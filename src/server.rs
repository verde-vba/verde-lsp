use dashmap::DashMap;
use tower_lsp::jsonrpc::Result;
use tower_lsp::lsp_types::*;
use tower_lsp::{Client, LanguageServer};

use crate::analysis::AnalysisHost;
use crate::parser;

pub struct VbaLanguageServer {
    client: Client,
    documents: DashMap<Url, String>,
    analysis: AnalysisHost,
    root_uri: std::sync::RwLock<Option<Url>>,
}

impl VbaLanguageServer {
    pub fn new(client: Client) -> Self {
        Self {
            client,
            documents: DashMap::new(),
            analysis: AnalysisHost::new(),
            root_uri: std::sync::RwLock::new(None),
        }
    }

    /// Scan the workspace root for `.bas`, `.cls`, `.frm` files and load
    /// each into `AnalysisHost` from disk. Files already opened by the
    /// client (present in `documents`) are skipped — the editor version
    /// takes precedence.
    fn load_workspace_files(&self, root: &std::path::Path) {
        const VBA_EXTENSIONS: &[&str] = &["bas", "cls", "frm"];

        let entries = match std::fs::read_dir(root) {
            Ok(entries) => entries,
            Err(e) => {
                log::warn!("[load_workspace_files] failed to read dir {:?}: {}", root, e);
                return;
            }
        };

        let mut count = 0u32;
        for entry in entries.filter_map(|e| e.ok()) {
            let path = entry.path();
            if !path.is_file() {
                continue;
            }
            let ext = path
                .extension()
                .and_then(|e| e.to_str())
                .unwrap_or("");
            if !VBA_EXTENSIONS.iter().any(|&v| v.eq_ignore_ascii_case(ext)) {
                continue;
            }
            if let Ok(uri) = Url::from_file_path(&path) {
                if self.documents.contains_key(&uri) {
                    continue; // client-managed — skip
                }
                if self.load_file_from_disk(&uri, &path) {
                    count += 1;
                }
            }
        }
        log::info!("[load_workspace_files] loaded {} VBA files from {:?}", count, root);
    }

    /// Read a single file from disk, parse it, and store the result in
    /// `AnalysisHost`. Does NOT insert into `documents` (reserved for
    /// client-opened files) and does NOT publish diagnostics.
    ///
    /// VBA modules exported from Japanese Excel are often encoded in
    /// Shift-JIS (CP932).  We try UTF-8 first and fall back to Shift-JIS
    /// so cross-module features work regardless of the source encoding.
    fn load_file_from_disk(&self, uri: &Url, path: &std::path::Path) -> bool {
        let bytes = match std::fs::read(path) {
            Ok(b) => b,
            Err(e) => {
                log::warn!("[load_file_from_disk] failed to read {:?}: {}", path, e);
                return false;
            }
        };

        let text = match String::from_utf8(bytes.clone()) {
            Ok(s) => s,
            Err(_) => {
                // Fallback: try Shift-JIS (common for Japanese VBA exports).
                let (decoded, _encoding, had_errors) =
                    encoding_rs::SHIFT_JIS.decode(&bytes);
                if had_errors {
                    log::warn!(
                        "[load_file_from_disk] {:?}: not valid UTF-8 or Shift-JIS, using lossy decode",
                        path
                    );
                }
                log::debug!("[load_file_from_disk] {:?}: decoded as Shift-JIS", path);
                decoded.into_owned()
            }
        };

        let parse_result = parser::parse(&text);
        self.analysis.update(uri.clone(), text, parse_result);
        log::debug!("[load_file_from_disk] loaded {:?}", path);
        true
    }

    async fn on_change(&self, uri: Url, text: String) {
        log::debug!("[on_change] uri={} text_len={}", uri, text.len());
        let parse_result = parser::parse(&text);
        log::debug!("[on_change] parsed — errors={} tokens={}", parse_result.errors.len(), parse_result.tokens.len());
        self.analysis
            .update(uri.clone(), text.clone(), parse_result);
        self.documents.insert(uri.clone(), text);

        let diagnostics = self.analysis.diagnostics(&uri);
        log::debug!("[on_change] publishing {} diagnostics for {}", diagnostics.len(), uri);
        self.client
            .publish_diagnostics(uri, diagnostics, None)
            .await;
    }
}

fn document_end_position(src: &str) -> Position {
    let mut line = 0u32;
    let mut last_line_start = 0usize;
    for (i, ch) in src.char_indices() {
        if ch == '\n' {
            line += 1;
            last_line_start = i + 1;
        }
    }
    let character = src[last_line_start..].encode_utf16().count() as u32;
    Position::new(line, character)
}

fn server_capabilities() -> ServerCapabilities {
    ServerCapabilities {
        position_encoding: Some(PositionEncodingKind::UTF16),
        text_document_sync: Some(TextDocumentSyncCapability::Kind(TextDocumentSyncKind::FULL)),
        completion_provider: Some(CompletionOptions {
            trigger_characters: Some(vec![".".to_string()]),
            ..Default::default()
        }),
        hover_provider: Some(HoverProviderCapability::Simple(true)),
        signature_help_provider: Some(SignatureHelpOptions {
            trigger_characters: Some(vec!["(".to_string(), ",".to_string()]),
            ..Default::default()
        }),
        definition_provider: Some(OneOf::Left(true)),
        type_definition_provider: Some(TypeDefinitionProviderCapability::Simple(true)),
        rename_provider: Some(OneOf::Left(true)),
        references_provider: Some(OneOf::Left(true)),
        code_action_provider: Some(CodeActionProviderCapability::Simple(true)),
        folding_range_provider: Some(FoldingRangeProviderCapability::Simple(true)),
        workspace_symbol_provider: Some(OneOf::Left(true)),
        document_highlight_provider: Some(OneOf::Left(true)),
        document_symbol_provider: Some(OneOf::Left(true)),
        document_formatting_provider: Some(OneOf::Left(true)),
        inlay_hint_provider: Some(OneOf::Left(true)),
        semantic_tokens_provider: Some(SemanticTokensServerCapabilities::SemanticTokensOptions(
            SemanticTokensOptions {
                legend: crate::semantic_tokens::legend(),
                full: Some(SemanticTokensFullOptions::Bool(true)),
                range: None,
                ..Default::default()
            },
        )),
        call_hierarchy_provider: Some(CallHierarchyServerCapability::Simple(true)),
        ..Default::default()
    }
}

#[tower_lsp::async_trait]
impl LanguageServer for VbaLanguageServer {
    async fn initialize(&self, params: InitializeParams) -> Result<InitializeResult> {
        log::info!("[initialize] received — root_uri={:?} client={:?}",
            params.root_uri, params.client_info);
        // Capture workspace root for workbook-context.json discovery.
        let root = params
            .root_uri
            .or_else(|| params.workspace_folders?.into_iter().next().map(|f| f.uri));
        log::info!("[initialize] resolved root={:?}", root);
        *self.root_uri.write().unwrap_or_else(|e| e.into_inner()) = root;

        let caps = server_capabilities();
        log::info!("[initialize] returning capabilities — completion={} hover={} definition={} rename={} formatting={} semanticTokens={}",
            caps.completion_provider.is_some(),
            caps.hover_provider.is_some(),
            caps.definition_provider.is_some(),
            caps.rename_provider.is_some(),
            caps.document_formatting_provider.is_some(),
            caps.semantic_tokens_provider.is_some(),
        );
        Ok(InitializeResult {
            capabilities: caps,
            ..Default::default()
        })
    }

    async fn initialized(&self, _: InitializedParams) {
        log::info!("[initialized] notification received");
        // Clone before awaiting to drop the RwLockReadGuard first.
        let root = self
            .root_uri
            .read()
            .unwrap_or_else(|e| e.into_inner())
            .clone();
        log::info!("[initialized] root_uri={:?}", root);
        if let Some(root) = root {
            if let Ok(base) = root.to_file_path() {
                self.analysis
                    .reload_workbook_context_from_path(&base.join("workbook-context.json"));
                // Scan workspace for VBA files and load them into analysis
                // so cross-module features work even before the client opens
                // each file in an editor tab.
                self.load_workspace_files(&base);
            }
            // Register a file watcher so clients notify us when workbook-context.json changes.
            let _ = self
                .client
                .register_capability(vec![Registration {
                    id: "workbook-context-watcher".to_string(),
                    method: "workspace/didChangeWatchedFiles".to_string(),
                    register_options: serde_json::to_value(
                        DidChangeWatchedFilesRegistrationOptions {
                            watchers: vec![FileSystemWatcher {
                                glob_pattern: GlobPattern::String(
                                    "**/workbook-context.json".to_string(),
                                ),
                                kind: None,
                            }],
                        },
                    )
                    .ok(),
                }])
                .await;
        }
        self.client
            .log_message(MessageType::INFO, "verde-lsp initialized")
            .await;
    }

    async fn did_change_watched_files(&self, params: DidChangeWatchedFilesParams) {
        log::info!("[didChangeWatchedFiles] {} changes", params.changes.len());
        for change in &params.changes {
            log::info!("[didChangeWatchedFiles] uri={} type={:?}", change.uri, change.typ);
            if change.uri.path().ends_with("workbook-context.json") {
                if let Ok(path) = change.uri.to_file_path() {
                    let ok = self.analysis.reload_workbook_context_from_path(&path);
                    log::info!("[didChangeWatchedFiles] reloaded workbook-context.json: success={}", ok);
                }
            }
        }
    }

    async fn shutdown(&self) -> Result<()> {
        log::info!("[shutdown] received");
        Ok(())
    }

    async fn did_open(&self, params: DidOpenTextDocumentParams) {
        log::info!("[didOpen] uri={} languageId={} version={} text_len={}",
            params.text_document.uri,
            params.text_document.language_id,
            params.text_document.version,
            params.text_document.text.len(),
        );
        self.on_change(params.text_document.uri, params.text_document.text)
            .await;
    }

    async fn did_change(&self, params: DidChangeTextDocumentParams) {
        log::info!("[didChange] uri={} version={} changes={}",
            params.text_document.uri,
            params.text_document.version,
            params.content_changes.len(),
        );
        if let Some(change) = params.content_changes.into_iter().last() {
            self.on_change(params.text_document.uri, change.text).await;
        }
    }

    async fn did_close(&self, params: DidCloseTextDocumentParams) {
        let uri = params.text_document.uri;
        log::info!("[didClose] uri={}", uri);
        self.documents.remove(&uri);

        // Fall back to the on-disk version so cross-module features
        // (completion, references, go-to-definition) keep working for
        // symbols defined in this file.
        if let Ok(path) = uri.to_file_path() {
            if self.load_file_from_disk(&uri, &path) {
                log::info!("[didClose] fell back to disk version for {}", uri);
                return;
            }
        }
        // Disk read failed (file deleted?) — remove from analysis.
        log::info!("[didClose] disk fallback failed, removing {}", uri);
        self.analysis.remove(&uri);
    }

    async fn completion(&self, params: CompletionParams) -> Result<Option<CompletionResponse>> {
        let uri = &params.text_document_position.text_document.uri;
        let position = params.text_document_position.position;
        log::debug!("[completion] uri={} pos={}:{}", uri, position.line, position.character);
        let items = crate::completion::complete(&self.analysis, uri, position);
        log::debug!("[completion] returning {} items", items.len());
        Ok(Some(CompletionResponse::Array(items)))
    }

    async fn signature_help(&self, params: SignatureHelpParams) -> Result<Option<SignatureHelp>> {
        let uri = &params.text_document_position_params.text_document.uri;
        let position = params.text_document_position_params.position;
        log::debug!("[signatureHelp] uri={} pos={}:{}", uri, position.line, position.character);
        let result = crate::signature_help::signature_help(
            &self.analysis,
            uri,
            position,
        );
        log::debug!("[signatureHelp] result={}", result.is_some());
        Ok(result)
    }

    async fn hover(&self, params: HoverParams) -> Result<Option<Hover>> {
        let uri = &params.text_document_position_params.text_document.uri;
        let position = params.text_document_position_params.position;
        log::debug!("[hover] uri={} pos={}:{}", uri, position.line, position.character);
        let result = crate::hover::hover(&self.analysis, uri, position);
        log::debug!("[hover] result={}", result.is_some());
        Ok(result)
    }

    async fn goto_definition(
        &self,
        params: GotoDefinitionParams,
    ) -> Result<Option<GotoDefinitionResponse>> {
        let uri = &params.text_document_position_params.text_document.uri;
        let position = params.text_document_position_params.position;
        log::debug!("[gotoDefinition] uri={} pos={}:{}", uri, position.line, position.character);
        let result = crate::definition::goto_definition(
            &self.analysis,
            uri,
            position,
        );
        log::debug!("[gotoDefinition] result={}", result.is_some());
        Ok(result)
    }

    async fn goto_type_definition(
        &self,
        params: GotoDefinitionParams,
    ) -> Result<Option<GotoDefinitionResponse>> {
        let uri = &params.text_document_position_params.text_document.uri;
        let position = params.text_document_position_params.position;
        log::debug!("[gotoTypeDefinition] uri={} pos={}:{}", uri, position.line, position.character);
        let result = crate::type_definition::goto_type_definition(
            &self.analysis,
            uri,
            position,
        );
        log::debug!("[gotoTypeDefinition] result={}", result.is_some());
        Ok(result)
    }

    async fn references(&self, params: ReferenceParams) -> Result<Option<Vec<Location>>> {
        let uri = &params.text_document_position.text_document.uri;
        let position = params.text_document_position.position;
        log::debug!("[references] uri={} pos={}:{}", uri, position.line, position.character);
        let locs = crate::references::find_references(&self.analysis, uri, position);
        log::debug!("[references] found {} locations", locs.len());
        Ok(if locs.is_empty() { None } else { Some(locs) })
    }

    async fn code_action(&self, params: CodeActionParams) -> Result<Option<CodeActionResponse>> {
        let uri = &params.text_document.uri;
        let range = params.range;
        let diags = &params.context.diagnostics;
        log::debug!("[codeAction] uri={} range={}:{}-{}:{} diags={}",
            uri, range.start.line, range.start.character, range.end.line, range.end.character, diags.len());
        let actions = crate::code_action::code_actions(&self.analysis, uri, range, diags);
        log::debug!("[codeAction] returning {} actions", actions.len());
        if actions.is_empty() {
            return Ok(None);
        }
        Ok(Some(
            actions
                .into_iter()
                .map(CodeActionOrCommand::CodeAction)
                .collect(),
        ))
    }

    async fn folding_range(&self, params: FoldingRangeParams) -> Result<Option<Vec<FoldingRange>>> {
        let uri = &params.text_document.uri;
        log::debug!("[foldingRange] uri={}", uri);
        let ranges = crate::folding_range::folding_ranges(&self.analysis, uri);
        log::debug!("[foldingRange] returning {} ranges", ranges.len());
        Ok(if ranges.is_empty() {
            None
        } else {
            Some(ranges)
        })
    }

    async fn symbol(
        &self,
        params: WorkspaceSymbolParams,
    ) -> Result<Option<Vec<SymbolInformation>>> {
        log::debug!("[workspaceSymbol] query={:?}", params.query);
        let syms = crate::workspace_symbol::workspace_symbols(&self.analysis, &params.query);
        log::debug!("[workspaceSymbol] returning {} symbols", syms.len());
        Ok(if syms.is_empty() { None } else { Some(syms) })
    }

    async fn document_highlight(
        &self,
        params: DocumentHighlightParams,
    ) -> Result<Option<Vec<DocumentHighlight>>> {
        let uri = &params.text_document_position_params.text_document.uri;
        let position = params.text_document_position_params.position;
        log::debug!("[documentHighlight] uri={} pos={}:{}", uri, position.line, position.character);
        let result = crate::document_highlight::document_highlight(
            &self.analysis,
            uri,
            position,
        );
        log::debug!("[documentHighlight] result={:?}", result.as_ref().map(|v| v.len()));
        Ok(result)
    }

    async fn document_symbol(
        &self,
        params: DocumentSymbolParams,
    ) -> Result<Option<DocumentSymbolResponse>> {
        let uri = &params.text_document.uri;
        log::debug!("[documentSymbol] uri={}", uri);
        let syms = crate::document_symbol::document_symbols(&self.analysis, uri);
        log::debug!("[documentSymbol] returning {} symbols", syms.len());
        if syms.is_empty() {
            Ok(None)
        } else {
            Ok(Some(DocumentSymbolResponse::Nested(syms)))
        }
    }

    async fn rename(&self, params: RenameParams) -> Result<Option<WorkspaceEdit>> {
        let uri = &params.text_document_position.text_document.uri;
        let position = params.text_document_position.position;
        let new_name = &params.new_name;
        log::debug!("[rename] uri={} pos={}:{} newName={:?}", uri, position.line, position.character, new_name);
        let result = crate::rename::rename(
            &self.analysis,
            uri,
            position,
            new_name,
        );
        log::debug!("[rename] result={}", result.is_some());
        Ok(result)
    }

    async fn formatting(&self, params: DocumentFormattingParams) -> Result<Option<Vec<TextEdit>>> {
        let uri = &params.text_document.uri;
        log::debug!("[formatting] uri={}", uri);
        let src = match self.documents.get(uri) {
            Some(s) => s.clone(),
            None => {
                log::debug!("[formatting] document not found");
                return Ok(None);
            }
        };
        let formatted = crate::formatting::apply_formatting(&src);
        if formatted == src {
            log::debug!("[formatting] no changes");
            return Ok(None);
        }
        log::debug!("[formatting] returning edit (src_len={} → formatted_len={})", src.len(), formatted.len());
        let end = document_end_position(&src);
        Ok(Some(vec![TextEdit::new(
            Range::new(Position::new(0, 0), end),
            formatted,
        )]))
    }

    async fn inlay_hint(&self, params: InlayHintParams) -> Result<Option<Vec<InlayHint>>> {
        let uri = &params.text_document.uri;
        log::debug!("[inlayHint] uri={} range={}:{}-{}:{}",
            uri, params.range.start.line, params.range.start.character,
            params.range.end.line, params.range.end.character);
        let hints = self.analysis.inlay_hints(uri, Some(params.range));
        log::debug!("[inlayHint] returning {} hints", hints.len());
        Ok(Some(hints))
    }

    async fn semantic_tokens_full(
        &self,
        params: SemanticTokensParams,
    ) -> Result<Option<SemanticTokensResult>> {
        let uri = &params.text_document.uri;
        log::debug!("[semanticTokensFull] uri={}", uri);
        let result = crate::semantic_tokens::semantic_tokens(&self.analysis, uri);
        log::debug!("[semanticTokensFull] result={}", result.is_some());
        Ok(result)
    }

    async fn prepare_call_hierarchy(
        &self,
        params: CallHierarchyPrepareParams,
    ) -> Result<Option<Vec<CallHierarchyItem>>> {
        let uri = &params.text_document_position_params.text_document.uri;
        let position = params.text_document_position_params.position;
        log::debug!("[prepareCallHierarchy] uri={} pos={}:{}", uri, position.line, position.character);
        let result = self.analysis.prepare_call_hierarchy(uri, position);
        log::debug!("[prepareCallHierarchy] result={:?}", result.as_ref().map(|v| v.len()));
        Ok(result)
    }

    async fn incoming_calls(
        &self,
        params: CallHierarchyIncomingCallsParams,
    ) -> Result<Option<Vec<CallHierarchyIncomingCall>>> {
        let item = &params.item;
        log::debug!("[incomingCalls] item={:?}", item.name);
        let calls = self.analysis.incoming_calls(item);
        log::debug!("[incomingCalls] returning {} calls", calls.len());
        Ok(if calls.is_empty() { None } else { Some(calls) })
    }

    async fn outgoing_calls(
        &self,
        params: CallHierarchyOutgoingCallsParams,
    ) -> Result<Option<Vec<CallHierarchyOutgoingCall>>> {
        let item = &params.item;
        log::debug!("[outgoingCalls] item={:?}", item.name);
        let calls = self.analysis.outgoing_calls(item);
        log::debug!("[outgoingCalls] returning {} calls", calls.len());
        Ok(if calls.is_empty() { None } else { Some(calls) })
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn server_capabilities_declares_utf16_position_encoding() {
        // LSP 3.17: server should explicitly negotiate positionEncoding.
        // We declare UTF-16 to match the resolve.rs implementation
        // (PBI-31 Sprint N+33).
        let caps = server_capabilities();
        assert_eq!(caps.position_encoding, Some(PositionEncodingKind::UTF16));
    }

    #[test]
    fn server_capabilities_declares_inlay_hint_provider() {
        let caps = server_capabilities();
        assert!(
            caps.inlay_hint_provider.is_some(),
            "inlayHintProvider must be declared in server capabilities"
        );
    }

    #[test]
    fn server_capabilities_declares_semantic_tokens_provider() {
        let caps = server_capabilities();
        assert!(
            caps.semantic_tokens_provider.is_some(),
            "semanticTokensProvider must be declared in server capabilities"
        );
    }

    #[test]
    fn server_capabilities_declares_call_hierarchy_provider() {
        let caps = server_capabilities();
        assert!(
            caps.call_hierarchy_provider.is_some(),
            "callHierarchyProvider must be declared in server capabilities"
        );
    }
}
