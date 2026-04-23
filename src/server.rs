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

    async fn on_change(&self, uri: Url, text: String) {
        let parse_result = parser::parse(&text);
        self.analysis
            .update(uri.clone(), text.clone(), parse_result);
        self.documents.insert(uri.clone(), text);

        let diagnostics = self.analysis.diagnostics(&uri);
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
        semantic_tokens_provider: Some(
            SemanticTokensServerCapabilities::SemanticTokensOptions(SemanticTokensOptions {
                legend: crate::semantic_tokens::legend(),
                full: Some(SemanticTokensFullOptions::Bool(true)),
                range: None,
                ..Default::default()
            }),
        ),
        call_hierarchy_provider: Some(CallHierarchyServerCapability::Simple(true)),
        ..Default::default()
    }
}

#[tower_lsp::async_trait]
impl LanguageServer for VbaLanguageServer {
    async fn initialize(&self, params: InitializeParams) -> Result<InitializeResult> {
        // Capture workspace root for workbook-context.json discovery.
        let root = params
            .root_uri
            .or_else(|| params.workspace_folders?.into_iter().next().map(|f| f.uri));
        *self.root_uri.write().unwrap_or_else(|e| e.into_inner()) = root;

        Ok(InitializeResult {
            capabilities: server_capabilities(),
            ..Default::default()
        })
    }

    async fn initialized(&self, _: InitializedParams) {
        // Clone before awaiting to drop the RwLockReadGuard first.
        let root = self.root_uri.read().unwrap_or_else(|e| e.into_inner()).clone();
        if let Some(root) = root {
            if let Ok(base) = root.to_file_path() {
                self.analysis
                    .reload_workbook_context_from_path(&base.join("workbook-context.json"));
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
        for change in &params.changes {
            if change.uri.path().ends_with("workbook-context.json") {
                if let Ok(path) = change.uri.to_file_path() {
                    self.analysis.reload_workbook_context_from_path(&path);
                }
            }
        }
    }

    async fn shutdown(&self) -> Result<()> {
        Ok(())
    }

    async fn did_open(&self, params: DidOpenTextDocumentParams) {
        self.on_change(params.text_document.uri, params.text_document.text)
            .await;
    }

    async fn did_change(&self, params: DidChangeTextDocumentParams) {
        if let Some(change) = params.content_changes.into_iter().last() {
            self.on_change(params.text_document.uri, change.text).await;
        }
    }

    async fn did_close(&self, params: DidCloseTextDocumentParams) {
        self.documents.remove(&params.text_document.uri);
        self.analysis.remove(&params.text_document.uri);
    }

    async fn completion(&self, params: CompletionParams) -> Result<Option<CompletionResponse>> {
        let uri = &params.text_document_position.text_document.uri;
        let position = params.text_document_position.position;
        let items = crate::completion::complete(&self.analysis, uri, position);
        Ok(Some(CompletionResponse::Array(items)))
    }

    async fn signature_help(&self, params: SignatureHelpParams) -> Result<Option<SignatureHelp>> {
        let uri = &params.text_document_position_params.text_document.uri;
        let position = params.text_document_position_params.position;
        Ok(crate::signature_help::signature_help(
            &self.analysis,
            uri,
            position,
        ))
    }

    async fn hover(&self, params: HoverParams) -> Result<Option<Hover>> {
        let uri = &params.text_document_position_params.text_document.uri;
        let position = params.text_document_position_params.position;
        Ok(crate::hover::hover(&self.analysis, uri, position))
    }

    async fn goto_definition(
        &self,
        params: GotoDefinitionParams,
    ) -> Result<Option<GotoDefinitionResponse>> {
        let uri = &params.text_document_position_params.text_document.uri;
        let position = params.text_document_position_params.position;
        Ok(crate::definition::goto_definition(
            &self.analysis,
            uri,
            position,
        ))
    }

    async fn goto_type_definition(
        &self,
        params: GotoDefinitionParams,
    ) -> Result<Option<GotoDefinitionResponse>> {
        let uri = &params.text_document_position_params.text_document.uri;
        let position = params.text_document_position_params.position;
        Ok(crate::type_definition::goto_type_definition(
            &self.analysis,
            uri,
            position,
        ))
    }

    async fn references(&self, params: ReferenceParams) -> Result<Option<Vec<Location>>> {
        let uri = &params.text_document_position.text_document.uri;
        let position = params.text_document_position.position;
        let locs = crate::references::find_references(&self.analysis, uri, position);
        Ok(if locs.is_empty() { None } else { Some(locs) })
    }

    async fn code_action(&self, params: CodeActionParams) -> Result<Option<CodeActionResponse>> {
        let uri = &params.text_document.uri;
        let range = params.range;
        let diags = &params.context.diagnostics;
        let actions = crate::code_action::code_actions(&self.analysis, uri, range, diags);
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
        let ranges = crate::folding_range::folding_ranges(&self.analysis, uri);
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
        let syms = crate::workspace_symbol::workspace_symbols(&self.analysis, &params.query);
        Ok(if syms.is_empty() { None } else { Some(syms) })
    }

    async fn document_highlight(
        &self,
        params: DocumentHighlightParams,
    ) -> Result<Option<Vec<DocumentHighlight>>> {
        let uri = &params.text_document_position_params.text_document.uri;
        let position = params.text_document_position_params.position;
        Ok(crate::document_highlight::document_highlight(
            &self.analysis,
            uri,
            position,
        ))
    }

    async fn document_symbol(
        &self,
        params: DocumentSymbolParams,
    ) -> Result<Option<DocumentSymbolResponse>> {
        let uri = &params.text_document.uri;
        let syms = crate::document_symbol::document_symbols(&self.analysis, uri);
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
        Ok(crate::rename::rename(
            &self.analysis,
            uri,
            position,
            new_name,
        ))
    }

    async fn formatting(&self, params: DocumentFormattingParams) -> Result<Option<Vec<TextEdit>>> {
        let uri = &params.text_document.uri;
        let src = match self.documents.get(uri) {
            Some(s) => s.clone(),
            None => return Ok(None),
        };
        let formatted = crate::formatting::apply_formatting(&src);
        if formatted == src {
            return Ok(None);
        }
        let end = document_end_position(&src);
        Ok(Some(vec![TextEdit::new(
            Range::new(Position::new(0, 0), end),
            formatted,
        )]))
    }

    async fn inlay_hint(&self, params: InlayHintParams) -> Result<Option<Vec<InlayHint>>> {
        let uri = &params.text_document.uri;
        let hints = self.analysis.inlay_hints(uri, Some(params.range));
        Ok(Some(hints))
    }

    async fn semantic_tokens_full(
        &self,
        params: SemanticTokensParams,
    ) -> Result<Option<SemanticTokensResult>> {
        let uri = &params.text_document.uri;
        Ok(crate::semantic_tokens::semantic_tokens(&self.analysis, uri))
    }

    async fn prepare_call_hierarchy(
        &self,
        params: CallHierarchyPrepareParams,
    ) -> Result<Option<Vec<CallHierarchyItem>>> {
        let uri = &params.text_document_position_params.text_document.uri;
        let position = params.text_document_position_params.position;
        Ok(self.analysis.prepare_call_hierarchy(uri, position))
    }

    async fn incoming_calls(
        &self,
        params: CallHierarchyIncomingCallsParams,
    ) -> Result<Option<Vec<CallHierarchyIncomingCall>>> {
        let item = &params.item;
        let calls = self.analysis.incoming_calls(item);
        Ok(if calls.is_empty() { None } else { Some(calls) })
    }

    async fn outgoing_calls(
        &self,
        params: CallHierarchyOutgoingCallsParams,
    ) -> Result<Option<Vec<CallHierarchyOutgoingCall>>> {
        let item = &params.item;
        let calls = self.analysis.outgoing_calls(item);
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
