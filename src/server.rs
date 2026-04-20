use dashmap::DashMap;
use tower_lsp::jsonrpc::Result;
use tower_lsp::lsp_types::*;
use tower_lsp::{Client, LanguageServer};

use crate::analysis::{AnalysisHost, WorkbookContext};
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

#[tower_lsp::async_trait]
impl LanguageServer for VbaLanguageServer {
    async fn initialize(&self, params: InitializeParams) -> Result<InitializeResult> {
        // Capture workspace root for workbook-context.json discovery.
        let root = params
            .root_uri
            .or_else(|| params.workspace_folders?.into_iter().next().map(|f| f.uri));
        *self.root_uri.write().unwrap() = root;

        Ok(InitializeResult {
            capabilities: ServerCapabilities {
                text_document_sync: Some(TextDocumentSyncCapability::Kind(
                    TextDocumentSyncKind::FULL,
                )),
                completion_provider: Some(CompletionOptions {
                    trigger_characters: Some(vec![".".to_string()]),
                    ..Default::default()
                }),
                hover_provider: Some(HoverProviderCapability::Simple(true)),
                definition_provider: Some(OneOf::Left(true)),
                rename_provider: Some(OneOf::Left(true)),
                ..Default::default()
            },
            ..Default::default()
        })
    }

    async fn initialized(&self, _: InitializedParams) {
        // Clone before .await to avoid holding RwLockReadGuard across await point.
        let root = self.root_uri.read().unwrap().clone();
        if let Some(root) = root {
            if let Ok(path) = root.to_file_path() {
                let json_path = path.join("workbook-context.json");
                if let Ok(content) = tokio::fs::read_to_string(&json_path).await {
                    if let Ok(ctx) = serde_json::from_str::<WorkbookContext>(&content) {
                        self.analysis.set_workbook_context(ctx);
                    }
                }
            }
        }
        self.client
            .log_message(MessageType::INFO, "verde-lsp initialized")
            .await;
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
}
