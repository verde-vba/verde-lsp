use tower_lsp::{LspService, Server};
use verde_lsp::server;

#[tokio::main]
async fn main() {
    env_logger::Builder::from_default_env()
        .target(env_logger::Target::Stderr)
        .init();

    let stdin = tokio::io::stdin();
    let stdout = tokio::io::stdout();

    let (service, socket) = LspService::new(server::VbaLanguageServer::new);

    // Spawn serve() on a task so we can catch panics caused by broken
    // stdio pipes (e.g. when the Tauri frontend terminates).
    let handle = tokio::spawn(async move {
        Server::new(stdin, stdout, socket).serve(service).await;
    });

    if let Err(e) = handle.await {
        // The client has already disconnected — exit cleanly.
        log::warn!("LSP server exited: {e}");
    }
}
