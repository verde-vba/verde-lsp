use tower_lsp::{LspService, Server};
use verde_lsp::server;

#[tokio::main]
async fn main() {
    // [DEBUG LOGGING] デフォルトを debug に設定して詳細ログを出力
    env_logger::Builder::from_default_env()
        .filter_level(log::LevelFilter::Debug)
        .target(env_logger::Target::Stderr)
        .init();

    log::info!("verde-lsp starting — pid={}", std::process::id());
    log::info!("verde-lsp cwd={:?}", std::env::current_dir().ok());
    log::info!("verde-lsp args={:?}", std::env::args().collect::<Vec<_>>());

    let stdin = tokio::io::stdin();
    let stdout = tokio::io::stdout();

    log::info!("verde-lsp creating LspService…");
    let (service, socket) = LspService::new(server::VbaLanguageServer::new);

    log::info!("verde-lsp starting server on stdin/stdout…");
    // Spawn serve() on a task so we can catch panics caused by broken
    // stdio pipes (e.g. when the Tauri frontend terminates).
    let handle = tokio::spawn(async move {
        Server::new(stdin, stdout, socket).serve(service).await;
    });

    if let Err(e) = handle.await {
        // The client has already disconnected — exit cleanly.
        log::warn!("LSP server exited: {e}");
    }
    log::info!("verde-lsp process exiting");
}
