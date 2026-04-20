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
    Server::new(stdin, stdout, socket).serve(service).await;
}
