mod parser;
mod analysis;
mod completion;
mod hover;
mod definition;
mod rename;
mod excel_model;
mod vba_builtins;
mod server;

use tower_lsp::{LspService, Server};

#[tokio::main]
async fn main() {
    env_logger::init();

    let stdin = tokio::io::stdin();
    let stdout = tokio::io::stdout();

    let (service, socket) = LspService::new(server::VbaLanguageServer::new);
    Server::new(stdin, stdout, socket).serve(service).await;
}
