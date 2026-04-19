use tower_lsp::lsp_types::*;

use crate::analysis::AnalysisHost;
use crate::analysis::resolve;

pub fn goto_definition(
    host: &AnalysisHost,
    uri: &Url,
    position: Position,
) -> Option<GotoDefinitionResponse> {
    host.with_symbols(uri, |symbols| {
        let word = resolve::find_word_at_position("", position)?;
        let matches = resolve::find_symbol_by_name(symbols, &word);
        let sym = matches.first()?;

        let range = resolve::text_range_to_lsp_range("", sym.span);
        Some(GotoDefinitionResponse::Scalar(Location::new(
            uri.clone(),
            range,
        )))
    })?
}
