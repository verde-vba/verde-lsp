use tower_lsp::lsp_types::*;

use crate::analysis::resolve;
use crate::analysis::AnalysisHost;

pub fn goto_definition(
    host: &AnalysisHost,
    uri: &Url,
    position: Position,
) -> Option<GotoDefinitionResponse> {
    host.with_source(uri, |symbols, source| {
        let word = resolve::find_word_at_position(source, position)?;
        let matches = resolve::find_symbol_by_name(symbols, &word);
        let sym = matches.first()?;

        let range = resolve::text_range_to_lsp_range(source, sym.span);
        Some(GotoDefinitionResponse::Scalar(Location::new(
            uri.clone(),
            range,
        )))
    })?
}
