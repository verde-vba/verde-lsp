use tower_lsp::lsp_types::*;

use crate::analysis::resolve;
use crate::analysis::AnalysisHost;

pub fn goto_definition(
    host: &AnalysisHost,
    uri: &Url,
    position: Position,
) -> Option<GotoDefinitionResponse> {
    // Try current file first
    let result = host.with_source(uri, |symbols, source| {
        let word = resolve::find_word_at_position(source, position)?;
        let matches = resolve::find_symbol_by_name(symbols, &word);
        let sym = matches.first()?;

        let range = resolve::text_range_to_lsp_range(source, sym.span);
        Some(GotoDefinitionResponse::Scalar(Location::new(
            uri.clone(),
            range,
        )))
    });
    if let Some(Some(r)) = result {
        return Some(r);
    }

    // Fallback: cross-module public symbols
    let word = host.with_source(uri, |_, source| resolve::find_word_at_position(source, position))??;
    let (other_uri, sym) = host.find_public_symbol_in_other_files(uri, &word)?;
    let range = host.with_source(&other_uri, |_, source| {
        resolve::text_range_to_lsp_range(source, sym.span)
    })?;
    Some(GotoDefinitionResponse::Scalar(Location::new(other_uri, range)))
}
