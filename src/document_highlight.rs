use tower_lsp::lsp_types::*;

use crate::analysis::resolve::{
    find_all_word_occurrences, find_word_at_position, text_range_to_lsp_range,
};
use crate::analysis::AnalysisHost;

pub fn document_highlight(
    host: &AnalysisHost,
    uri: &Url,
    position: Position,
) -> Option<Vec<DocumentHighlight>> {
    let word = host.with_source(uri, |_, source| find_word_at_position(source, position))??;

    let highlights = host.with_source(uri, |_, source| {
        find_all_word_occurrences(source, &word)
            .into_iter()
            .map(|range| DocumentHighlight {
                range: text_range_to_lsp_range(source, range),
                kind: Some(DocumentHighlightKind::TEXT),
            })
            .collect::<Vec<_>>()
    })?;

    if highlights.is_empty() {
        None
    } else {
        Some(highlights)
    }
}
