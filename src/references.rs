use tower_lsp::lsp_types::*;

use crate::analysis::resolve::{find_all_word_occurrences, find_word_at_position, text_range_to_lsp_range};
use crate::analysis::AnalysisHost;

pub fn find_references(host: &AnalysisHost, uri: &Url, position: Position) -> Vec<Location> {
    host.with_source(uri, |_, source| {
        let word = match find_word_at_position(source, position) {
            Some(w) => w,
            None => return Vec::new(),
        };
        find_all_word_occurrences(source, &word)
            .into_iter()
            .map(|range| Location {
                uri: uri.clone(),
                range: text_range_to_lsp_range(source, range),
            })
            .collect()
    })
    .unwrap_or_default()
}
