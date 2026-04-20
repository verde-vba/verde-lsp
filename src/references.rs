use tower_lsp::lsp_types::*;

use crate::analysis::resolve::{
    find_all_word_occurrences, find_word_at_position, text_range_to_lsp_range,
};
use crate::analysis::AnalysisHost;

pub fn find_references(host: &AnalysisHost, uri: &Url, position: Position) -> Vec<Location> {
    let word = match host.with_source(uri, |_, source| find_word_at_position(source, position)) {
        Some(Some(w)) => w,
        _ => return Vec::new(),
    };

    let mut locations = Vec::new();
    for (file_uri, source) in host.all_file_sources() {
        for range in find_all_word_occurrences(&source, &word) {
            locations.push(Location {
                uri: file_uri.clone(),
                range: text_range_to_lsp_range(&source, range),
            });
        }
    }
    locations
}
