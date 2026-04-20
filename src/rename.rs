use std::collections::HashMap;
use tower_lsp::lsp_types::*;

use crate::analysis::resolve;
use crate::analysis::AnalysisHost;

pub fn rename(
    host: &AnalysisHost,
    uri: &Url,
    position: Position,
    new_name: &str,
) -> Option<WorkspaceEdit> {
    host.with_source(uri, |symbols, source| {
        let word = resolve::find_word_at_position(source, position)?;

        // Guard: only rename identifiers that are known declared symbols.
        if resolve::find_symbol_by_name(symbols, &word).is_empty() {
            return None;
        }

        let edits: Vec<TextEdit> = resolve::find_all_word_occurrences(source, &word)
            .into_iter()
            .map(|range| TextEdit::new(resolve::text_range_to_lsp_range(source, range), new_name.to_string()))
            .collect();

        let mut changes = HashMap::new();
        changes.insert(uri.clone(), edits);

        Some(WorkspaceEdit {
            changes: Some(changes),
            ..Default::default()
        })
    })?
}
