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
    // Determine the word at cursor and guard: only rename known declared symbols
    // (either in the current file or as a Public symbol in another file).
    let word = host.with_source(uri, |symbols, source| {
        let word = resolve::find_word_at_position(source, position)?;
        let found_locally = !resolve::find_symbol_by_name(symbols, &word).is_empty();
        let found_cross = !found_locally
            && host.find_public_symbol_in_other_files(uri, &word).is_some();
        if found_locally || found_cross { Some(word) } else { None }
    })??;

    // Collect edits across all registered files.
    let mut changes: HashMap<Url, Vec<TextEdit>> = HashMap::new();
    for (file_uri, source) in host.all_file_sources() {
        let edits: Vec<TextEdit> = resolve::find_all_word_occurrences(&source, &word)
            .into_iter()
            .map(|range| {
                TextEdit::new(
                    resolve::text_range_to_lsp_range(&source, range),
                    new_name.to_string(),
                )
            })
            .collect();
        if !edits.is_empty() {
            changes.insert(file_uri, edits);
        }
    }

    if changes.is_empty() {
        None
    } else {
        Some(WorkspaceEdit { changes: Some(changes), ..Default::default() })
    }
}
