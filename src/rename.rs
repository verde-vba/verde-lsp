use std::collections::HashMap;
use tower_lsp::lsp_types::*;

use crate::analysis::resolve;
use crate::analysis::AnalysisHost;
use crate::parser::ast::Visibility;

pub fn rename(
    host: &AnalysisHost,
    uri: &Url,
    position: Position,
    new_name: &str,
) -> Option<WorkspaceEdit> {
    // Determine the word at cursor, guard for known symbols, and decide scope.
    // cross_file is true only when the symbol is Public and module-level —
    // Private symbols (including local variables) must not propagate across files.
    let (word, cross_file) = host.with_source(uri, |symbols, source| {
        let word = resolve::find_word_at_position(source, position)?;
        let local_syms = resolve::find_symbol_by_name(symbols, &word);
        if !local_syms.is_empty() {
            let is_public_module_level = local_syms
                .iter()
                .any(|s| s.visibility == Visibility::Public && s.proc_scope.is_none());
            Some((word, is_public_module_level))
        } else {
            let found_cross = host.find_public_symbol_in_other_files(uri, &word).is_some();
            if found_cross {
                Some((word, true))
            } else {
                None
            }
        }
    })??;

    let files: Vec<(Url, String)> = if cross_file {
        host.all_file_sources()
    } else {
        host.with_source(uri, |_, source| vec![(uri.clone(), source.to_string())])
            .unwrap_or_default()
    };

    // Collect edits across the determined file set.
    let mut changes: HashMap<Url, Vec<TextEdit>> = HashMap::new();
    for (file_uri, source) in files {
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
        Some(WorkspaceEdit {
            changes: Some(changes),
            ..Default::default()
        })
    }
}
