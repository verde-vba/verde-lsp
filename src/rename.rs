use std::collections::HashMap;
use tower_lsp::lsp_types::*;

use crate::analysis::AnalysisHost;
use crate::analysis::resolve;

pub fn rename(
    host: &AnalysisHost,
    uri: &Url,
    position: Position,
    new_name: &str,
) -> Option<WorkspaceEdit> {
    host.with_symbols(uri, |symbols| {
        let word = resolve::find_word_at_position("", position)?;
        let matches = resolve::find_symbol_by_name(symbols, &word);

        if matches.is_empty() {
            return None;
        }

        let edits: Vec<TextEdit> = matches
            .iter()
            .map(|sym| {
                let range = resolve::text_range_to_lsp_range("", sym.span);
                TextEdit::new(range, new_name.to_string())
            })
            .collect();

        let mut changes = HashMap::new();
        changes.insert(uri.clone(), edits);

        Some(WorkspaceEdit {
            changes: Some(changes),
            ..Default::default()
        })
    })?
}
