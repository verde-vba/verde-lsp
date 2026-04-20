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
        let matches = resolve::find_symbol_by_name(symbols, &word);

        if matches.is_empty() {
            return None;
        }

        let edits: Vec<TextEdit> = matches
            .iter()
            .map(|sym| {
                let range = resolve::text_range_to_lsp_range(source, sym.span);
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
