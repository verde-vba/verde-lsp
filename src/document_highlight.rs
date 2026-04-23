use tower_lsp::lsp_types::*;

use crate::analysis::resolve::{
    find_all_word_occurrences, find_word_at_position, position_to_offset, text_range_to_lsp_range,
};
use crate::analysis::symbols::SymbolKind;
use crate::analysis::AnalysisHost;

pub fn document_highlight(
    host: &AnalysisHost,
    uri: &Url,
    position: Position,
) -> Option<Vec<DocumentHighlight>> {
    let word = host.with_source(uri, |_, source| find_word_at_position(source, position))??;

    let highlights = host.with_source(uri, |symbols, source| {
        // Check if the word refers to a proc-scoped symbol in the procedure at cursor.
        let proc_scope_range = (|| {
            let offset = position_to_offset(source, position)?;
            let containing_proc = symbols
                .proc_ranges
                .iter()
                .find(|(_, range)| {
                    offset >= range.start as usize && offset <= range.end as usize
                })?;
            let is_proc_scoped = symbols.symbols.iter().any(|sym| {
                sym.name.eq_ignore_ascii_case(&word)
                    && sym.proc_scope.as_deref() == Some(containing_proc.0.as_str())
                    && matches!(
                        sym.kind,
                        SymbolKind::Variable | SymbolKind::Parameter | SymbolKind::Constant
                    )
            });
            if is_proc_scoped {
                Some(containing_proc.1)
            } else {
                None
            }
        })();

        let occurrences = find_all_word_occurrences(source, &word);
        let filtered: Vec<_> = if let Some(proc_range) = proc_scope_range {
            occurrences
                .into_iter()
                .filter(|r| r.start >= proc_range.start && r.end <= proc_range.end)
                .collect()
        } else {
            occurrences
        };

        filtered
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

#[cfg(test)]
mod tests {
    use super::*;
    use crate::analysis::AnalysisHost;
    use crate::parser;

    #[test]
    fn proc_scoped_variable_highlights_only_within_procedure() {
        let host = AnalysisHost::new();
        let source = "Sub A()\n    Dim x As Long\n    x = 1\nEnd Sub\nSub B()\n    Dim x As Long\n    x = 2\nEnd Sub\n";
        let uri = Url::parse("file:///test.bas").unwrap();
        let parse_result = parser::parse(source);
        host.update(uri.clone(), source.to_string(), parse_result);

        // Cursor on `x` inside Sub A (line 1, col 8 => "Dim x")
        let highlights = document_highlight(&host, &uri, Position::new(1, 8));
        let highlights = highlights.expect("expected some highlights");
        // Should only find highlights within Sub A, not Sub B
        assert!(
            highlights.len() == 2,
            "expected 2 highlights for proc-scoped x in Sub A, got {}",
            highlights.len()
        );
        for h in &highlights {
            assert!(
                h.range.start.line < 4,
                "highlight should be within Sub A, but found at line {}",
                h.range.start.line
            );
        }
    }
}
