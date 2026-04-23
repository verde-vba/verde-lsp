use tower_lsp::lsp_types::*;

use crate::analysis::resolve::{
    find_all_word_occurrences, find_word_at_position, position_to_offset, text_range_to_lsp_range,
};
use crate::analysis::symbols::SymbolKind;
use crate::analysis::AnalysisHost;

pub fn find_references(host: &AnalysisHost, uri: &Url, position: Position) -> Vec<Location> {
    let word = match host.with_source(uri, |_, source| find_word_at_position(source, position)) {
        Some(Some(w)) => w,
        _ => return Vec::new(),
    };

    // Check if the word refers to a proc-scoped symbol (Variable, Parameter, or Constant
    // with proc_scope = Some(...)) in the procedure containing the cursor.
    let proc_scope_range = host.with_source(uri, |symbols, source| {
        let offset = position_to_offset(source, position)?;
        // Find the procedure containing the cursor.
        let containing_proc = symbols
            .proc_ranges
            .iter()
            .find(|(_, range)| offset >= range.start as usize && offset <= range.end as usize)?;
        // Check if the word matches a proc-scoped symbol within that procedure.
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
    });

    let mut locations = Vec::new();

    if let Some(Some(proc_range)) = proc_scope_range {
        // Proc-scoped: only return occurrences within the procedure's byte range in this file.
        if let Some(locs) = host.with_source(uri, |_, source| {
            find_all_word_occurrences(source, &word)
                .into_iter()
                .filter(|r| r.start >= proc_range.start && r.end <= proc_range.end)
                .map(|r| Location {
                    uri: uri.clone(),
                    range: text_range_to_lsp_range(source, r),
                })
                .collect::<Vec<_>>()
        }) {
            locations = locs;
        }
    } else {
        // Module-level: search all files.
        for (file_uri, source) in host.all_file_sources() {
            for range in find_all_word_occurrences(&source, &word) {
                locations.push(Location {
                    uri: file_uri.clone(),
                    range: text_range_to_lsp_range(&source, range),
                });
            }
        }
    }

    locations
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::analysis::AnalysisHost;
    use crate::parser;

    #[test]
    fn proc_scoped_variable_references_only_within_procedure() {
        let host = AnalysisHost::new();
        let source = "Sub A()\n    Dim x As Long\n    x = 1\nEnd Sub\nSub B()\n    Dim x As Long\n    x = 2\nEnd Sub\n";
        let uri = Url::parse("file:///test.bas").unwrap();
        let parse_result = parser::parse(source);
        host.update(uri.clone(), source.to_string(), parse_result);

        // Cursor on `x` inside Sub A (line 1, col 8 => "Dim x")
        let refs = find_references(&host, &uri, Position::new(1, 8));
        // Should only find references within Sub A, not Sub B
        assert!(
            refs.len() == 2,
            "expected 2 references for proc-scoped x in Sub A, got {}",
            refs.len()
        );
        // All references should be within Sub A's lines (lines 0-3)
        for loc in &refs {
            assert!(
                loc.range.start.line < 4,
                "reference should be within Sub A, but found at line {}",
                loc.range.start.line
            );
        }
    }
}
