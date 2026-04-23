use tower_lsp::lsp_types::*;

use crate::analysis::resolve::offset_to_position;
use crate::analysis::symbols::SymbolKind;
use crate::analysis::AnalysisHost;

pub fn folding_ranges(host: &AnalysisHost, uri: &Url) -> Vec<FoldingRange> {
    host.with_source(uri, |symbols, source| {
        let mut ranges: Vec<FoldingRange> = symbols
            .proc_ranges
            .iter()
            .map(|(_, range)| {
                let start = offset_to_position(source, range.start as usize);
                let end = offset_to_position(source, range.end as usize);
                FoldingRange {
                    start_line: start.line,
                    start_character: None,
                    end_line: end.line,
                    end_character: None,
                    kind: Some(FoldingRangeKind::Region),
                    collapsed_text: None,
                }
            })
            .collect();

        // Control structure blocks (If/For/With/Select/Do/While) from block_ranges
        for br in &symbols.block_ranges {
            let start = offset_to_position(source, br.start as usize);
            let end = offset_to_position(source, br.end as usize);
            ranges.push(FoldingRange {
                start_line: start.line,
                start_character: None,
                end_line: end.line,
                end_character: None,
                kind: Some(FoldingRangeKind::Region),
                collapsed_text: None,
            });
        }

        // Type and Enum blocks
        for sym in &symbols.symbols {
            if matches!(sym.kind, SymbolKind::TypeDef | SymbolKind::EnumDef) {
                let start = offset_to_position(source, sym.span.start as usize);
                let end = offset_to_position(source, sym.span.end as usize);
                ranges.push(FoldingRange {
                    start_line: start.line,
                    start_character: None,
                    end_line: end.line,
                    end_character: None,
                    kind: Some(FoldingRangeKind::Region),
                    collapsed_text: None,
                });
            }
        }

        ranges
    })
    .unwrap_or_default()
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::analysis::AnalysisHost;

    fn setup_host(source: &str) -> (AnalysisHost, Url) {
        let host = AnalysisHost::new();
        let uri = Url::parse("file:///test.bas").unwrap();
        let parse_result = crate::parser::parse(source);
        host.update(uri.clone(), source.to_string(), parse_result);
        (host, uri)
    }

    #[test]
    fn type_block_generates_folding_range() {
        let source = "Type MyType\n    x As Long\n    y As String\nEnd Type\n";
        let (host, uri) = setup_host(source);
        let ranges = folding_ranges(&host, &uri);
        assert!(!ranges.is_empty(), "expected folding range for Type block");
        assert_eq!(ranges[0].start_line, 0);
    }

    #[test]
    fn enum_block_generates_folding_range() {
        let source = "Enum Color\n    Red\n    Green\n    Blue\nEnd Enum\n";
        let (host, uri) = setup_host(source);
        let ranges = folding_ranges(&host, &uri);
        assert!(!ranges.is_empty(), "expected folding range for Enum block");
        assert_eq!(ranges[0].start_line, 0);
    }

    #[test]
    fn procedure_folding_still_works() {
        let source = "Sub Foo()\n    Dim x As Long\nEnd Sub\n";
        let (host, uri) = setup_host(source);
        let ranges = folding_ranges(&host, &uri);
        assert!(!ranges.is_empty(), "expected folding range for Sub block");
    }

    // ── PLAN-14: Control structure folding ───────────────────────────

    #[test]
    fn if_block_generates_folding_range() {
        let source = "Sub Foo()\n    If True Then\n        x = 1\n    End If\nEnd Sub\n";
        let (host, uri) = setup_host(source);
        let ranges = folding_ranges(&host, &uri);
        // Should have Sub range + If range
        assert!(
            ranges.len() >= 2,
            "expected at least 2 folding ranges (Sub + If)"
        );
    }

    #[test]
    fn for_block_generates_folding_range() {
        let source = "Sub Foo()\n    For i = 1 To 10\n        x = i\n    Next i\nEnd Sub\n";
        let (host, uri) = setup_host(source);
        let ranges = folding_ranges(&host, &uri);
        assert!(
            ranges.len() >= 2,
            "expected at least 2 folding ranges (Sub + For)"
        );
    }

    #[test]
    fn with_block_generates_folding_range() {
        let source = "Sub Foo()\n    With rng\n        .Value = 1\n    End With\nEnd Sub\n";
        let (host, uri) = setup_host(source);
        let ranges = folding_ranges(&host, &uri);
        assert!(
            ranges.len() >= 2,
            "expected at least 2 folding ranges (Sub + With)"
        );
    }

    #[test]
    fn mixed_type_enum_and_proc_all_fold() {
        let source =
            "Type MyType\n    x As Long\nEnd Type\nEnum Color\n    Red\nEnd Enum\nSub Foo()\nEnd Sub\n";
        let (host, uri) = setup_host(source);
        let ranges = folding_ranges(&host, &uri);
        assert_eq!(
            ranges.len(),
            3,
            "expected 3 folding ranges (Type + Enum + Sub)"
        );
    }
}
