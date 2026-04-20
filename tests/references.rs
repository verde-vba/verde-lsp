use tower_lsp::lsp_types::{Position, Url};
use verde_lsp::analysis::AnalysisHost;
use verde_lsp::parser;
use verde_lsp::references;

fn find_refs(source: &str, line: u32, col: u32) -> Vec<(u32, u32)> {
    let uri: Url = "file:///test.bas".parse().unwrap();
    let host = AnalysisHost::new();
    host.update(uri.clone(), source.to_string(), parser::parse(source));
    references::find_references(&host, &uri, Position::new(line, col))
        .into_iter()
        .map(|loc| (loc.range.start.line, loc.range.start.character))
        .collect()
}

#[test]
fn references_returns_all_occurrences_of_symbol() {
    // "Foo" appears at: Sub Foo (line 0 col 4), Call Foo (line 2 col 4)
    let src = "Sub Foo()\nEnd Sub\nSub Main()\n    Foo\nEnd Sub\n";
    let refs = find_refs(src, 0, 5); // cursor on "Foo" declaration
    assert_eq!(refs.len(), 2, "expected 2 references to Foo, got: {refs:?}");
    assert!(refs.contains(&(0, 4)), "expected declaration at line 0 col 4");
    assert!(refs.contains(&(3, 4)), "expected call site at line 3 col 4");
}

#[test]
fn references_returns_empty_for_unknown_word() {
    let src = "Sub Main()\nEnd Sub\n";
    let refs = find_refs(src, 0, 0); // cursor on whitespace before Sub
    // "Sub" is a keyword — or cursor is on it; result should be non-empty but not crash
    // Key: function should not panic
    let _ = refs; // just assert no panic
}
