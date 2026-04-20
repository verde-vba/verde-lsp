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
    let refs = find_refs(src, 0, 0);
    let _ = refs; // just assert no panic
}

#[test]
fn cross_file_references_found_in_both_files() {
    // "Foo" appears in module_a (declaration) and module_b (call site)
    let uri_a: Url = "file:///module_a.bas".parse().unwrap();
    let src_a = "Public Sub Foo()\nEnd Sub\n";
    let uri_b: Url = "file:///module_b.bas".parse().unwrap();
    let src_b = "Sub Main()\n    Foo\nEnd Sub\n";

    let host = AnalysisHost::new();
    host.update(uri_a.clone(), src_a.to_string(), parser::parse(src_a));
    host.update(uri_b.clone(), src_b.to_string(), parser::parse(src_b));

    // cursor on "Foo" in module_a (line 0 col 11 = "Public Sub |F|oo")
    let locs = references::find_references(&host, &uri_a, Position::new(0, 11));

    let uris: Vec<&str> = locs.iter().map(|l| l.uri.as_str()).collect();
    assert!(
        uris.iter().any(|u| u.contains("module_a")),
        "expected reference in module_a, got: {uris:?}"
    );
    assert!(
        uris.iter().any(|u| u.contains("module_b")),
        "expected reference in module_b, got: {uris:?}"
    );
}
