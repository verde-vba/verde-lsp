use tower_lsp::lsp_types::Url;
use verde_lsp::analysis::AnalysisHost;
use verde_lsp::parser;
use verde_lsp::workspace_symbol::workspace_symbols;

#[test]
fn workspace_symbol_returns_matching_public_procedures() {
    let uri_a: Url = "file:///module_a.bas".parse().unwrap();
    let src_a = "Public Sub FooBar()\nEnd Sub\n\nPublic Sub Baz()\nEnd Sub\n";

    let uri_b: Url = "file:///module_b.bas".parse().unwrap();
    let src_b = "Public Sub FooHelper()\nEnd Sub\n";

    let host = AnalysisHost::new();
    host.update(uri_a.clone(), src_a.to_string(), parser::parse(src_a));
    host.update(uri_b.clone(), src_b.to_string(), parser::parse(src_b));

    // query "Foo" should match FooBar and FooHelper but not Baz
    let results = workspace_symbols(&host, "Foo");
    assert_eq!(
        results.len(),
        2,
        "expected 2 symbols matching 'Foo', got: {results:?}"
    );
    assert!(
        results.iter().any(|s| s.name == "FooBar"),
        "expected FooBar in results"
    );
    assert!(
        results.iter().any(|s| s.name == "FooHelper"),
        "expected FooHelper in results"
    );
}

#[test]
fn workspace_symbol_empty_query_returns_all_module_level_symbols() {
    let uri: Url = "file:///test.bas".parse().unwrap();
    let src = "Public Sub Foo()\nEnd Sub\n\nPublic Function Bar() As Integer\nEnd Function\n";
    let host = AnalysisHost::new();
    host.update(uri.clone(), src.to_string(), parser::parse(src));

    let results = workspace_symbols(&host, "");
    assert!(
        results.len() >= 2,
        "expected at least Foo and Bar, got: {results:?}"
    );
}
