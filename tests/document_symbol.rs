use tower_lsp::lsp_types::{SymbolKind, Url};
use verde_lsp::analysis::AnalysisHost;
use verde_lsp::document_symbol;
use verde_lsp::parser;

fn make_symbols(source: &str) -> Vec<tower_lsp::lsp_types::DocumentSymbol> {
    let uri: Url = "file:///test.bas".parse().unwrap();
    let host = AnalysisHost::new();
    host.update(uri.clone(), source.to_string(), parser::parse(source));
    document_symbol::document_symbols(&host, &uri)
}

#[test]
fn sub_is_top_level_function_symbol() {
    let src = "Sub Foo()\nEnd Sub\n";
    let syms = make_symbols(src);
    assert_eq!(syms.len(), 1);
    assert_eq!(syms[0].name, "Foo");
    assert_eq!(syms[0].kind, SymbolKind::FUNCTION);
}

#[test]
fn function_is_top_level_function_symbol() {
    let src = "Function Bar() As String\nEnd Function\n";
    let syms = make_symbols(src);
    assert_eq!(syms.len(), 1);
    assert_eq!(syms[0].name, "Bar");
    assert_eq!(syms[0].kind, SymbolKind::FUNCTION);
}

#[test]
fn procedure_parameters_are_children() {
    let src = "Sub Foo(x As Integer, y As String)\nEnd Sub\n";
    let syms = make_symbols(src);
    let children = syms[0].children.as_ref().expect("Foo should have children");
    let names: Vec<&str> = children.iter().map(|s| s.name.as_str()).collect();
    assert!(names.contains(&"x"), "expected param x, got: {names:?}");
    assert!(names.contains(&"y"), "expected param y, got: {names:?}");
}

#[test]
fn local_dim_is_child_of_procedure() {
    let src = "Sub Foo()\n    Dim counter As Long\nEnd Sub\n";
    let syms = make_symbols(src);
    let children = syms[0].children.as_ref().expect("Foo should have children");
    let names: Vec<&str> = children.iter().map(|s| s.name.as_str()).collect();
    assert!(names.contains(&"counter"), "expected local 'counter', got: {names:?}");
}

#[test]
fn module_level_dim_is_top_level() {
    let src = "Dim total As Long\nSub Foo()\nEnd Sub\n";
    let syms = make_symbols(src);
    let names: Vec<&str> = syms.iter().map(|s| s.name.as_str()).collect();
    assert!(names.contains(&"total"), "expected module-level 'total', got: {names:?}");
    let total = syms.iter().find(|s| s.name == "total").unwrap();
    assert_eq!(total.kind, SymbolKind::VARIABLE);
}

#[test]
fn empty_module_returns_empty() {
    let syms = make_symbols("");
    assert!(syms.is_empty());
}

#[test]
fn procedure_selection_range_is_name_span() {
    // "Sub Foo" — "Foo" starts at col 4
    let src = "Sub Foo()\nEnd Sub\n";
    let syms = make_symbols(src);
    assert_eq!(syms[0].selection_range.start.character, 4);
}

#[test]
fn procedure_range_covers_full_body() {
    // Full procedure spans lines 0-1
    let src = "Sub Foo()\nEnd Sub\n";
    let syms = make_symbols(src);
    assert_eq!(syms[0].range.start.line, 0);
    assert_eq!(syms[0].range.end.line, 1);
}
