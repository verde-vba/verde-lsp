use tower_lsp::lsp_types::{CompletionItemKind, Position, Url};
use verde_lsp::analysis::AnalysisHost;
use verde_lsp::completion;
use verde_lsp::parser;

fn do_complete(source: &str) -> Vec<(String, Option<CompletionItemKind>, Option<String>)> {
    do_complete_at(source, 0, 0)
}

fn do_complete_at(
    source: &str,
    line: u32,
    col: u32,
) -> Vec<(String, Option<CompletionItemKind>, Option<String>)> {
    let uri: Url = "file:///test.bas".parse().unwrap();
    let host = AnalysisHost::new();
    let parse_result = parser::parse(source);
    host.update(uri.clone(), source.to_string(), parse_result);
    completion::complete(&host, &uri, Position::new(line, col))
        .into_iter()
        .map(|item| (item.label, item.kind, item.detail))
        .collect()
}

#[test]
fn completion_includes_local_dim_variable() {
    let source = "Sub Foo()\n    Dim x As String\nEnd Sub\n";
    let items = do_complete(source);
    let found = items.iter().find(|(label, _, _)| label == "x");
    assert!(found.is_some(), "expected 'x' in completion items");
    let (_, kind, detail) = found.unwrap();
    assert_eq!(*kind, Some(CompletionItemKind::VARIABLE));
    assert_eq!(detail.as_deref(), Some("String"));
}

#[test]
fn completion_includes_local_const() {
    let source = "Sub Foo()\n    Const PI As Double = 3.14\nEnd Sub\n";
    let items = do_complete(source);
    let found = items.iter().find(|(label, _, _)| label == "PI");
    assert!(found.is_some(), "expected 'PI' in completion items");
    let (_, kind, detail) = found.unwrap();
    assert_eq!(*kind, Some(CompletionItemKind::CONSTANT));
    assert_eq!(detail.as_deref(), Some("Double"));
}

// ── Scope-aware filtering (PBI-08) ─────────────────────────────────────────

#[test]
fn completion_local_var_not_visible_in_other_proc() {
    // "Sub A()\n    Dim x As Long\nEnd Sub\nSub B()\n    \nEnd Sub\n"
    // cursor at line 4, col 4 → inside Sub B body
    let source = "Sub A()\n    Dim x As Long\nEnd Sub\nSub B()\n    \nEnd Sub\n";
    let items = do_complete_at(source, 4, 4);
    let found = items.iter().find(|(label, _, _)| label == "x");
    assert!(found.is_none(), "expected 'x' NOT in completion inside Sub B");
}

#[test]
fn completion_module_var_visible_everywhere() {
    // Module-level var (Public, no Dim) must appear inside Sub A.
    // Note: bare module-level `Dim` is a parser limitation; use `Public` form.
    let source = "Public m As String\nSub A()\n    \nEnd Sub\n";
    let items = do_complete_at(source, 2, 4);
    let found = items.iter().find(|(label, _, _)| label == "m");
    assert!(found.is_some(), "expected module var 'm' visible inside Sub A");
}

#[test]
fn completion_param_visible_in_own_proc_only() {
    // "Sub A(p As String)\n    \nEnd Sub\nSub B()\n    \nEnd Sub\n"
    let source = "Sub A(p As String)\n    \nEnd Sub\nSub B()\n    \nEnd Sub\n";
    // inside Sub A: param must appear
    let items_a = do_complete_at(source, 1, 4);
    let found_a = items_a.iter().find(|(label, _, _)| label == "p");
    assert!(found_a.is_some(), "expected param 'p' visible inside Sub A");
    // inside Sub B: param must NOT appear
    let items_b = do_complete_at(source, 4, 4);
    let found_b = items_b.iter().find(|(label, _, _)| label == "p");
    assert!(found_b.is_none(), "expected param 'p' NOT visible inside Sub B");
}
