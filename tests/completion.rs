use tower_lsp::lsp_types::{CompletionItemKind, Position, Url};
use verde_lsp::analysis::{AnalysisHost, WorkbookContext};
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
    assert!(
        found.is_none(),
        "expected 'x' NOT in completion inside Sub B"
    );
}

#[test]
fn completion_module_var_visible_everywhere() {
    // Module-level Dim var must appear as a completion candidate inside Sub A.
    let source = "Dim m As String\nSub A()\n    \nEnd Sub\n";
    let items = do_complete_at(source, 2, 4);
    let found = items.iter().find(|(label, _, _)| label == "m");
    assert!(
        found.is_some(),
        "expected module var 'm' visible inside Sub A"
    );
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
    assert!(
        found_b.is_none(),
        "expected param 'p' NOT visible inside Sub B"
    );
}

#[test]
fn completion_includes_public_symbols_from_other_files() {
    // File A defines a Public Sub
    let uri_a: Url = "file:///module_a.bas".parse().unwrap();
    let source_a = "Public Sub Foo()\nEnd Sub\n";
    let host = AnalysisHost::new();
    host.update(uri_a, source_a.to_string(), parser::parse(source_a));

    // File B is a separate module
    let uri_b: Url = "file:///module_b.bas".parse().unwrap();
    let source_b = "\n";
    host.update(uri_b.clone(), source_b.to_string(), parser::parse(source_b));

    // Completing in file B should surface Foo from file A
    let items: Vec<String> = completion::complete(&host, &uri_b, Position::new(0, 0))
        .into_iter()
        .map(|item| item.label)
        .collect();

    assert!(
        items.iter().any(|s| s == "Foo"),
        "expected 'Foo' from module_a in completion candidates for module_b, got: {:?}",
        items
    );
}

#[test]
fn completion_includes_workbook_sheet_names() {
    let uri: Url = "file:///test.bas".parse().unwrap();
    let src = "Sub Main()\nEnd Sub\n";
    let host = AnalysisHost::new();
    host.update(uri.clone(), src.to_string(), parser::parse(src));
    host.set_workbook_context(WorkbookContext {
        sheets: vec!["Sheet1".to_string(), "DataSheet".to_string()],
        ..Default::default()
    });

    let items: Vec<String> = completion::complete(&host, &uri, Position::new(0, 0))
        .into_iter()
        .map(|i| i.label)
        .collect();

    assert!(
        items.iter().any(|s| s == "Sheet1"),
        "expected 'Sheet1' from workbook context in completions, got: {items:?}"
    );
    assert!(
        items.iter().any(|s| s == "DataSheet"),
        "expected 'DataSheet' from workbook context in completions, got: {items:?}"
    );
}

#[test]
fn completion_includes_workbook_table_names() {
    let uri: Url = "file:///test.bas".parse().unwrap();
    let src = "Sub Main()\nEnd Sub\n";
    let host = AnalysisHost::new();
    host.update(uri.clone(), src.to_string(), parser::parse(src));
    host.set_workbook_context(WorkbookContext {
        tables: vec!["SalesTable".to_string()],
        ..Default::default()
    });

    let items: Vec<String> = completion::complete(&host, &uri, Position::new(0, 0))
        .into_iter()
        .map(|i| i.label)
        .collect();

    assert!(
        items.iter().any(|s| s == "SalesTable"),
        "expected 'SalesTable' from workbook tables in completions, got: {items:?}"
    );
}

// ── UDT dot-access completion (PBI-43) ───────────────────────────────────

/// `f.` after `Dim f As MyType` should offer only MyType's members — no keywords/builtins.
#[test]
fn dot_access_prefix_triggers_udt_member_completion() {
    // line 0: Type MyType
    // line 1:     x As Long
    // line 2:     name As String
    // line 3: End Type
    // line 4: Sub Test()
    // line 5:     Dim f As MyType
    // line 6:     f.        <- cursor col 6 (after dot)
    // line 7: End Sub
    let source =
        "Type MyType\n    x As Long\n    name As String\nEnd Type\nSub Test()\n    Dim f As MyType\n    f.\nEnd Sub\n";
    let items = do_complete_at(source, 6, 6);
    let labels: Vec<&str> = items.iter().map(|(l, _, _)| l.as_str()).collect();
    assert!(
        labels.contains(&"x"),
        "expected 'x' in dot-access completion, got: {labels:?}"
    );
    assert!(
        labels.contains(&"name"),
        "expected 'name' in dot-access completion, got: {labels:?}"
    );
    assert!(
        !labels.contains(&"Dim"),
        "keyword 'Dim' must not appear in dot-access completion, got: {labels:?}"
    );
}

/// Module-level variable `x` must not leak into dot-access results for a different type.
#[test]
fn dot_access_filters_non_udt_symbols() {
    // line 0: Dim x As Long      <- module var named 'x', should NOT appear
    // line 1: Type MyType
    // line 2:     y As Long      <- only UDT member should appear
    // line 3: End Type
    // line 4: Sub Test()
    // line 5:     Dim f As MyType
    // line 6:     f.             <- cursor col 6
    // line 7: End Sub
    let source =
        "Dim x As Long\nType MyType\n    y As Long\nEnd Type\nSub Test()\n    Dim f As MyType\n    f.\nEnd Sub\n";
    let items = do_complete_at(source, 6, 6);
    let labels: Vec<&str> = items.iter().map(|(l, _, _)| l.as_str()).collect();
    assert!(
        labels.contains(&"y"),
        "expected UDT member 'y' in dot-access completion, got: {labels:?}"
    );
    assert!(
        !labels.contains(&"x"),
        "module var 'x' must not appear in UDT dot-access, got: {labels:?}"
    );
}

/// Dot-access on a variable whose type has no TypeDef must return empty.
#[test]
fn dot_access_unknown_type_returns_empty() {
    // line 0: Sub Test()
    // line 1:     Dim g As Unknown
    // line 2:     g.             <- cursor col 6
    // line 3: End Sub
    let source = "Sub Test()\n    Dim g As Unknown\n    g.\nEnd Sub\n";
    let items = do_complete_at(source, 2, 6);
    assert!(
        items.is_empty(),
        "expected empty completion for unknown type dot-access, got: {items:?}"
    );
}

/// Proc-scoped `Dim f As MyType` must also resolve dot-access members.
#[test]
fn dot_access_procedure_scoped_variable_resolves() {
    // line 0: Type MyType
    // line 1:     x As Long
    // line 2: End Type
    // line 3: Sub Test()
    // line 4:     Dim f As MyType
    // line 5:     f.             <- cursor col 6
    // line 6: End Sub
    let source =
        "Type MyType\n    x As Long\nEnd Type\nSub Test()\n    Dim f As MyType\n    f.\nEnd Sub\n";
    let items = do_complete_at(source, 5, 6);
    let labels: Vec<&str> = items.iter().map(|(l, _, _)| l.as_str()).collect();
    assert!(
        labels.contains(&"x"),
        "expected 'x' from proc-scoped UDT variable, got: {labels:?}"
    );
}

#[test]
fn completion_includes_workbook_named_ranges() {
    let uri: Url = "file:///test.bas".parse().unwrap();
    let src = "Sub Main()\nEnd Sub\n";
    let host = AnalysisHost::new();
    host.update(uri.clone(), src.to_string(), parser::parse(src));
    host.set_workbook_context(WorkbookContext {
        named_ranges: vec!["MyRange".to_string()],
        ..Default::default()
    });

    let items: Vec<String> = completion::complete(&host, &uri, Position::new(0, 0))
        .into_iter()
        .map(|i| i.label)
        .collect();

    assert!(
        items.iter().any(|s| s == "MyRange"),
        "expected 'MyRange' from workbook named_ranges in completions, got: {items:?}"
    );
}

// ── Class module Me. completion (PBI-44) ─────────────────────────────────────

/// `Me.` should offer the class module's own procedures and module-level variables.
#[test]
fn me_dot_completion_returns_module_members() {
    // line 0: Sub DoWork()
    // line 1: End Sub
    // line 2: Private m_count As Long
    // line 3: Sub Test()
    // line 4:     Me.          <- cursor col 7 (after dot)
    // line 5: End Sub
    let source = "Sub DoWork()\nEnd Sub\nPrivate m_count As Long\nSub Test()\n    Me.\nEnd Sub\n";
    let items = do_complete_at(source, 4, 7);
    let labels: Vec<&str> = items.iter().map(|(l, _, _)| l.as_str()).collect();
    assert!(
        labels.contains(&"DoWork"),
        "expected 'DoWork' in Me. completion, got: {labels:?}"
    );
    assert!(
        labels.contains(&"m_count"),
        "expected 'm_count' in Me. completion, got: {labels:?}"
    );
    assert!(
        !labels.contains(&"Dim"),
        "keyword 'Dim' must not appear in Me. dot-access completion, got: {labels:?}"
    );
}

/// `Me.` completion must not offer proc-scoped local variables.
#[test]
fn me_dot_completion_excludes_local_variables() {
    // line 0: Sub Test()
    // line 1:     Dim local As Long
    // line 2:     Me.          <- cursor col 7
    // line 3: End Sub
    let source = "Sub Test()\n    Dim local As Long\n    Me.\nEnd Sub\n";
    let items = do_complete_at(source, 2, 7);
    let labels: Vec<&str> = items.iter().map(|(l, _, _)| l.as_str()).collect();
    assert!(
        !labels.contains(&"local"),
        "proc-local 'local' must not appear in Me. dot-access, got: {labels:?}"
    );
}

/// `Me.partial` partial input should still resolve (cursor mid-word after dot).
#[test]
fn me_dot_completion_with_partial_member() {
    // line 0: Sub GetValue() As Long
    // line 1: End Sub
    // line 2: Sub Test()
    // line 3:     Me.Get       <- cursor col 10 (after "Get")
    // line 4: End Sub
    let source = "Sub GetValue()\nEnd Sub\nSub Test()\n    Me.Get\nEnd Sub\n";
    let items = do_complete_at(source, 3, 10);
    let labels: Vec<&str> = items.iter().map(|(l, _, _)| l.as_str()).collect();
    assert!(
        labels.contains(&"GetValue"),
        "expected 'GetValue' in Me.Get partial completion, got: {labels:?}"
    );
}

// ── Excel builtin type dot-access completion (PBI-45) ───────────────────────

/// `Dim pt As PivotTable / pt.` should offer PivotTable's properties and methods.
#[test]
fn pt_dot_completion_returns_pivottable_members() {
    // line 0: Sub Test()
    // line 1:     Dim pt As PivotTable
    // line 2:     pt.          <- cursor col 7 (after dot)
    // line 3: End Sub
    let source = "Sub Test()\n    Dim pt As PivotTable\n    pt.\nEnd Sub\n";
    let items = do_complete_at(source, 2, 7);
    let labels: Vec<&str> = items.iter().map(|(l, _, _)| l.as_str()).collect();
    assert!(
        labels.contains(&"Name"),
        "expected 'Name' in PivotTable dot-access completion, got: {labels:?}"
    );
    assert!(
        labels.contains(&"DataBodyRange"),
        "expected 'DataBodyRange' in PivotTable dot-access completion, got: {labels:?}"
    );
    assert!(
        labels.contains(&"RefreshTable"),
        "expected 'RefreshTable' in PivotTable dot-access completion, got: {labels:?}"
    );
    assert!(
        !labels.contains(&"Dim"),
        "keyword 'Dim' must not appear in dot-access completion, got: {labels:?}"
    );
}

/// `Dim ch As Chart / ch.` should offer Chart's properties and methods.
#[test]
fn chart_dot_completion_returns_chart_members() {
    // line 0: Sub Test()
    // line 1:     Dim ch As Chart
    // line 2:     ch.          <- cursor col 7 (after dot)
    // line 3: End Sub
    let source = "Sub Test()\n    Dim ch As Chart\n    ch.\nEnd Sub\n";
    let items = do_complete_at(source, 2, 7);
    let labels: Vec<&str> = items.iter().map(|(l, _, _)| l.as_str()).collect();
    assert!(
        labels.contains(&"Name"),
        "expected 'Name' in Chart dot-access completion, got: {labels:?}"
    );
    assert!(
        labels.contains(&"ChartType"),
        "expected 'ChartType' in Chart dot-access completion, got: {labels:?}"
    );
    assert!(
        labels.contains(&"SetSourceData"),
        "expected 'SetSourceData' in Chart dot-access completion, got: {labels:?}"
    );
    assert!(
        !labels.contains(&"Dim"),
        "keyword 'Dim' must not appear in dot-access completion, got: {labels:?}"
    );
}

/// `Dim sh As Shape / sh.` should offer Shape's properties and methods.
#[test]
fn shape_dot_completion_returns_shape_members() {
    // line 0: Sub Test()
    // line 1:     Dim sh As Shape
    // line 2:     sh.          <- cursor col 7 (after dot)
    // line 3: End Sub
    let source = "Sub Test()\n    Dim sh As Shape\n    sh.\nEnd Sub\n";
    let items = do_complete_at(source, 2, 7);
    let labels: Vec<&str> = items.iter().map(|(l, _, _)| l.as_str()).collect();
    assert!(
        labels.contains(&"Name"),
        "expected 'Name' in Shape dot-access completion, got: {labels:?}"
    );
    assert!(
        labels.contains(&"Width"),
        "expected 'Width' in Shape dot-access completion, got: {labels:?}"
    );
    assert!(
        labels.contains(&"Visible"),
        "expected 'Visible' in Shape dot-access completion, got: {labels:?}"
    );
    assert!(
        !labels.contains(&"Dim"),
        "keyword 'Dim' must not appear in dot-access completion, got: {labels:?}"
    );
}

/// Existing Range dot-access completion must continue to work (regression guard).
#[test]
fn existing_range_dot_completion_still_works() {
    // line 0: Sub Test()
    // line 1:     Dim rng As Range
    // line 2:     rng.         <- cursor col 8 (after dot)
    // line 3: End Sub
    let source = "Sub Test()\n    Dim rng As Range\n    rng.\nEnd Sub\n";
    let items = do_complete_at(source, 2, 8);
    let labels: Vec<&str> = items.iter().map(|(l, _, _)| l.as_str()).collect();
    assert!(
        labels.contains(&"Value"),
        "expected 'Value' in Range dot-access completion, got: {labels:?}"
    );
    assert!(
        labels.contains(&"Address"),
        "expected 'Address' in Range dot-access completion, got: {labels:?}"
    );
    assert!(
        labels.contains(&"Select"),
        "expected 'Select' in Range dot-access completion, got: {labels:?}"
    );
}
