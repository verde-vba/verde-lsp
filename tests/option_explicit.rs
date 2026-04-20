use tower_lsp::lsp_types::{Diagnostic, DiagnosticSeverity, Url};
use verde_lsp::analysis::AnalysisHost;
use verde_lsp::parser;

fn diagnose(source: &str) -> Vec<Diagnostic> {
    let uri: Url = "file:///test.bas".parse().unwrap();
    let host = AnalysisHost::new();
    let parse_result = parser::parse(source);
    host.update(uri.clone(), source.to_string(), parse_result);
    host.diagnostics(&uri)
}

fn assert_no_diagnostics(source: &str, context: &str) {
    let diagnostics = diagnose(source);
    assert!(
        diagnostics.is_empty(),
        "expected zero diagnostics for {}, got {}: [{}]",
        context,
        diagnostics.len(),
        diagnostics
            .iter()
            .map(|d| format!("{:?}: {}", d.severity, d.message))
            .collect::<Vec<_>>()
            .join(", ")
    );
}

#[test]
fn warns_on_undeclared_variable_when_option_explicit_is_set() {
    let source = "Option Explicit\n\nSub Main()\n    x = 10\nEnd Sub\n";
    let diagnostics = diagnose(source);

    assert_eq!(
        diagnostics.len(),
        1,
        "expected exactly 1 diagnostic, got {}: {:?}",
        diagnostics.len(),
        diagnostics
    );
    let diag = &diagnostics[0];
    assert_eq!(
        diag.severity,
        Some(DiagnosticSeverity::WARNING),
        "expected Warning severity, got {:?}",
        diag.severity
    );
    assert!(
        diag.message.contains("x"),
        "expected message to contain 'x', got: {:?}",
        diag.message
    );
    assert!(
        diag.message.contains("Option Explicit"),
        "expected message to contain 'Option Explicit', got: {:?}",
        diag.message
    );
}

#[test]
fn does_not_warn_when_option_explicit_is_absent() {
    assert_no_diagnostics(
        "Sub Main()\n    zzz = 42\n    aaa = zzz + 1\nEnd Sub\n",
        "source without Option Explicit",
    );
}

#[test]
fn does_not_warn_on_for_next_loop() {
    assert_no_diagnostics(
        "Option Explicit\n\nSub Main()\n    Dim i As Long\n    For i = 1 To 10\n    Next i\nEnd Sub\n",
        "For/Next loop with declared variable",
    );
}

#[test]
fn does_not_warn_on_procedure_parameter_usage() {
    assert_no_diagnostics(
        "Option Explicit\n\nSub Foo(ByVal name As String)\n    MsgBox name\nEnd Sub\n",
        "procedure parameter usage in body",
    );
}

#[test]
fn does_not_warn_for_member_access_rhs_identifiers() {
    assert_no_diagnostics(
        "Option Explicit\n\nSub Main()\n    Dim ws As Worksheet\n    ws.Range(\"A1\").Value = 10\nEnd Sub\n",
        "member access on declared variable",
    );
}

#[test]
fn does_not_warn_on_for_each_with_declared_items() {
    assert_no_diagnostics(
        "Option Explicit\n\nSub ProcessSheets()\n    Dim ws As Worksheet\n    For Each ws In ActiveWorkbook.Worksheets\n        ws.Cells(1, 1).Value = \"x\"\n    Next ws\nEnd Sub\n",
        "For Each with declared loop variable",
    );
}

#[test]
fn option_explicit_flags_undeclared_in_if_header() {
    let source = "Option Explicit\n\nSub Demo()\n    If undeclaredFlag Then\n        Debug.Print 1\n    End If\nEnd Sub\n";
    let diagnostics = diagnose(source);

    assert!(
        diagnostics
            .iter()
            .any(|d| d.severity == Some(DiagnosticSeverity::WARNING)
                && d.message.contains("undeclaredFlag")),
        "expected a Warning diagnostic naming 'undeclaredFlag', got: {:?}",
        diagnostics
    );
}

#[test]
fn option_explicit_flags_undeclared_in_set_rhs() {
    let source = "Option Explicit\n\nSub Demo()\n    Dim target As Object\n    Set target = undeclaredSource\nEnd Sub\n";
    let diagnostics = diagnose(source);

    assert!(
        diagnostics
            .iter()
            .any(|d| d.severity == Some(DiagnosticSeverity::WARNING)
                && d.message.contains("undeclaredSource")),
        "expected a Warning diagnostic naming 'undeclaredSource', got: {:?}",
        diagnostics
    );
    assert!(
        !diagnostics.iter().any(|d| d.message.contains("target")),
        "did not expect any diagnostic naming 'target', got: {:?}",
        diagnostics
    );
}

#[test]
fn option_explicit_flags_undeclared_in_for_header() {
    // VBA: For loop where the bound expression uses an undeclared identifier
    // Dim lo As Long is declared; upperBound is NOT declared
    let src = r#"Option Explicit
Sub Demo()
    Dim lo As Long
    For lo = 1 To upperBound
        Debug.Print lo
    Next lo
End Sub"#;
    let diags = diagnose(src);
    assert!(
        diags.iter().any(|d| {
            d.severity == Some(DiagnosticSeverity::WARNING)
                && d.message.to_lowercase().contains("upperbound")
        }),
        "expected warning for undeclared `upperBound`, got: {diags:?}"
    );
    // lo is declared — must NOT be warned
    assert!(
        !diags
            .iter()
            .any(|d| d.message.to_lowercase().contains("'lo'")),
        "unexpected warning for declared `lo`, got: {diags:?}"
    );
}

#[test]
fn option_explicit_flags_undeclared_in_while_header() {
    let source = "Option Explicit\n\nSub Demo()\n    While undeclaredCond\n    Wend\nEnd Sub\n";
    let diags = diagnose(source);
    assert!(
        diags.iter().any(|d| {
            d.severity == Some(DiagnosticSeverity::WARNING) && d.message.contains("undeclaredCond")
        }),
        "expected Warning for undeclared 'undeclaredCond' in While header, got: {diags:?}"
    );
}

#[test]
fn does_not_warn_on_while_with_declared_condition() {
    assert_no_diagnostics(
        "Option Explicit\n\nSub Demo()\n    Dim running As Boolean\n    running = True\n    While running\n    Wend\nEnd Sub\n",
        "While loop with declared condition variable",
    );
}

#[test]
fn option_explicit_flags_undeclared_in_redim_bounds() {
    let source = "Option Explicit\nSub Demo()\n    Dim arr() As Long\n    ReDim arr(undeclaredSize)\nEnd Sub\n";
    let diags = diagnose(source);
    assert!(
        diags.iter().any(|d| d.message.contains("undeclaredSize")),
        "expected warning for undeclaredSize in ReDim bounds, got: {diags:?}"
    );
}

#[test]
fn option_explicit_flags_undeclared_in_do_while_header() {
    let source = "Option Explicit\nSub Demo()\n    Do While undeclaredCond\n    Loop\nEnd Sub\n";
    let diags = diagnose(source);
    assert!(
        diags.iter().any(|d| d.message.contains("undeclaredCond")),
        "expected warning for undeclaredCond in Do While header, got: {diags:?}"
    );
}

#[test]
fn option_explicit_flags_undeclared_in_do_until_header() {
    let source = "Option Explicit\nSub Demo()\n    Do Until undeclaredCond\n    Loop\nEnd Sub\n";
    let diags = diagnose(source);
    assert!(
        diags.iter().any(|d| d.message.contains("undeclaredCond")),
        "expected warning for undeclaredCond in Do Until header, got: {diags:?}"
    );
}

#[test]
fn does_not_warn_on_do_while_with_declared_condition() {
    assert_no_diagnostics(
        "Option Explicit\n\nSub Demo()\n    Dim running As Boolean\n    running = True\n    Do While running\n    Loop\nEnd Sub\n",
        "Do While loop with declared condition variable",
    );
}

#[test]
fn qualified_module_name_not_flagged_as_undeclared() {
    // ModuleA.bas defines Public Sub Foo; module_b.bas calls ModuleA.Foo
    // "ModuleA" is extracted from the URI filename — must NOT produce undeclared warning
    let uri_a: Url = "file:///workspace/ModuleA.bas".parse().unwrap();
    let src_a = "Public Sub Foo()\nEnd Sub\n";

    let uri_b: Url = "file:///workspace/module_b.bas".parse().unwrap();
    // "M" in "ModuleA" is col=4; after_dot skips "Foo" — only "ModuleA" is at risk
    let src_b = "Option Explicit\n\nSub Main()\n    ModuleA.Foo\nEnd Sub\n";

    let host = AnalysisHost::new();
    host.update(uri_a.clone(), src_a.to_string(), parser::parse(src_a));
    host.update(uri_b.clone(), src_b.to_string(), parser::parse(src_b));

    let diags = host.diagnostics(&uri_b);
    assert!(
        !diags
            .iter()
            .any(|d| d.message.to_lowercase().contains("modulea")),
        "expected no undeclared warning for ModuleA (it is a known module name), got: {diags:?}"
    );
}

#[test]
fn qualified_module_name_truly_unknown_still_detected() {
    // UnknownMod is not a registered file — must still warn
    let uri_a: Url = "file:///workspace/ModuleA.bas".parse().unwrap();
    let src_a = "Public Sub Foo()\nEnd Sub\n";

    let uri_b: Url = "file:///workspace/module_b.bas".parse().unwrap();
    let src_b = "Option Explicit\n\nSub Main()\n    UnknownMod.Bar\nEnd Sub\n";

    let host = AnalysisHost::new();
    host.update(uri_a.clone(), src_a.to_string(), parser::parse(src_a));
    host.update(uri_b.clone(), src_b.to_string(), parser::parse(src_b));

    let diags = host.diagnostics(&uri_b);
    assert!(
        diags
            .iter()
            .any(|d| d.message.to_lowercase().contains("unknownmod")),
        "expected undeclared warning for UnknownMod (not a registered module), got: {diags:?}"
    );
}

#[test]
fn cross_module_public_symbol_not_flagged_as_undeclared() {
    // Module A: "F" in "Foo" is col=0 on its own line
    let uri_a: Url = "file:///module_a.bas".parse().unwrap();
    let src_a = "Public Sub Foo()\nEnd Sub\n";

    // Module B: Option Explicit + calls Foo — must NOT produce an undeclared warning for Foo
    let uri_b: Url = "file:///module_b.bas".parse().unwrap();
    let src_b = "Option Explicit\n\nSub Main()\n    Foo\nEnd Sub\n";

    let host = AnalysisHost::new();
    host.update(uri_a.clone(), src_a.to_string(), parser::parse(src_a));
    host.update(uri_b.clone(), src_b.to_string(), parser::parse(src_b));

    let diags = host.diagnostics(&uri_b);
    assert!(
        !diags.iter().any(|d| d.message.contains("Foo")),
        "expected no undeclared warning for Foo (defined as Public in other module), got: {diags:?}"
    );
}

#[test]
fn cross_module_truly_undeclared_still_detected() {
    let uri_a: Url = "file:///module_a.bas".parse().unwrap();
    let src_a = "Public Sub Foo()\nEnd Sub\n";

    // Bar is NOT defined in any module — must still warn
    let uri_b: Url = "file:///module_b.bas".parse().unwrap();
    let src_b = "Option Explicit\n\nSub Main()\n    Bar\nEnd Sub\n";

    let host = AnalysisHost::new();
    host.update(uri_a.clone(), src_a.to_string(), parser::parse(src_a));
    host.update(uri_b.clone(), src_b.to_string(), parser::parse(src_b));

    let diags = host.diagnostics(&uri_b);
    assert!(
        diags.iter().any(|d| d.message.contains("Bar")),
        "expected undeclared warning for Bar (not defined anywhere), got: {diags:?}"
    );
}
