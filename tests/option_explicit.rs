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
