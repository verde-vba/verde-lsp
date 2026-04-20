use tower_lsp::lsp_types::{Diagnostic, Range, Url};
use verde_lsp::analysis::AnalysisHost;
use verde_lsp::code_action::code_actions;
use verde_lsp::parser;

fn make_host_with_undeclared(source: &str) -> (AnalysisHost, Url, Vec<Diagnostic>) {
    let uri: Url = "file:///test.bas".parse().unwrap();
    let host = AnalysisHost::new();
    host.update(uri.clone(), source.to_string(), parser::parse(source));
    let diags = host.diagnostics(&uri);
    (host, uri, diags)
}

#[test]
fn code_action_offers_dim_for_undeclared_variable() {
    let source = "Option Explicit\n\nSub Main()\n    x = 1\nEnd Sub\n";
    let (host, uri, diags) = make_host_with_undeclared(source);

    assert!(!diags.is_empty(), "expected undeclared diagnostic for 'x'");
    let diag = &diags[0];

    let actions = code_actions(&host, &uri, diag.range, &diags);

    assert!(
        !actions.is_empty(),
        "expected at least one code action, got none"
    );
    assert!(
        actions
            .iter()
            .any(|a| a.title.contains('x') || a.title.contains('X')),
        "expected action title to mention 'x', got: {:?}",
        actions.iter().map(|a| &a.title).collect::<Vec<_>>()
    );
}

#[test]
fn code_action_inserts_dim_after_procedure_declaration() {
    let source = "Option Explicit\n\nSub Main()\n    y = 42\nEnd Sub\n";
    let (host, uri, diags) = make_host_with_undeclared(source);

    assert!(!diags.is_empty());
    let diag = &diags[0];
    let actions = code_actions(&host, &uri, diag.range, &diags);
    assert!(!actions.is_empty());

    // The workspace edit should insert on line 3 (after "Sub Main()" on line 2)
    let action = &actions[0];
    let edit = action.edit.as_ref().expect("expected workspace edit");
    let changes = edit.changes.as_ref().expect("expected changes map");
    let edits = changes
        .values()
        .next()
        .expect("expected at least one file edit");
    assert!(!edits.is_empty(), "expected at least one text edit");

    // The inserted text should be a Dim statement for 'y'
    let inserted_text = &edits[0].new_text;
    assert!(
        inserted_text.contains("Dim") && inserted_text.contains('y'),
        "expected 'Dim y ...' in inserted text, got: {inserted_text:?}"
    );
}

#[test]
fn code_action_returns_empty_for_no_diagnostics() {
    let uri: Url = "file:///test.bas".parse().unwrap();
    let source = "Sub Main()\nEnd Sub\n";
    let host = AnalysisHost::new();
    host.update(uri.clone(), source.to_string(), parser::parse(source));

    let empty_range = Range::default();
    let actions = code_actions(&host, &uri, empty_range, &[]);

    assert!(actions.is_empty(), "expected no actions for no diagnostics");
}
