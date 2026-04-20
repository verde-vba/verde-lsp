use tower_lsp::lsp_types::{Position, Url};
use verde_lsp::analysis::AnalysisHost;
use verde_lsp::parser;
use verde_lsp::signature_help::signature_help;

#[test]
fn signature_help_returns_label_for_sub_with_single_param() {
    let uri: Url = "file:///test.bas".parse().unwrap();
    let source = "Sub Foo(x As Integer)\nEnd Sub\n\nSub Main()\n    Foo()\nEnd Sub\n";
    let host = AnalysisHost::new();
    host.update(uri.clone(), source.to_string(), parser::parse(source));

    // cursor at {4, 8} = just inside `Foo(|)`, after the `(`
    let result = signature_help(&host, &uri, Position::new(4, 8));

    assert!(result.is_some(), "expected SignatureHelp, got None");
    let sh = result.unwrap();
    assert!(!sh.signatures.is_empty(), "expected at least one signature");
    assert_eq!(sh.signatures[0].label, "Sub Foo(x As Integer)");
}

#[test]
fn signature_help_returns_active_parameter_index_for_second_arg() {
    let uri: Url = "file:///test.bas".parse().unwrap();
    let source =
        "Sub Foo(x As Integer, y As String)\nEnd Sub\n\nSub Main()\n    Foo(1, 2)\nEnd Sub\n";
    let host = AnalysisHost::new();
    host.update(uri.clone(), source.to_string(), parser::parse(source));

    // cursor at {4, 11} = at `2` in `    Foo(1, 2)` — second argument
    let result = signature_help(&host, &uri, Position::new(4, 11));

    assert!(result.is_some());
    let sh = result.unwrap();
    assert_eq!(
        sh.active_parameter,
        Some(1),
        "expected active_parameter=1 for second argument"
    );
}

#[test]
fn signature_help_cross_module_public_function() {
    let uri_a: Url = "file:///module_a.bas".parse().unwrap();
    let src_a = "Public Sub Helper(n As Long)\nEnd Sub\n";

    let uri_b: Url = "file:///module_b.bas".parse().unwrap();
    let src_b = "Sub Main()\n    Helper()\nEnd Sub\n";

    let host = AnalysisHost::new();
    host.update(uri_a.clone(), src_a.to_string(), parser::parse(src_a));
    host.update(uri_b.clone(), src_b.to_string(), parser::parse(src_b));

    // cursor at {1, 11} = just inside `Helper(|)` in module_b
    let result = signature_help(&host, &uri_b, Position::new(1, 11));

    assert!(
        result.is_some(),
        "expected SignatureHelp for cross-module Public function"
    );
    let sh = result.unwrap();
    assert_eq!(sh.signatures[0].label, "Sub Helper(n As Long)");
}
