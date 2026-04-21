use tower_lsp::lsp_types::*;
use verde_lsp::analysis::AnalysisHost;
use verde_lsp::parser;

fn make_host(uri: &str, src: &str) -> (AnalysisHost, Url) {
    let url = Url::parse(uri).unwrap();
    let host = AnalysisHost::new();
    let parse_result = parser::parse(src);
    host.update(url.clone(), src.to_string(), parse_result);
    (host, url)
}

#[test]
fn inlay_hint_shows_dim_variable_type() {
    let src = "Sub Foo()\n    Dim x As String\nEnd Sub\n";
    let (host, uri) = make_host("file:///test.bas", src);
    let hints = host.inlay_hints(&uri, None);
    let hint = hints
        .iter()
        .find(|h| matches!(&h.label, InlayHintLabel::String(s) if s == ": String"));
    assert!(hint.is_some(), "expected ': String' hint, got: {:?}", hints);
}

#[test]
fn inlay_hint_shows_variant_for_untyped_dim() {
    let src = "Sub Foo()\n    Dim x\nEnd Sub\n";
    let (host, uri) = make_host("file:///test.bas", src);
    let hints = host.inlay_hints(&uri, None);
    let hint = hints
        .iter()
        .find(|h| matches!(&h.label, InlayHintLabel::String(s) if s == ": Variant"));
    assert!(
        hint.is_some(),
        "expected ': Variant' hint, got: {:?}",
        hints
    );
}

#[test]
fn inlay_hint_shows_const_type() {
    let src = "Sub Foo()\n    Const PI As Double = 3.14\nEnd Sub\n";
    let (host, uri) = make_host("file:///test.bas", src);
    let hints = host.inlay_hints(&uri, None);
    let hint = hints
        .iter()
        .find(|h| matches!(&h.label, InlayHintLabel::String(s) if s == ": Double"));
    assert!(hint.is_some(), "expected ': Double' hint, got: {:?}", hints);
}

#[test]
fn inlay_hint_no_hint_for_procedures() {
    let src = "Sub Foo()\nEnd Sub\n";
    let (host, uri) = make_host("file:///test.bas", src);
    let hints = host.inlay_hints(&uri, None);
    assert!(
        hints.is_empty(),
        "procedures should not generate inlay hints"
    );
}
