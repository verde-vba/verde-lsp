use tower_lsp::lsp_types::{Position, Url};
use verde_lsp::analysis::AnalysisHost;
use verde_lsp::document_highlight::document_highlight;
use verde_lsp::parser;

#[test]
fn document_highlight_returns_all_occurrences_in_file() {
    let uri: Url = "file:///test.bas".parse().unwrap();
    let source = "Sub Main()\n    Dim x As Integer\n    x = 1\n    x = x + 1\nEnd Sub\n";
    let host = AnalysisHost::new();
    host.update(uri.clone(), source.to_string(), parser::parse(source));

    // cursor on `x` at line 1 col 8 (in `Dim x As Integer`)
    let result = document_highlight(&host, &uri, Position::new(1, 8));

    assert!(result.is_some(), "expected document highlights, got None");
    let highlights = result.unwrap();
    // x appears 4 times: Dim x, x = 1, x = x + 1 (x twice)
    assert_eq!(
        highlights.len(),
        4,
        "expected 4 highlights, got {highlights:?}"
    );
}

#[test]
fn document_highlight_returns_none_for_unknown_position() {
    let uri: Url = "file:///test.bas".parse().unwrap();
    let source = "Sub Main()\nEnd Sub\n";
    let host = AnalysisHost::new();
    host.update(uri.clone(), source.to_string(), parser::parse(source));

    // cursor on `)` — not an identifier character
    let result = document_highlight(&host, &uri, Position::new(0, 9));

    assert!(
        result.is_none() || result.as_ref().map_or(false, |h| h.is_empty()),
        "expected None or empty for whitespace cursor"
    );
}
