use tower_lsp::lsp_types::{HoverContents, Position, Url};
use verde_lsp::analysis::AnalysisHost;
use verde_lsp::hover;
use verde_lsp::parser;

fn do_hover(source: &str, position: Position) -> Option<String> {
    let uri: Url = "file:///test.bas".parse().unwrap();
    let host = AnalysisHost::new();
    let parse_result = parser::parse(source);
    host.update(uri.clone(), source.to_string(), parse_result);
    let result = hover::hover(&host, &uri, position)?;
    match result.contents {
        HoverContents::Markup(markup) => Some(markup.value),
        HoverContents::Scalar(_) => None,
        HoverContents::Array(_) => None,
    }
}

#[test]
fn hover_on_sub_name_shows_parameter_list() {
    // "Sub Foo(x As Long, y As String)"
    //  0123456789...
    // "Foo" starts at column 4, line 0
    let source = "Sub Foo(x As Long, y As String)\nEnd Sub\n";
    let position = Position::new(0, 4); // on 'F' of "Foo"

    let content = do_hover(source, position)
        .expect("expected hover result, got None");

    assert!(
        content.contains("x As Long"),
        "expected hover content to contain 'x As Long', got: {:?}",
        content
    );
    assert!(
        content.contains("y As String"),
        "expected hover content to contain 'y As String', got: {:?}",
        content
    );
}
