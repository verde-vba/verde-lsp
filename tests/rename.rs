use tower_lsp::lsp_types::{Position, Range, Url};
use verde_lsp::analysis::AnalysisHost;
use verde_lsp::parser;
use verde_lsp::rename;

fn do_rename(source: &str, position: Position, new_name: &str) -> Option<tower_lsp::lsp_types::WorkspaceEdit> {
    let uri: Url = "file:///test.bas".parse().unwrap();
    let host = AnalysisHost::new();
    let parse_result = parser::parse(source);
    host.update(uri.clone(), source.to_string(), parse_result);
    rename::rename(&host, &uri, position, new_name)
}

#[test]
fn rename_procedure_name_returns_workspace_edit() {
    // "Sub Foo()\nEnd Sub"
    //  0123456789
    // "Foo" starts at column 4, line 0
    let source = "Sub Foo()\nEnd Sub";
    let position = Position::new(0, 4); // on 'F' of "Foo"

    let edit = do_rename(source, position, "Bar")
        .expect("expected WorkspaceEdit, got None");

    let changes = edit.changes.expect("expected changes map");
    let uri: Url = "file:///test.bas".parse().unwrap();
    let edits = changes.get(&uri).expect("expected edits for file URI");

    assert_eq!(edits.len(), 1, "expected exactly one text edit");
    let text_edit = &edits[0];
    assert_eq!(
        text_edit.new_text, "Bar",
        "expected new_text to be 'Bar', got: {:?}",
        text_edit.new_text
    );
    // "Foo" is at line 0, columns 4..7
    let expected_range = Range::new(Position::new(0, 4), Position::new(0, 7));
    assert_eq!(
        text_edit.range, expected_range,
        "expected range {:?}, got {:?}",
        expected_range, text_edit.range
    );
}
