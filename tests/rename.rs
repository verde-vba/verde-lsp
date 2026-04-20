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

#[test]
fn rename_includes_call_site_in_workspace_edit() {
    // Sub Foo() — declaration at line 0, col 4
    // Sub Bar()
    //     Call Foo() — call site at line 3, col 9
    let source = "Sub Foo()\nEnd Sub\nSub Bar()\n    Call Foo()\nEnd Sub";
    let position = Position::new(0, 4); // on 'F' of declaration "Foo"

    let edit = do_rename(source, position, "Baz")
        .expect("expected WorkspaceEdit");

    let changes = edit.changes.expect("expected changes map");
    let uri: Url = "file:///test.bas".parse().unwrap();
    let edits = changes.get(&uri).expect("expected edits for file URI");

    assert_eq!(edits.len(), 2, "expected 2 edits: declaration + call site, got {}", edits.len());
}

#[test]
fn cross_file_rename_includes_other_file_occurrences() {
    // ModuleA has "Public Sub Foo()" — cursor on Foo (line 0, col 11)
    let uri_a: Url = "file:///module_a.bas".parse().unwrap();
    let src_a = "Public Sub Foo()\nEnd Sub\n";
    // ModuleB calls Foo at line 2, col 4
    let uri_b: Url = "file:///module_b.bas".parse().unwrap();
    let src_b = "Sub Main()\n    Foo\nEnd Sub\n";

    let host = AnalysisHost::new();
    host.update(uri_a.clone(), src_a.to_string(), parser::parse(src_a));
    host.update(uri_b.clone(), src_b.to_string(), parser::parse(src_b));

    let edit = rename::rename(&host, &uri_a, Position::new(0, 11), "Bar")
        .expect("expected WorkspaceEdit");
    let changes = edit.changes.expect("expected changes");

    assert!(
        changes.contains_key(&uri_b),
        "expected rename to include module_b, got keys: {:?}",
        changes.keys().collect::<Vec<_>>()
    );
}

#[test]
fn rename_from_call_site_in_other_file() {
    // cursor is in module_b on the call site "Foo" — Foo is defined in module_a
    let uri_a: Url = "file:///module_a.bas".parse().unwrap();
    let src_a = "Public Sub Foo()\nEnd Sub\n";
    let uri_b: Url = "file:///module_b.bas".parse().unwrap();
    // "Foo" in module_b: line 1 col 4
    let src_b = "Sub Main()\n    Foo\nEnd Sub\n";

    let host = AnalysisHost::new();
    host.update(uri_a.clone(), src_a.to_string(), parser::parse(src_a));
    host.update(uri_b.clone(), src_b.to_string(), parser::parse(src_b));

    // cursor on Foo call site in module_b
    let edit = rename::rename(&host, &uri_b, Position::new(1, 4), "Bar");
    assert!(
        edit.is_some(),
        "expected rename from call site in module_b to succeed, got None"
    );
    let changes = edit.unwrap().changes.unwrap();
    assert!(
        changes.contains_key(&uri_a),
        "expected module_a to be included in rename changes, got: {:?}",
        changes.keys().collect::<Vec<_>>()
    );
}
