use tower_lsp::lsp_types::{Position, Range, Url};
use verde_lsp::analysis::AnalysisHost;
use verde_lsp::parser;
use verde_lsp::rename;

fn make_host_two_files(uri_a: &Url, src_a: &str, uri_b: &Url, src_b: &str) -> AnalysisHost {
    let host = AnalysisHost::new();
    host.update(uri_a.clone(), src_a.to_string(), parser::parse(src_a));
    host.update(uri_b.clone(), src_b.to_string(), parser::parse(src_b));
    host
}

fn do_rename(
    source: &str,
    position: Position,
    new_name: &str,
) -> Option<tower_lsp::lsp_types::WorkspaceEdit> {
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

    let edit = do_rename(source, position, "Bar").expect("expected WorkspaceEdit, got None");

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

    let edit = do_rename(source, position, "Baz").expect("expected WorkspaceEdit");

    let changes = edit.changes.expect("expected changes map");
    let uri: Url = "file:///test.bas".parse().unwrap();
    let edits = changes.get(&uri).expect("expected edits for file URI");

    assert_eq!(
        edits.len(),
        2,
        "expected 2 edits: declaration + call site, got {}",
        edits.len()
    );
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

    let edit =
        rename::rename(&host, &uri_a, Position::new(0, 11), "Bar").expect("expected WorkspaceEdit");
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

#[test]
fn rename_private_sub_stays_in_single_file() {
    // "Private Sub Foo()\nEnd Sub\n"
    //  0         1
    //  0123456789012345
    // 'F' is at column 12
    let uri_a: Url = "file:///module_a.bas".parse().unwrap();
    let src_a = "Private Sub Foo()\nEnd Sub\n";
    let uri_b: Url = "file:///module_b.bas".parse().unwrap();
    let src_b = "Private Sub Foo()\nEnd Sub\n";

    let host = make_host_two_files(&uri_a, src_a, &uri_b, src_b);
    let edit =
        rename::rename(&host, &uri_a, Position::new(0, 12), "Bar").expect("expected WorkspaceEdit");
    let changes = edit.changes.unwrap();

    assert!(changes.contains_key(&uri_a), "module_a should be renamed");
    assert!(
        !changes.contains_key(&uri_b),
        "module_b must NOT be renamed — Private symbols are module-scoped"
    );
}

#[test]
fn rename_local_var_stays_within_its_own_procedure() {
    // Single file with two Sub procs both declaring a local "x".
    // Renaming "x" in Sub Foo must NOT affect "x" in Sub Bar.
    //
    // Sub Foo()             <- line 0
    //     Dim x As Integer  <- line 1, 'x' at col 8
    //     x = 1             <- line 2, 'x' at col 4
    // End Sub               <- line 3
    // Sub Bar()             <- line 4
    //     Dim x As String   <- line 5
    //     x = 2             <- line 6
    // End Sub               <- line 7
    let source = "Sub Foo()\n    Dim x As Integer\n    x = 1\nEnd Sub\nSub Bar()\n    Dim x As String\n    x = 2\nEnd Sub\n";
    let position = Position::new(1, 8); // 'x' in Foo's Dim declaration

    let edit = do_rename(source, position, "myVar").expect("expected WorkspaceEdit");
    let changes = edit.changes.unwrap();
    let uri: Url = "file:///test.bas".parse().unwrap();
    let edits = changes.get(&uri).expect("expected edits for file URI");

    assert_eq!(
        edits.len(),
        2,
        "expected 2 renames (Foo's x only: decl + use), got {} — Bar's x must NOT be renamed",
        edits.len()
    );
    for e in edits {
        assert!(
            e.range.start.line < 4,
            "rename must stay within Sub Foo (lines 0-3), found edit at line {}",
            e.range.start.line
        );
    }
}

#[test]
fn rename_from_use_site_stays_within_its_procedure() {
    // Same two-procedure source: cursor is on the *use site* of x inside Foo (not the Dim).
    // The rename must still be constrained to Sub Foo.
    let source = "Sub Foo()\n    Dim x As Integer\n    x = 1\nEnd Sub\nSub Bar()\n    Dim x As String\n    x = 2\nEnd Sub\n";
    let position = Position::new(2, 4); // 'x' in "    x = 1" (use site in Foo)

    let edit = do_rename(source, position, "myVar").expect("expected WorkspaceEdit");
    let changes = edit.changes.unwrap();
    let uri: Url = "file:///test.bas".parse().unwrap();
    let edits = changes.get(&uri).expect("expected edits for file URI");

    assert_eq!(
        edits.len(),
        2,
        "expected 2 renames (Foo's x: decl + use site), got {} — Bar's x must NOT be renamed",
        edits.len()
    );
    for e in edits {
        assert!(
            e.range.start.line < 4,
            "rename must stay within Sub Foo (lines 0-3), found edit at line {}",
            e.range.start.line
        );
    }
}

#[test]
fn rename_local_variable_stays_in_single_file() {
    // "Sub Foo()\n    Dim x As Integer\n    x = 1\nEnd Sub\n"
    // line 1: "    Dim x As Integer"
    //          0123456789
    // 'x' is at column 8
    let uri_a: Url = "file:///module_a.bas".parse().unwrap();
    let src_a = "Sub Foo()\n    Dim x As Integer\n    x = 1\nEnd Sub\n";
    let uri_b: Url = "file:///module_b.bas".parse().unwrap();
    let src_b = "Sub Bar()\n    Dim x As String\n    x = \"hi\"\nEnd Sub\n";

    let host = make_host_two_files(&uri_a, src_a, &uri_b, src_b);
    let edit = rename::rename(&host, &uri_a, Position::new(1, 8), "myVar")
        .expect("expected WorkspaceEdit");
    let changes = edit.changes.unwrap();

    assert!(changes.contains_key(&uri_a), "module_a should be renamed");
    assert!(
        !changes.contains_key(&uri_b),
        "module_b must NOT be renamed — local variables are procedure-scoped"
    );
}
