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
fn hover_on_local_variable_shows_type() {
    // "Sub Foo()\n    Dim x As String\nEnd Sub\n"
    //  line 1:  "    Dim x As String"
    //                    ^ col 8
    let source = "Sub Foo()\n    Dim x As String\nEnd Sub\n";
    let position = Position::new(1, 8); // on 'x'
    let content = do_hover(source, position).expect("expected hover result for local variable");
    assert!(
        content.contains("x") && content.contains("String"),
        "expected hover to show 'x' with type 'String', got: {:?}",
        content
    );
}

#[test]
fn hover_on_symbol_from_other_module_shows_signature() {
    // File A defines Public Sub Foo
    let uri_a: Url = "file:///module_a.bas".parse().unwrap();
    let source_a = "Public Sub Foo()\nEnd Sub\n";
    let host = AnalysisHost::new();
    host.update(uri_a, source_a.to_string(), parser::parse(source_a));

    // File B references Foo
    let uri_b: Url = "file:///module_b.bas".parse().unwrap();
    let source_b = "Sub Bar()\n    Foo\nEnd Sub\n";
    host.update(uri_b.clone(), source_b.to_string(), parser::parse(source_b));

    // Hover on "Foo" in file B (line 1, col 4)
    let result = hover::hover(&host, &uri_b, Position::new(1, 4));
    assert!(
        result.is_some(),
        "expected hover result for 'Foo' from other module, got None"
    );
    let content = match result.unwrap().contents {
        tower_lsp::lsp_types::HoverContents::Markup(m) => m.value,
        _ => panic!("expected markup content"),
    };
    assert!(
        content.contains("Foo"),
        "expected hover content to mention 'Foo', got: {:?}",
        content
    );
}

#[test]
fn hover_on_sub_name_shows_parameter_list() {
    // "Sub Foo(x As Long, y As String)"
    //  0123456789...
    // "Foo" starts at column 4, line 0
    let source = "Sub Foo(x As Long, y As String)\nEnd Sub\n";
    let position = Position::new(0, 4); // on 'F' of "Foo"

    let content = do_hover(source, position).expect("expected hover result, got None");

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

// Two Sub procs both declaring a parameter named "x" but with different types.
// Hover in Sub A's body must show "Integer", hover in Sub B's body must show "String".
//
// Sub A(x As Integer)  <- line 0, 'x' at col 6
//     x = 1            <- line 1, 'x' at col 4  <- cursor (Test 1)
// End Sub              <- line 2
// Sub B(x As String)   <- line 3, 'x' at col 6
//     x = "hi"         <- line 4, 'x' at col 4  <- cursor (Test 2)
// End Sub              <- line 5
#[test]
fn hover_parameter_in_first_proc_shows_its_type() {
    let source =
        "Sub A(x As Integer)\n    x = 1\nEnd Sub\nSub B(x As String)\n    x = \"hi\"\nEnd Sub\n";
    let position = Position::new(1, 4); // 'x' in Sub A's body

    let content = do_hover(source, position).expect("expected hover result for 'x' in Sub A");
    assert!(
        content.contains("Integer"),
        "expected hover to show 'Integer' (Sub A's x), got: {:?}",
        content
    );
}

/// Hovering on `x` in `f.x` must show the UDT member's type annotation.
#[test]
fn hover_on_udt_member_access_shows_type() {
    // line 0: Type MyType
    // line 1:     x As Long
    // line 2: End Type
    // line 3: Sub Test()
    // line 4:     Dim f As MyType
    // line 5:     f.x               <- cursor on 'x', col 6
    // line 6: End Sub
    let source =
        "Type MyType\n    x As Long\nEnd Type\nSub Test()\n    Dim f As MyType\n    f.x\nEnd Sub\n";
    let position = Position::new(5, 6); // on 'x' in 'f.x'
    let content = do_hover(source, position).expect("expected hover result for UDT member access");
    assert!(
        content.contains("x") && content.contains("Long"),
        "expected hover to show 'x As Long', got: {content:?}"
    );
}

#[test]
fn hover_parameter_in_second_proc_shows_its_type() {
    let source =
        "Sub A(x As Integer)\n    x = 1\nEnd Sub\nSub B(x As String)\n    x = \"hi\"\nEnd Sub\n";
    let position = Position::new(4, 4); // 'x' in Sub B's body

    let content = do_hover(source, position).expect("expected hover result for 'x' in Sub B");
    assert!(
        content.contains("String"),
        "expected hover to show 'String' (Sub B's x), got: {:?}",
        content
    );
}

// ── Me. hover (PBI-44) ────────────────────────────────────────────────────────

/// Hover on the member part of `Me.Method` should show the method signature.
#[test]
fn hover_me_dot_member_shows_signature() {
    // line 0: Sub DoWork()
    // line 1: End Sub
    // line 2: Sub Test()
    // line 3:     Me.DoWork   <- cursor on 'D' of DoWork at col 7
    // line 4: End Sub
    let source = "Sub DoWork()\nEnd Sub\nSub Test()\n    Me.DoWork\nEnd Sub\n";
    let position = Position::new(3, 7); // on 'D' of DoWork
    let content = do_hover(source, position).expect("expected hover for Me.DoWork");
    assert!(
        content.contains("DoWork"),
        "expected hover to show 'DoWork', got: {content:?}"
    );
}
