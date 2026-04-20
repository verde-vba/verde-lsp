use tower_lsp::lsp_types::{GotoDefinitionResponse, Position, Url};
use verde_lsp::analysis::AnalysisHost;
use verde_lsp::definition;
use verde_lsp::parser;

fn do_goto(source: &str, position: Position) -> Option<Position> {
    let uri: Url = "file:///test.bas".parse().unwrap();
    let host = AnalysisHost::new();
    let parse_result = parser::parse(source);
    host.update(uri.clone(), source.to_string(), parse_result);
    let resp = definition::goto_definition(&host, &uri, position)?;
    match resp {
        GotoDefinitionResponse::Scalar(loc) => Some(loc.range.start),
        _ => None,
    }
}

// "Sub Foo()\nEnd Sub\n\nSub Bar()\n    Call Foo\nEnd Sub\n"
//  line 0: Sub Foo()      — "Foo" at col 4
//  line 3: Sub Bar()
//  line 4:     Call Foo   — "Foo" at col 9
#[test]
fn goto_definition_from_call_statement_jumps_to_sub() {
    let source = "Sub Foo()\nEnd Sub\n\nSub Bar()\n    Call Foo\nEnd Sub\n";
    let call_pos = Position::new(4, 9); // on 'F' of "Foo" in "    Call Foo"
    let def = do_goto(source, call_pos).expect("expected definition location");
    // "Foo" declaration is at line 0, col 4
    assert_eq!(
        def.line, 0,
        "expected definition on line 0 (Sub Foo), got {}",
        def.line
    );
    assert_eq!(
        def.character, 4,
        "expected col 4 ('F' of Foo), got {}",
        def.character
    );
}

// "Sub Foo()\nEnd Sub\n\nSub Bar()\n    Foo 1\nEnd Sub\n"
//  line 4:     Foo 1  — bare call without Call keyword, "Foo" at col 4
#[test]
fn goto_definition_from_bare_call_jumps_to_sub() {
    let source = "Sub Foo()\nEnd Sub\n\nSub Bar()\n    Foo 1\nEnd Sub\n";
    let call_pos = Position::new(4, 4); // on 'F' of "Foo" in "    Foo 1"
    let def = do_goto(source, call_pos).expect("expected definition location");
    assert_eq!(
        def.line, 0,
        "expected definition on line 0 (Sub Foo), got {}",
        def.line
    );
    assert_eq!(
        def.character, 4,
        "expected col 4 ('F' of Foo), got {}",
        def.character
    );
}

// Local variable goto definition: cursor on "x" usage → jumps to Dim x declaration
// "Sub Foo()\n    Dim x As String\n    x = \"hi\"\nEnd Sub\n"
//  line 0: Sub Foo()
//  line 1:     Dim x As String  — "x" at col 8
//  line 2:     x = "hi"        — "x" at col 4 (usage)
#[test]
fn goto_definition_from_local_variable_usage_jumps_to_dim() {
    let source = "Sub Foo()\n    Dim x As String\n    x = \"hi\"\nEnd Sub\n";
    let usage_pos = Position::new(2, 4); // on 'x' in "    x = ..."
    let def = do_goto(source, usage_pos).expect("expected definition location for local variable");
    // Should jump to 'x' in "    Dim x As String" — col 8
    assert_eq!(
        def.line, 1,
        "expected definition on line 1 (Dim x), got {}",
        def.line
    );
    assert_eq!(
        def.character, 8,
        "expected col 8 ('x' in Dim), got {}",
        def.character
    );
}

#[test]
fn goto_definition_crosses_module_boundary() {
    // File A defines Public Sub Foo
    let uri_a: Url = "file:///module_a.bas".parse().unwrap();
    let source_a = "Public Sub Foo()\nEnd Sub\n";
    let host = AnalysisHost::new();
    host.update(uri_a.clone(), source_a.to_string(), parser::parse(source_a));

    // File B calls Foo
    let uri_b: Url = "file:///module_b.bas".parse().unwrap();
    let source_b = "Sub Bar()\n    Call Foo\nEnd Sub\n";
    host.update(uri_b.clone(), source_b.to_string(), parser::parse(source_b));

    // goto_definition from "Foo" in file B (line 1, col 9)
    let resp = definition::goto_definition(&host, &uri_b, Position::new(1, 9));
    assert!(
        resp.is_some(),
        "expected goto_definition result for cross-module 'Foo', got None"
    );
    match resp.unwrap() {
        tower_lsp::lsp_types::GotoDefinitionResponse::Scalar(loc) => {
            assert_eq!(
                loc.uri, uri_a,
                "expected definition to point to module_a.bas, got {:?}",
                loc.uri
            );
            assert_eq!(
                loc.range.start.line, 0,
                "expected definition on line 0 of module_a, got {}",
                loc.range.start.line
            );
        }
        _ => panic!("expected scalar goto_definition response"),
    }
}
