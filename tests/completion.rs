use tower_lsp::lsp_types::{CompletionItemKind, Position, Url};
use verde_lsp::analysis::AnalysisHost;
use verde_lsp::completion;
use verde_lsp::parser;

fn do_complete(source: &str) -> Vec<(String, Option<CompletionItemKind>, Option<String>)> {
    let uri: Url = "file:///test.bas".parse().unwrap();
    let host = AnalysisHost::new();
    let parse_result = parser::parse(source);
    host.update(uri.clone(), source.to_string(), parse_result);
    completion::complete(&host, &uri, Position::new(0, 0))
        .into_iter()
        .map(|item| (item.label, item.kind, item.detail))
        .collect()
}

#[test]
fn completion_includes_local_dim_variable() {
    let source = "Sub Foo()\n    Dim x As String\nEnd Sub\n";
    let items = do_complete(source);
    let found = items.iter().find(|(label, _, _)| label == "x");
    assert!(found.is_some(), "expected 'x' in completion items");
    let (_, kind, detail) = found.unwrap();
    assert_eq!(*kind, Some(CompletionItemKind::VARIABLE));
    assert_eq!(detail.as_deref(), Some("String"));
}

#[test]
fn completion_includes_local_const() {
    let source = "Sub Foo()\n    Const PI As Double = 3.14\nEnd Sub\n";
    let items = do_complete(source);
    let found = items.iter().find(|(label, _, _)| label == "PI");
    assert!(found.is_some(), "expected 'PI' in completion items");
    let (_, kind, detail) = found.unwrap();
    assert_eq!(*kind, Some(CompletionItemKind::CONSTANT));
    assert_eq!(detail.as_deref(), Some("Double"));
}
