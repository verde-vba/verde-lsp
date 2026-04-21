use tower_lsp::lsp_types::*;
use verde_lsp::analysis::AnalysisHost;

fn make_host_single(uri: &str, src: &str) -> (AnalysisHost, Url) {
    let url = Url::parse(uri).unwrap();
    let host = AnalysisHost::new();
    let parse_result = verde_lsp::parser::parse(src);
    host.update(url.clone(), src.to_string(), parse_result);
    (host, url)
}

fn make_host_two(uri1: &str, src1: &str, uri2: &str, src2: &str) -> (AnalysisHost, Url, Url) {
    let url1 = Url::parse(uri1).unwrap();
    let url2 = Url::parse(uri2).unwrap();
    let host = AnalysisHost::new();
    host.update(
        url1.clone(),
        src1.to_string(),
        verde_lsp::parser::parse(src1),
    );
    host.update(
        url2.clone(),
        src2.to_string(),
        verde_lsp::parser::parse(src2),
    );
    (host, url1, url2)
}

#[test]
fn prepare_call_hierarchy_returns_item_for_sub() {
    let src = "Sub Foo()\nEnd Sub\n";
    let (host, uri) = make_host_single("file:///test.bas", src);
    // cursor on "Foo" at line 0, character 4
    let pos = Position::new(0, 4);
    let items = host.prepare_call_hierarchy(&uri, pos);
    assert!(items.is_some(), "expected CallHierarchyItem for Sub");
    let items = items.unwrap();
    assert!(!items.is_empty());
    assert_eq!(items[0].name, "Foo");
}

#[test]
fn incoming_calls_returns_callers() {
    let src = "Sub Foo()\nEnd Sub\nSub Bar()\n    Foo\nEnd Sub\n";
    let (host, uri) = make_host_single("file:///test.bas", src);
    // prepare item for Foo
    let items = host
        .prepare_call_hierarchy(&uri, Position::new(0, 4))
        .unwrap();
    let item = items[0].clone();
    let calls = host.incoming_calls(&item);
    assert!(!calls.is_empty(), "expected Bar as caller of Foo");
    assert!(calls.iter().any(|c| c.from.name == "Bar"));
}

#[test]
fn incoming_calls_cross_file() {
    let src1 = "Sub Foo()\nEnd Sub\n";
    let src2 = "Sub Bar()\n    Foo\nEnd Sub\n";
    let (host, _uri1, _uri2) = make_host_two("file:///a.bas", src1, "file:///b.bas", src2);
    let url1 = Url::parse("file:///a.bas").unwrap();
    let items = host
        .prepare_call_hierarchy(&url1, Position::new(0, 4))
        .unwrap();
    let item = items[0].clone();
    let calls = host.incoming_calls(&item);
    assert!(!calls.is_empty(), "expected Bar in b.bas as caller");
    assert!(calls.iter().any(|c| c.from.name == "Bar"));
}

#[test]
fn outgoing_calls_returns_callees() {
    let src = "Sub Foo()\n    Bar\n    Baz\nEnd Sub\nSub Bar()\nEnd Sub\nSub Baz()\nEnd Sub\n";
    let (host, uri) = make_host_single("file:///test.bas", src);
    let items = host
        .prepare_call_hierarchy(&uri, Position::new(0, 4))
        .unwrap();
    let item = items[0].clone();
    let calls = host.outgoing_calls(&item);
    let names: Vec<&str> = calls.iter().map(|c| c.to.name.as_str()).collect();
    assert!(names.contains(&"Bar"), "expected Bar in outgoing calls");
    assert!(names.contains(&"Baz"), "expected Baz in outgoing calls");
}

#[test]
fn prepare_call_hierarchy_returns_none_for_non_procedure() {
    let src = "Sub Foo()\n    Dim x As String\nEnd Sub\n";
    let (host, uri) = make_host_single("file:///test.bas", src);
    // cursor on "x" (a variable, not a procedure)
    let items = host.prepare_call_hierarchy(&uri, Position::new(1, 8));
    // Either None or empty vec is acceptable
    let is_empty = items.map_or(true, |v| v.is_empty());
    assert!(is_empty, "variables should not return call hierarchy items");
}
