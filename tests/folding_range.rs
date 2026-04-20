use tower_lsp::lsp_types::Url;
use verde_lsp::analysis::AnalysisHost;
use verde_lsp::folding_range::folding_ranges;
use verde_lsp::parser;

#[test]
fn folding_ranges_returns_one_range_per_procedure() {
    let uri: Url = "file:///test.bas".parse().unwrap();
    let source = "Sub Foo()\n    Dim x As Integer\nEnd Sub\n\nSub Bar()\nEnd Sub\n";
    let host = AnalysisHost::new();
    host.update(uri.clone(), source.to_string(), parser::parse(source));

    let ranges = folding_ranges(&host, &uri);

    // Expect 2 folding ranges (one per Sub)
    assert_eq!(ranges.len(), 2, "expected 2 fold ranges, got: {ranges:?}");
}

#[test]
fn folding_range_spans_full_procedure_body() {
    let uri: Url = "file:///test.bas".parse().unwrap();
    let source = "Sub Foo()\n    Dim x As Integer\n    x = 1\nEnd Sub\n";
    let host = AnalysisHost::new();
    host.update(uri.clone(), source.to_string(), parser::parse(source));

    let ranges = folding_ranges(&host, &uri);

    assert_eq!(ranges.len(), 1, "expected 1 fold range");
    let r = &ranges[0];
    assert_eq!(r.start_line, 0, "Sub starts at line 0");
    assert_eq!(r.end_line, 3, "End Sub is at line 3");
}
