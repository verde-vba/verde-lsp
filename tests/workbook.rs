use verde_lsp::analysis::AnalysisHost;

#[test]
fn reload_workbook_context_from_json_path() {
    let dir = std::env::temp_dir();
    let path = dir.join("verde_lsp_test_workbook_context.json");
    std::fs::write(
        &path,
        r#"{"sheets": ["Sheet1", "Summary"], "tables": ["T1"], "named_ranges": ["MyRange"]}"#,
    )
    .unwrap();

    let host = AnalysisHost::new();
    let ok = host.reload_workbook_context_from_path(&path);

    assert!(ok, "expected reload to succeed");
    assert!(
        host.workbook_sheets().contains(&"Sheet1".to_string()),
        "expected Sheet1 in sheets after reload"
    );
    assert!(
        host.workbook_tables().contains(&"T1".to_string()),
        "expected T1 in tables after reload"
    );
    assert!(
        host.workbook_named_ranges().contains(&"MyRange".to_string()),
        "expected MyRange in named_ranges after reload"
    );

    let _ = std::fs::remove_file(&path);
}

#[test]
fn reload_workbook_context_returns_false_for_missing_file() {
    let host = AnalysisHost::new();
    let ok = host.reload_workbook_context_from_path(std::path::Path::new("/nonexistent/path.json"));
    assert!(!ok, "expected reload to fail for missing file");
}
