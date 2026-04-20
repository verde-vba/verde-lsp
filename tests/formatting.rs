use verde_lsp::formatting::apply_formatting;

// ── PBI-46 α: keyword case normalization + trailing whitespace removal ──

#[test]
fn format_keyword_case_normalizes_dim_as_integer() {
    let input = "dim x as integer\n";
    assert_eq!(apply_formatting(input), "Dim x As Integer\n");
}

#[test]
fn format_keyword_case_normalizes_sub_end_sub() {
    let input = "sub Foo()\nend sub\n";
    assert_eq!(apply_formatting(input), "Sub Foo()\nEnd Sub\n");
}

#[test]
fn format_keyword_case_normalizes_function() {
    let input = "function Bar() as long\nend function\n";
    assert_eq!(apply_formatting(input), "Function Bar() As Long\nEnd Function\n");
}

#[test]
fn format_keyword_case_normalizes_mixed_block() {
    let input = "sub Foo()\ndim x as integer\nif x > 0 then\nx = 1\nend if\nend sub\n";
    assert_eq!(
        apply_formatting(input),
        "Sub Foo()\nDim x As Integer\nIf x > 0 Then\nx = 1\nEnd If\nEnd Sub\n"
    );
}

#[test]
fn format_trailing_whitespace_removed() {
    let input = "Dim x As Integer   \nSub Foo()   \nEnd Sub\n";
    assert_eq!(apply_formatting(input), "Dim x As Integer\nSub Foo()\nEnd Sub\n");
}

#[test]
fn format_already_canonical_is_unchanged() {
    let input = "Sub Foo()\n    Dim x As Integer\nEnd Sub\n";
    assert_eq!(apply_formatting(input), "Sub Foo()\n    Dim x As Integer\nEnd Sub\n");
}

#[test]
fn format_preserves_string_literals_intact() {
    let input = "Dim s As String\ns = \"hello DIM world\"\n";
    assert_eq!(apply_formatting(input), "Dim s As String\ns = \"hello DIM world\"\n");
}

#[test]
fn format_preserves_identifiers_case() {
    let input = "Sub myProcedure()\nDim myVar As Long\nEnd Sub\n";
    assert_eq!(
        apply_formatting(input),
        "Sub myProcedure()\nDim myVar As Long\nEnd Sub\n"
    );
}
