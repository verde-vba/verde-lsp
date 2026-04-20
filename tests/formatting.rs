use verde_lsp::formatting::apply_formatting;

// ── PBI-46 α: keyword case normalization + trailing whitespace removal ──
// ── PBI-46 β: indent normalization (depth tracking) ──

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
    assert_eq!(
        apply_formatting(input),
        "Function Bar() As Long\nEnd Function\n"
    );
}

#[test]
fn format_keyword_case_normalizes_mixed_block() {
    let input = "sub Foo()\ndim x as integer\nif x > 0 then\nx = 1\nend if\nend sub\n";
    assert_eq!(
        apply_formatting(input),
        "Sub Foo()\n    Dim x As Integer\n    If x > 0 Then\n        x = 1\n    End If\nEnd Sub\n"
    );
}

#[test]
fn format_trailing_whitespace_removed() {
    let input = "Dim x As Integer   \nSub Foo()   \nEnd Sub\n";
    assert_eq!(
        apply_formatting(input),
        "Dim x As Integer\nSub Foo()\nEnd Sub\n"
    );
}

#[test]
fn format_already_canonical_is_unchanged() {
    let input = "Sub Foo()\n    Dim x As Integer\nEnd Sub\n";
    assert_eq!(
        apply_formatting(input),
        "Sub Foo()\n    Dim x As Integer\nEnd Sub\n"
    );
}

#[test]
fn format_preserves_string_literals_intact() {
    let input = "Dim s As String\ns = \"hello DIM world\"\n";
    assert_eq!(
        apply_formatting(input),
        "Dim s As String\ns = \"hello DIM world\"\n"
    );
}

#[test]
fn format_preserves_identifiers_case() {
    let input = "Sub myProcedure()\nDim myVar As Long\nEnd Sub\n";
    assert_eq!(
        apply_formatting(input),
        "Sub myProcedure()\n    Dim myVar As Long\nEnd Sub\n"
    );
}

#[test]
fn format_indent_sub_body_indented_4_spaces() {
    let input = "Sub Foo()\nDim x As Long\nEnd Sub\n";
    assert_eq!(
        apply_formatting(input),
        "Sub Foo()\n    Dim x As Long\nEnd Sub\n"
    );
}

#[test]
fn format_indent_nested_if_double_indented() {
    let input = "Sub Foo()\nIf True Then\nx = 1\nEnd If\nEnd Sub\n";
    assert_eq!(
        apply_formatting(input),
        "Sub Foo()\n    If True Then\n        x = 1\n    End If\nEnd Sub\n"
    );
}

#[test]
fn format_indent_else_if_aligned_with_if() {
    // ElseIf/Else は If と同じ depth に揃える (VBA 慣習 — depth 一時 -1)
    let input = "Sub Foo()\nIf x > 0 Then\nx = 1\nElseIf x < 0 Then\nx = -1\nElse\nx = 0\nEnd If\nEnd Sub\n";
    assert_eq!(
        apply_formatting(input),
        "Sub Foo()\n    If x > 0 Then\n        x = 1\n    ElseIf x < 0 Then\n        x = -1\n    Else\n        x = 0\n    End If\nEnd Sub\n"
    );
}

#[test]
fn format_indent_select_case_aligned_with_select() {
    // Case は Select Case と同じ depth (depth 一時 -1)
    let input = "Sub Foo()\nSelect Case x\nCase 1\nx = 1\nCase Else\nx = 0\nEnd Select\nEnd Sub\n";
    assert_eq!(
        apply_formatting(input),
        "Sub Foo()\n    Select Case x\n    Case 1\n        x = 1\n    Case Else\n        x = 0\n    End Select\nEnd Sub\n"
    );
}

#[test]
fn format_indent_for_loop_body() {
    let input = "Sub Foo()\nFor i = 1 To 10\nx = i\nNext i\nEnd Sub\n";
    assert_eq!(
        apply_formatting(input),
        "Sub Foo()\n    For i = 1 To 10\n        x = i\n    Next i\nEnd Sub\n"
    );
}

#[test]
fn format_indent_with_block() {
    let input = "Sub Foo()\nWith obj\n.x = 1\n.y = 2\nEnd With\nEnd Sub\n";
    assert_eq!(
        apply_formatting(input),
        "Sub Foo()\n    With obj\n        .x = 1\n        .y = 2\n    End With\nEnd Sub\n"
    );
}

#[test]
fn format_indent_do_loop() {
    let input = "Sub Foo()\nDo While x > 0\nx = x - 1\nLoop\nEnd Sub\n";
    assert_eq!(
        apply_formatting(input),
        "Sub Foo()\n    Do While x > 0\n        x = x - 1\n    Loop\nEnd Sub\n"
    );
}

#[test]
fn format_indent_type_block() {
    let input = "Type MyType\nx As Long\ny As String\nEnd Type\n";
    assert_eq!(
        apply_formatting(input),
        "Type MyType\n    x As Long\n    y As String\nEnd Type\n"
    );
}

#[test]
fn format_indent_public_sub_open_token() {
    // Public/Private 修飾子があっても Sub 本体を正しくインデントする
    let input = "Public Sub Foo()\nDim x As Long\nEnd Sub\n";
    assert_eq!(
        apply_formatting(input),
        "Public Sub Foo()\n    Dim x As Long\nEnd Sub\n"
    );
}

#[test]
fn format_indent_preserves_blank_lines() {
    let input = "Sub Foo()\nDim x As Long\n\nDim y As Long\nEnd Sub\n";
    assert_eq!(
        apply_formatting(input),
        "Sub Foo()\n    Dim x As Long\n\n    Dim y As Long\nEnd Sub\n"
    );
}
