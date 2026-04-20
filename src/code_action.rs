use std::collections::HashMap;

use tower_lsp::lsp_types::*;

use crate::analysis::resolve::{offset_to_position, position_to_offset};
use crate::analysis::AnalysisHost;

const UNDECLARED_MARKER: &str = "is not declared (Option Explicit)";

pub fn code_actions(
    host: &AnalysisHost,
    uri: &Url,
    _range: Range,
    diagnostics: &[Diagnostic],
) -> Vec<CodeAction> {
    let mut actions = Vec::new();

    for diag in diagnostics {
        let Some(var_name) = extract_undeclared_name(&diag.message) else {
            continue;
        };
        let Some(insert_range) = find_dim_insert_position(host, uri, diag.range) else {
            continue;
        };
        let indent = "    ";
        let new_text = format!("{indent}Dim {var_name} As Variant\n");

        let mut changes: HashMap<Url, Vec<TextEdit>> = HashMap::new();
        changes.insert(
            uri.clone(),
            vec![TextEdit {
                range: Range::new(insert_range, insert_range),
                new_text,
            }],
        );

        actions.push(CodeAction {
            title: format!("Declare '{var_name}' As Variant"),
            kind: Some(CodeActionKind::QUICKFIX),
            diagnostics: Some(vec![diag.clone()]),
            edit: Some(WorkspaceEdit {
                changes: Some(changes),
                ..Default::default()
            }),
            ..Default::default()
        });
    }

    actions
}

/// Extract the variable name from messages like:
/// `"Variable 'x' is not declared (Option Explicit)"`
fn extract_undeclared_name(message: &str) -> Option<String> {
    if !message.contains(UNDECLARED_MARKER) {
        return None;
    }
    let start = message.find('\'')?;
    let rest = &message[start + 1..];
    let end = rest.find('\'')?;
    Some(rest[..end].to_string())
}

/// Find the position to insert a `Dim` statement: the beginning of the line
/// immediately after the procedure header that contains `diag_range`.
fn find_dim_insert_position(host: &AnalysisHost, uri: &Url, diag_range: Range) -> Option<Position> {
    host.with_source(uri, |symbols, source| {
        let cursor_offset = position_to_offset(source, diag_range.start)?;

        // Find the procedure whose range contains the diagnostic.
        let proc_start = symbols
            .proc_ranges
            .iter()
            .find(|(_, r)| cursor_offset >= r.start as usize && cursor_offset <= r.end as usize)
            .map(|(_, r)| r.start as usize)?;

        // The insert position is the start of the line after the proc header line.
        let proc_header_pos = offset_to_position(source, proc_start);
        Some(Position::new(proc_header_pos.line + 1, 0))
    })?
}
