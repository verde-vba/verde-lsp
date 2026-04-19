use tower_lsp::lsp_types::*;

use super::symbols::SymbolTable;
use crate::parser::ParseResult;
use super::resolve::text_range_to_lsp_range;

pub fn compute(parse_result: &ParseResult, _symbols: &SymbolTable) -> Vec<Diagnostic> {
    let mut diagnostics = Vec::new();
    let source = "";

    for error in &parse_result.errors {
        diagnostics.push(Diagnostic {
            range: Range::new(
                Position::new(0, error.span.start as u32),
                Position::new(0, error.span.end as u32),
            ),
            severity: Some(DiagnosticSeverity::ERROR),
            source: Some("verde-lsp".to_string()),
            message: error.message.clone(),
            ..Default::default()
        });
    }

    // TODO: Option Explicit diagnostics — warn on undeclared variables

    diagnostics
}
