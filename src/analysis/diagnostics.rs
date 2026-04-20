use tower_lsp::lsp_types::*;

use super::resolve::text_range_to_lsp_range;
use super::symbols::SymbolTable;
use crate::parser::ast::{Ast, AstNode, TextRange};
use crate::parser::lexer::{self, Token};
use crate::parser::ParseResult;
use crate::vba_builtins::{BUILTIN_FUNCTIONS, BUILTIN_TYPES, KEYWORDS};

pub fn compute(
    parse_result: &ParseResult,
    symbols: &SymbolTable,
    source: &str,
) -> Vec<Diagnostic> {
    let mut diagnostics = Vec::new();

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

    if parse_result.ast.option_explicit {
        diagnostics.extend(check_option_explicit(&parse_result.ast, source, symbols));
    }

    diagnostics
}

/// Happy-path Option Explicit diagnostic.
///
/// Scans all `Ident` tokens in the source and emits a Warning for any
/// identifier that is not:
/// - declared at module level (symbol table)
/// - a VBA keyword
/// - a builtin function
/// - a builtin type
///
/// Identifiers that fall inside a module-level *declaration span*
/// (Variable/Const/TypeDef/EnumDef) are skipped — those are the declaration
/// sites themselves, not references. Identifiers inside procedure bodies are
/// still scanned (the parser currently does not parse procedure bodies in
/// detail; a procedure's `span` covers the entire `Sub..End Sub`, so skipping
/// the procedure span wholesale would hide all body references we want to
/// check). As a compromise, we only skip the procedure's *signature* region
/// (name + declared parameters) heuristically by skipping the line containing
/// `Sub`/`Function`/`Property`.
///
/// Known deferred edge cases (Phase 2 happy path):
/// - Local `Dim` inside procedures: bodies aren't parsed, so locals will be
///   flagged as undeclared.
/// - Procedure parameters: not captured in the symbol table; also flagged.
/// - Member access (`x.Value`): the rhs after `.` is still lexed as `Ident`.
/// - `With` blocks, `For i = 1 To 10` loop variables, assignment lhs vs rhs.
pub fn check_option_explicit(
    ast: &Ast,
    source: &str,
    symbols: &SymbolTable,
) -> Vec<Diagnostic> {
    // Collect spans that should be skipped entirely. These cover module-level
    // *declarations* (Type/Enum/Variable/Const) where the contained
    // identifiers are declaration sites, not references.
    let mut skip_spans: Vec<TextRange> = Vec::new();

    // Collect procedure signature line spans — approximate: from procedure
    // span start to the first newline after it. Parameters declared in the
    // signature should not trigger undeclared warnings.
    let mut signature_spans: Vec<TextRange> = Vec::new();

    for (_, node) in ast.nodes.iter() {
        match node {
            AstNode::Variable(v) => skip_spans.push(v.span),
            AstNode::TypeDef(t) => skip_spans.push(t.span),
            AstNode::EnumDef(e) => skip_spans.push(e.span),
            AstNode::Procedure(p) => {
                let start = p.span.start as usize;
                let end_of_line = source[start..]
                    .find('\n')
                    .map(|off| start + off)
                    .unwrap_or(source.len());
                signature_spans.push(TextRange::new(start, end_of_line));
            }
            _ => {}
        }
    }

    // Build the "declared or builtin" lowercase set.
    let mut declared: std::collections::HashSet<String> =
        std::collections::HashSet::new();
    for sym in &symbols.symbols {
        declared.insert(sym.name.to_ascii_lowercase());
    }
    for kw in KEYWORDS {
        declared.insert(kw.to_ascii_lowercase());
    }
    for f in BUILTIN_FUNCTIONS {
        declared.insert(f.to_ascii_lowercase());
    }
    for t in BUILTIN_TYPES {
        declared.insert(t.to_ascii_lowercase());
    }

    let tokens = lexer::lex(source);
    let mut diagnostics = Vec::new();

    for spanned in &tokens {
        if spanned.token != Token::Identifier {
            continue;
        }
        let offset = spanned.span.start;

        if in_any_span(offset, &skip_spans) || in_any_span(offset, &signature_spans) {
            continue;
        }

        if declared.contains(&spanned.text.to_ascii_lowercase()) {
            continue;
        }

        let range = text_range_to_lsp_range(
            source,
            TextRange::new(spanned.span.start, spanned.span.end),
        );
        diagnostics.push(Diagnostic {
            range,
            severity: Some(DiagnosticSeverity::WARNING),
            source: Some("verde-lsp".to_string()),
            message: format!(
                "Variable '{}' is not declared (Option Explicit)",
                spanned.text
            ),
            ..Default::default()
        });
    }

    diagnostics
}

fn in_any_span(offset: usize, spans: &[TextRange]) -> bool {
    spans
        .iter()
        .any(|s| offset >= s.start as usize && offset < s.end as usize)
}
