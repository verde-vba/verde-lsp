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
/// For each procedure in the module, re-lex the procedure's body (scoped via
/// `ProcedureNode.body_range`) and walk tokens with a small state machine,
/// emitting a Warning for any identifier that is not:
/// - declared at module level (symbol table)
/// - a VBA keyword
/// - a builtin function or type
/// - a VBA-in-Excel global (Application members)
/// - a procedure parameter of the enclosing procedure
/// - a local declared via `Dim`/`Static`/`Const`/`ReDim` earlier in the body
///
/// Because scanning is scoped to `body_range`, module-level declaration sites
/// and procedure signatures are out of scope by construction — no heuristic
/// span-skipping is required.
///
/// Known deferred edge cases (Phase 2 happy path):
/// - `With` blocks, `For i = 1 To 10` loop variables, and assignment lhs vs
///   rhs are not yet distinguished — covered only by the existing state
///   machine's coarse handling.
/// - Nested procedures are not a concern (VBA disallows them).
pub fn check_option_explicit(
    ast: &Ast,
    source: &str,
    symbols: &SymbolTable,
) -> Vec<Diagnostic> {
    // Build the "declared or builtin" lowercase set. This set is
    // module-global: every procedure body scan sees the same set plus its
    // own procedure-scoped locals and parameters.
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
    // Excel `Application` members are VBA globals in the Excel host
    // (e.g. `ActiveWorkbook`, `Range`, `Cells`, `Worksheets`), so they
    // count as declared under Option Explicit.
    for name in crate::excel_model::types::application_globals() {
        declared.insert(name.to_ascii_lowercase());
    }

    let mut diagnostics = Vec::new();

    for (_, node) in ast.nodes.iter() {
        let proc = match node {
            AstNode::Procedure(p) => p,
            _ => continue,
        };

        let body_start = proc.body_range.start as usize;
        let body_end = proc.body_range.end as usize;
        if body_start >= body_end {
            continue;
        }
        let body_source = &source[body_start..body_end];

        // Seed per-procedure local scope with parameter names.
        let mut local_declared: std::collections::HashSet<String> =
            std::collections::HashSet::new();
        for param_id in &proc.params {
            if let AstNode::Parameter(param) = &ast.nodes[*param_id] {
                local_declared.insert(param.name.to_ascii_lowercase());
            }
        }

        scan_body(
            body_source,
            body_start,
            source,
            &declared,
            &mut local_declared,
            &mut diagnostics,
        );
    }

    diagnostics
}

/// Re-lex a procedure body slice and emit Option Explicit warnings for any
/// undeclared identifier references. Token spans from the body's lexer are
/// relative to the slice; we rebase them to absolute source offsets via
/// `body_start_abs` before producing LSP ranges.
fn scan_body(
    body_source: &str,
    body_start_abs: usize,
    source: &str,
    declared: &std::collections::HashSet<String>,
    local_declared: &mut std::collections::HashSet<String>,
    diagnostics: &mut Vec<Diagnostic>,
) {
    let tokens = lexer::lex(body_source);

    // Token-context state machine (per procedure):
    // - `prev_token` is the previous non-trivia token (trivia = Comment).
    //   Used to detect `.Ident` (member access RHS) and `As Ident` (type pos).
    // - `in_decl_list` is true inside a Dim/Static/Private/Public/Const/ReDim
    //   list, until the statement ends (Newline/Colon) or we hit `As` (after
    //   which the next identifier is a type).
    // - Declared local names from the lhs of Dim-style decls are accumulated
    //   into `local_declared` for the remainder of the scan.
    let mut prev_token: Option<Token> = None;
    let mut in_decl_list = false;
    let mut expecting_type = false;

    for spanned in &tokens {
        match spanned.token {
            Token::Dim | Token::Static | Token::Const | Token::ReDim => {
                in_decl_list = true;
                expecting_type = false;
            }
            Token::Private | Token::Public => {
                in_decl_list = true;
                expecting_type = false;
            }
            Token::As => {
                expecting_type = true;
            }
            Token::Comma => {
                if in_decl_list {
                    expecting_type = false;
                }
            }
            Token::Newline | Token::Colon => {
                in_decl_list = false;
                expecting_type = false;
            }
            _ => {}
        }

        if spanned.token != Token::Identifier {
            if !matches!(spanned.token, Token::Comment) {
                prev_token = Some(spanned.token.clone());
            }
            continue;
        }
        let lower = spanned.text.to_ascii_lowercase();
        let after_dot = matches!(prev_token, Some(Token::Dot));
        let after_as = matches!(prev_token, Some(Token::As));

        if in_decl_list && !expecting_type && !after_dot {
            local_declared.insert(lower.clone());
            prev_token = Some(Token::Identifier);
            continue;
        }

        if after_dot {
            prev_token = Some(Token::Identifier);
            continue;
        }

        if after_as {
            expecting_type = false;
            prev_token = Some(Token::Identifier);
            continue;
        }

        prev_token = Some(Token::Identifier);

        if declared.contains(&lower) || local_declared.contains(&lower) {
            continue;
        }

        // Rebase the body-relative token span to absolute source offsets.
        let abs_start = body_start_abs + spanned.span.start;
        let abs_end = body_start_abs + spanned.span.end;
        let range = text_range_to_lsp_range(source, TextRange::new(abs_start, abs_end));
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
}
