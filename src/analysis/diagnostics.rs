use tower_lsp::lsp_types::*;

use super::resolve::text_range_to_lsp_range;
use super::symbols::SymbolTable;
use crate::parser::ast::{Ast, AstNode, ProcedureNode, StatementNode, TextRange};
use crate::parser::lexer::Token;
use crate::parser::ParseResult;
use crate::vba_builtins::{BUILTIN_FUNCTIONS, BUILTIN_TYPES, KEYWORDS};

pub fn compute(parse_result: &ParseResult, symbols: &SymbolTable, source: &str) -> Vec<Diagnostic> {
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
/// For each procedure in the module, walk the procedure body via its parsed
/// AST statements and emit a Warning for any identifier reference that is not:
/// - declared at module level (symbol table)
/// - a VBA keyword
/// - a builtin function or type
/// - a VBA-in-Excel global (Application members)
/// - a procedure parameter of the enclosing procedure
/// - a local declared via `Dim`/`Static`/`Const`/`ReDim` earlier in the body
///
/// Walking the AST (rather than re-lexing the body) lets us recover local
/// declarations directly from `LocalDeclarationNode.names`, so the remaining
/// in-statement state machine only needs to skip identifiers in type position
/// (right after `As`) and member-access RHS (right after `.`).
///
/// Known deferred edge cases (Phase 2 happy path):
/// - `With` blocks, `For i = 1 To 10` loop variables, and assignment lhs vs
///   rhs are not yet distinguished — covered only by the coarse handling in
///   the expression-statement walker.
/// - Nested procedures are not a concern (VBA disallows them).
pub fn check_option_explicit(ast: &Ast, source: &str, symbols: &SymbolTable) -> Vec<Diagnostic> {
    let declared = collect_module_declared(symbols);
    let mut diagnostics = Vec::new();

    for (_, node) in ast.nodes.iter() {
        let proc = match node {
            AstNode::Procedure(p) => p,
            _ => continue,
        };
        scan_procedure(proc, ast, source, &declared, &mut diagnostics);
    }

    diagnostics
}

/// Build the module-global "declared or builtin" lowercase set. Every
/// procedure body scan sees the same set plus its own procedure-scoped locals
/// and parameters.
fn collect_module_declared(symbols: &SymbolTable) -> std::collections::HashSet<String> {
    let mut declared: std::collections::HashSet<String> = std::collections::HashSet::new();
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
    declared
}

/// Walk one procedure body, accumulating local declarations and emitting
/// undeclared-identifier warnings for expression statements.
fn scan_procedure(
    proc: &ProcedureNode,
    ast: &Ast,
    source: &str,
    declared: &std::collections::HashSet<String>,
    diagnostics: &mut Vec<Diagnostic>,
) {
    // Seed per-procedure local scope with parameter names.
    let mut local_declared: std::collections::HashSet<String> = std::collections::HashSet::new();
    for param_id in &proc.params {
        if let AstNode::Parameter(param) = &ast.nodes[*param_id] {
            local_declared.insert(param.name.to_ascii_lowercase());
        }
    }

    for stmt_id in &proc.body {
        let stmt = match &ast.nodes[*stmt_id] {
            AstNode::Statement(s) => s,
            _ => continue,
        };
        match stmt {
            StatementNode::LocalDeclaration(decl) => {
                for (name, _, _) in &decl.names {
                    local_declared.insert(name.to_ascii_lowercase());
                }
            }
            StatementNode::Expression(expr) => {
                scan_expression_tokens(
                    &expr.tokens,
                    source,
                    declared,
                    &local_declared,
                    diagnostics,
                );
            }
            StatementNode::If(node) => {
                scan_expression_tokens(
                    &node.tokens,
                    source,
                    declared,
                    &local_declared,
                    diagnostics,
                );
            }
            StatementNode::For(node) => {
                scan_expression_tokens(
                    &node.tokens,
                    source,
                    declared,
                    &local_declared,
                    diagnostics,
                );
            }
            StatementNode::With(node) => {
                scan_expression_tokens(
                    &node.tokens,
                    source,
                    declared,
                    &local_declared,
                    diagnostics,
                );
            }
            StatementNode::Select(node) => {
                scan_expression_tokens(
                    &node.tokens,
                    source,
                    declared,
                    &local_declared,
                    diagnostics,
                );
            }
            StatementNode::Call(node) => {
                scan_expression_tokens(
                    &node.tokens,
                    source,
                    declared,
                    &local_declared,
                    diagnostics,
                );
            }
            StatementNode::Set(node) => {
                scan_expression_tokens(
                    &node.tokens,
                    source,
                    declared,
                    &local_declared,
                    diagnostics,
                );
            }
        }
    }
}

/// Walk tokens of a single expression statement and emit Option Explicit
/// warnings for identifiers that are neither declared nor occupy a
/// non-reference position (type after `As`, member name after `.`).
fn scan_expression_tokens(
    tokens: &[crate::parser::lexer::SpannedToken],
    source: &str,
    declared: &std::collections::HashSet<String>,
    local_declared: &std::collections::HashSet<String>,
    diagnostics: &mut Vec<Diagnostic>,
) {
    // `prev_token` is the previous non-trivia token (trivia = Comment).
    // Used to detect `.Ident` (member access RHS) and `As Ident` (type pos).
    let mut prev_token: Option<Token> = None;

    for spanned in tokens {
        if spanned.token != Token::Identifier {
            if !matches!(spanned.token, Token::Comment) {
                prev_token = Some(spanned.token.clone());
            }
            continue;
        }

        let after_dot = matches!(prev_token, Some(Token::Dot));
        let after_as = matches!(prev_token, Some(Token::As));
        prev_token = Some(Token::Identifier);

        if after_dot || after_as {
            continue;
        }

        let lower = spanned.text.to_ascii_lowercase();
        if declared.contains(&lower) || local_declared.contains(&lower) {
            continue;
        }

        let range =
            text_range_to_lsp_range(source, TextRange::new(spanned.span.start, spanned.span.end));
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
