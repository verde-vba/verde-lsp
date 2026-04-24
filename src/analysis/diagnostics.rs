use smol_str::SmolStr;
use tower_lsp::lsp_types::*;

use super::resolve::text_range_to_lsp_range;
use super::symbols::SymbolTable;
use crate::parser::ast::{Ast, AstNode, ProcedureNode, StatementNode, TextRange};
use crate::parser::lexer::Token;
use crate::parser::ParseResult;
use crate::vba_builtins::{BUILTIN_FUNCTIONS, BUILTIN_TYPES, EXCEL_CONSTANTS, KEYWORDS, VBA_CONSTANTS};

/// VBA runtime global objects that are always available without declaration.
const VBA_GLOBAL_OBJECTS: &[&str] = &["Debug", "Err"];

pub fn compute(
    parse_result: &ParseResult,
    symbols: &SymbolTable,
    source: &str,
    cross_module_names: &std::collections::HashSet<SmolStr>,
) -> Vec<Diagnostic> {
    let mut diagnostics = Vec::new();

    for error in &parse_result.errors {
        let range =
            text_range_to_lsp_range(source, TextRange::new(error.span.start, error.span.end));
        diagnostics.push(Diagnostic {
            range,
            severity: Some(DiagnosticSeverity::ERROR),
            source: Some("verde-lsp".to_string()),
            message: error.message.clone(),
            ..Default::default()
        });
    }

    if parse_result.ast.option_explicit {
        diagnostics.extend(check_option_explicit(
            &parse_result.ast,
            source,
            symbols,
            cross_module_names,
        ));
    }

    diagnostics.extend(check_unused_variables(&parse_result.ast, source));

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
pub fn check_option_explicit(
    ast: &Ast,
    source: &str,
    symbols: &SymbolTable,
    cross_module_names: &std::collections::HashSet<SmolStr>,
) -> Vec<Diagnostic> {
    let mut declared = collect_module_declared(symbols);
    declared.extend(cross_module_names.iter().cloned());
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
fn collect_module_declared(symbols: &SymbolTable) -> std::collections::HashSet<SmolStr> {
    let mut declared: std::collections::HashSet<SmolStr> = std::collections::HashSet::new();
    for sym in &symbols.symbols {
        declared.insert(SmolStr::new(sym.name.to_ascii_lowercase()));
    }
    for kw in KEYWORDS {
        declared.insert(SmolStr::new(kw.to_ascii_lowercase()));
    }
    for f in BUILTIN_FUNCTIONS {
        declared.insert(SmolStr::new(f.to_ascii_lowercase()));
    }
    for t in BUILTIN_TYPES {
        declared.insert(SmolStr::new(t.to_ascii_lowercase()));
    }
    // VBA runtime global objects (e.g. `Debug.Print`, `Err.Raise`).
    for name in VBA_GLOBAL_OBJECTS {
        declared.insert(SmolStr::new(name.to_ascii_lowercase()));
    }
    for c in VBA_CONSTANTS {
        declared.insert(SmolStr::new(c.to_ascii_lowercase()));
    }
    for c in EXCEL_CONSTANTS {
        declared.insert(SmolStr::new(c.to_ascii_lowercase()));
    }
    // Excel `Application` members are VBA globals in the Excel host
    // (e.g. `ActiveWorkbook`, `Range`, `Cells`, `Worksheets`), so they
    // count as declared under Option Explicit.
    for name in crate::excel_model::types::application_globals() {
        declared.insert(SmolStr::new(name.to_ascii_lowercase()));
    }
    declared
}

/// Walk one procedure body, accumulating local declarations and emitting
/// undeclared-identifier warnings for expression statements.
fn scan_procedure(
    proc: &ProcedureNode,
    ast: &Ast,
    source: &str,
    declared: &std::collections::HashSet<SmolStr>,
    diagnostics: &mut Vec<Diagnostic>,
) {
    // Seed per-procedure local scope with parameter names.
    let mut local_declared: std::collections::HashSet<SmolStr> = std::collections::HashSet::new();
    for param_id in &proc.params {
        if let AstNode::Parameter(param) = &ast.nodes[*param_id] {
            local_declared.insert(SmolStr::new(param.name.to_ascii_lowercase()));
        }
    }

    // Pre-scan: collect GoTo/OnError label targets so that label definition
    // lines (parsed as bare-identifier ExpressionStatements) are not flagged.
    for stmt_id in &proc.body {
        if let AstNode::Statement(stmt) = &ast.nodes[*stmt_id] {
            let tokens = match stmt {
                StatementNode::GoTo(n) => &n.tokens,
                StatementNode::OnError(n) => &n.tokens,
                _ => continue,
            };
            for spanned in tokens {
                if spanned.token == Token::Identifier {
                    local_declared.insert(SmolStr::new(spanned.text.to_ascii_lowercase()));
                }
            }
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
                    local_declared.insert(SmolStr::new(name.to_ascii_lowercase()));
                }
            }
            // GoTo/OnError targets are label names, not variable references.
            StatementNode::GoTo(_) | StatementNode::OnError(_) => {}
            _ => {
                let tokens = stmt.tokens();
                if !tokens.is_empty() {
                    scan_expression_tokens(tokens, source, declared, &local_declared, diagnostics);
                }
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
    declared: &std::collections::HashSet<SmolStr>,
    local_declared: &std::collections::HashSet<SmolStr>,
    diagnostics: &mut Vec<Diagnostic>,
) {
    // `prev_token` is the previous non-trivia token (trivia = Comment).
    // Used to detect `.Ident` (member access RHS) and `As Ident` (type pos).
    let mut prev_token: Option<Token> = None;

    for spanned in tokens {
        if spanned.token != Token::Identifier {
            if !matches!(spanned.token, Token::Comment) {
                prev_token = Some(spanned.token);
            }
            continue;
        }

        let after_dot = matches!(prev_token, Some(Token::Dot));
        let after_as = matches!(prev_token, Some(Token::As));
        prev_token = Some(Token::Identifier);

        if after_dot || after_as {
            continue;
        }

        let lower = SmolStr::new(spanned.text.to_ascii_lowercase());
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

/// Walk every procedure in the AST and emit a WARNING for each local variable
/// that is declared (via `Dim`/`Static`/`Const`) but never referenced in any
/// non-declaration statement within that procedure. Variables whose name starts
/// with `_` are excluded (convention for intentionally unused bindings).
fn check_unused_variables(ast: &Ast, source: &str) -> Vec<Diagnostic> {
    let mut diagnostics = Vec::new();

    for &node_id in &ast.root {
        let proc = match &ast.nodes[node_id] {
            AstNode::Procedure(p) => p,
            _ => continue,
        };

        // Collect locally declared variable names and their spans.
        let mut locals: Vec<(SmolStr, TextRange)> = Vec::new();
        for &stmt_id in &proc.body {
            if let AstNode::Statement(StatementNode::LocalDeclaration(decl)) = &ast.nodes[stmt_id] {
                for (name, _, name_span) in &decl.names {
                    // Skip variables with _ prefix.
                    if name.starts_with('_') {
                        continue;
                    }
                    locals.push((name.clone(), *name_span));
                }
            }
        }

        if locals.is_empty() {
            continue;
        }

        // Collect all identifier texts (lowercased) from non-declaration body statements.
        let mut used_idents: std::collections::HashSet<SmolStr> = std::collections::HashSet::new();
        for &stmt_id in &proc.body {
            let stmt = match &ast.nodes[stmt_id] {
                AstNode::Statement(s) => s,
                _ => continue,
            };
            match stmt {
                StatementNode::LocalDeclaration(_) => continue,
                _ => {
                    for spanned in stmt.tokens() {
                        if spanned.token == Token::Identifier {
                            used_idents.insert(SmolStr::new(spanned.text.to_ascii_lowercase()));
                        }
                    }
                }
            }
        }

        // Emit diagnostics for declared-but-unused locals.
        for (name, span) in &locals {
            let lower_name = SmolStr::new(name.to_ascii_lowercase());
            let is_used = used_idents.contains(&lower_name);
            if !is_used {
                diagnostics.push(Diagnostic {
                    range: text_range_to_lsp_range(source, *span),
                    severity: Some(DiagnosticSeverity::WARNING),
                    source: Some("verde-lsp".to_string()),
                    message: format!("Variable '{}' is declared but never used", name),
                    ..Default::default()
                });
            }
        }
    }

    diagnostics
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn unused_variable_warning() {
        let source = "Sub Foo()\n    Dim x As Long\nEnd Sub\n";
        let result = crate::parser::parse(source);
        let symbols = crate::analysis::symbols::build_symbol_table(&result.ast);
        let diags = compute(
            &result,
            &symbols,
            source,
            &std::collections::HashSet::<SmolStr>::new(),
        );
        assert!(diags
            .iter()
            .any(|d| d.message.contains("x") && d.message.contains("never used")));
    }

    #[test]
    fn used_variable_no_warning() {
        let source = "Sub Foo()\n    Dim x As Long\n    x = 1\nEnd Sub\n";
        let result = crate::parser::parse(source);
        let symbols = crate::analysis::symbols::build_symbol_table(&result.ast);
        let diags = compute(
            &result,
            &symbols,
            source,
            &std::collections::HashSet::<SmolStr>::new(),
        );
        assert!(!diags.iter().any(|d| d.message.contains("never used")));
    }

    #[test]
    fn underscore_prefix_variable_excluded() {
        let source = "Sub Foo()\n    Dim _unused As Long\nEnd Sub\n";
        let result = crate::parser::parse(source);
        let symbols = crate::analysis::symbols::build_symbol_table(&result.ast);
        let diags = compute(
            &result,
            &symbols,
            source,
            &std::collections::HashSet::<SmolStr>::new(),
        );
        assert!(!diags.iter().any(|d| d.message.contains("_unused")));
    }
}
