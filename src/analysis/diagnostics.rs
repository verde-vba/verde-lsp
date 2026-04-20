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

    // Parameter names harvested from ProcedureNode.params (populated by the
    // parser). Used to seed `local_declared` so parameter references inside
    // the body aren't flagged under Option Explicit.
    let mut parameter_names: Vec<String> = Vec::new();

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
                for param_id in &p.params {
                    if let AstNode::Parameter(param) = &ast.nodes[*param_id] {
                        parameter_names.push(param.name.to_ascii_lowercase());
                    }
                }
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
    // Excel `Application` members are VBA globals in the Excel host
    // (e.g. `ActiveWorkbook`, `Range`, `Cells`, `Worksheets`), so they
    // count as declared under Option Explicit.
    for name in crate::excel_model::types::application_globals() {
        declared.insert(name.to_ascii_lowercase());
    }

    let tokens = lexer::lex(source);
    let mut diagnostics = Vec::new();

    // Token-context state machine:
    // - `prev_token` is the previous non-trivia token (trivia = Newline/Comment).
    //   Used to detect `.Ident` (member access RHS) and `As Ident` (type position).
    // - `in_decl_list` is true when we're inside a Dim/Static/Private/Public/
    //   Const/ReDim list, until the statement ends (Newline) or we hit `As`
    //   (after which the next identifier is a type, not a declared name).
    // - Declared local names from Dim/etc. LHS are accumulated into
    //   `local_declared` and added to the "is declared" set for the remainder
    //   of the scan.
    let mut prev_token: Option<Token> = None;
    let mut in_decl_list = false;
    // After `As` we expect a type identifier next, even if we were inside a
    // Dim list; after that identifier, control returns to the Dim list only if
    // a comma follows.
    let mut expecting_type = false;
    let mut local_declared: std::collections::HashSet<String> =
        std::collections::HashSet::new();
    // Seed with procedure parameter names so their use inside the body
    // doesn't trigger undeclared warnings.
    for name in &parameter_names {
        local_declared.insert(name.clone());
    }

    for spanned in &tokens {
        // Update decl-list state BEFORE handling identifiers so the keyword
        // that starts a decl list flips the flag in time for the next ident.
        match spanned.token {
            Token::Dim | Token::Static | Token::Const | Token::ReDim => {
                in_decl_list = true;
                expecting_type = false;
            }
            Token::Private | Token::Public => {
                // These can start decl lists at module level or qualify
                // procedures. Treat conservatively: enable decl list; it gets
                // cleared on Newline if it was a procedure decl (Sub/Function
                // lives on the signature line, whose identifiers are already
                // skipped via signature_spans).
                in_decl_list = true;
                expecting_type = false;
            }
            Token::As => {
                expecting_type = true;
            }
            Token::Comma => {
                // Inside a Dim list, a comma means another declared name
                // follows. Clear "expecting_type" so we treat the next ident
                // as a declaration name again.
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
            // Track previous non-trivia token for lookback on the next ident.
            if !matches!(spanned.token, Token::Comment) {
                prev_token = Some(spanned.token.clone());
            }
            continue;
        }
        let offset = spanned.span.start;
        let lower = spanned.text.to_ascii_lowercase();

        // Context-based suppression for this identifier.
        let after_dot = matches!(prev_token, Some(Token::Dot));
        let after_as = matches!(prev_token, Some(Token::As));

        // If we're inside a Dim-style decl list and NOT in a type position,
        // this identifier is a declared name: record it and skip warning.
        if in_decl_list && !expecting_type && !after_dot {
            local_declared.insert(lower.clone());
            prev_token = Some(Token::Identifier);
            continue;
        }

        // After `.`, skip entirely — it's a member access RHS.
        if after_dot {
            prev_token = Some(Token::Identifier);
            continue;
        }

        // After `As`, this identifier is a type reference. Skip it
        // (covers user-defined types, which the symbol table may or may not
        // contain; being conservative here).
        if after_as {
            // Clear expecting_type once the type ident is consumed.
            expecting_type = false;
            prev_token = Some(Token::Identifier);
            continue;
        }

        prev_token = Some(Token::Identifier);

        if in_any_span(offset, &skip_spans) || in_any_span(offset, &signature_spans) {
            continue;
        }

        if declared.contains(&lower) || local_declared.contains(&lower) {
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
