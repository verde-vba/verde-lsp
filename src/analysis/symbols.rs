use smol_str::SmolStr;

use crate::parser::ast::*;

#[derive(Debug, Clone)]
pub struct SymbolTable {
    pub symbols: Vec<Symbol>,
    pub option_explicit: bool,
    /// Full byte-range of each procedure in source order, used for cursor-scope detection.
    pub proc_ranges: Vec<(SmolStr, TextRange)>,
    /// Block ranges (With/If/For/Select/Do/While) within procedures,
    /// computed by scanning the token stream for open/close pairs.
    pub block_ranges: Vec<BlockRange>,
    /// Interface names declared via `Implements IFoo` at module level.
    pub implements: Vec<SmolStr>,
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum BlockKind {
    With,
    If,
    For,
    Select,
    Do,
    While,
}

#[derive(Debug, Clone)]
pub struct BlockRange {
    pub kind: BlockKind,
    pub start: u32,
    pub end: u32,
    pub proc_name: SmolStr,
    /// Extra data: for With blocks, the target expression (e.g. variable name).
    pub data: Option<SmolStr>,
}

#[derive(Debug, Clone)]
pub struct Symbol {
    pub name: SmolStr,
    pub kind: SymbolKind,
    pub type_name: Option<SmolStr>,
    pub visibility: Visibility,
    pub span: TextRange,
    pub detail: SymbolDetail,
    /// None = module-level (visible everywhere); Some(name) = visible only inside that procedure.
    pub proc_scope: Option<SmolStr>,
}

#[derive(Debug, Clone)]
pub enum SymbolKind {
    Procedure,
    Function,
    Property,
    Variable,
    Constant,
    Parameter,
    TypeDef,
    EnumDef,
    EnumMember,
    UdtMember,
}

impl SymbolKind {
    pub fn to_lsp_symbol_kind(&self) -> tower_lsp::lsp_types::SymbolKind {
        match self {
            SymbolKind::Procedure => tower_lsp::lsp_types::SymbolKind::FUNCTION,
            SymbolKind::Function => tower_lsp::lsp_types::SymbolKind::FUNCTION,
            SymbolKind::Property => tower_lsp::lsp_types::SymbolKind::PROPERTY,
            SymbolKind::Variable => tower_lsp::lsp_types::SymbolKind::VARIABLE,
            SymbolKind::Constant => tower_lsp::lsp_types::SymbolKind::CONSTANT,
            SymbolKind::Parameter => tower_lsp::lsp_types::SymbolKind::VARIABLE,
            SymbolKind::TypeDef => tower_lsp::lsp_types::SymbolKind::STRUCT,
            SymbolKind::EnumDef => tower_lsp::lsp_types::SymbolKind::ENUM,
            SymbolKind::EnumMember => tower_lsp::lsp_types::SymbolKind::ENUM_MEMBER,
            SymbolKind::UdtMember => tower_lsp::lsp_types::SymbolKind::FIELD,
        }
    }
}

#[derive(Debug, Clone)]
pub enum SymbolDetail {
    Procedure {
        kind: ProcedureKind,
        params: Vec<ParameterInfo>,
        return_type: Option<SmolStr>,
    },
    Parameter {
        type_name: Option<SmolStr>,
        passing: ParameterPassing,
        is_optional: bool,
    },
    Variable {
        is_static: bool,
    },
    TypeDef {
        members: Vec<(SmolStr, Option<SmolStr>)>,
    },
    EnumDef {
        members: Vec<(SmolStr, Option<i64>)>,
    },
    EnumMember {
        parent_enum: SmolStr,
        value: i64,
    },
    UdtMember {
        /// Name of the TypeDef this member belongs to.
        parent_type: SmolStr,
        type_name: SmolStr,
    },
}

#[derive(Debug, Clone)]
pub struct ParameterInfo {
    pub name: SmolStr,
    pub type_name: Option<SmolStr>,
    pub passing: ParameterPassing,
    pub is_optional: bool,
}

pub fn build_symbol_table(ast: &Ast) -> SymbolTable {
    let mut symbols = Vec::new();
    let mut proc_ranges = Vec::new();
    let mut block_ranges = Vec::new();

    for &node_id in &ast.root {
        match &ast.nodes[node_id] {
            AstNode::Procedure(proc) => {
                proc_ranges.push((proc.name.clone(), proc.span));

                let kind = match proc.kind {
                    ProcedureKind::Sub => SymbolKind::Procedure,
                    ProcedureKind::Function => SymbolKind::Function,
                    _ => SymbolKind::Property,
                };
                symbols.push(Symbol {
                    name: proc.name.clone(),
                    kind,
                    type_name: proc.return_type.clone(),
                    visibility: proc.visibility.clone(),
                    span: proc.name_span,
                    detail: SymbolDetail::Procedure {
                        kind: proc.kind.clone(),
                        params: proc
                            .params
                            .iter()
                            .filter_map(|&id| {
                                if let AstNode::Parameter(p) = &ast.nodes[id] {
                                    Some(ParameterInfo {
                                        name: p.name.clone(),
                                        type_name: p.type_name.clone(),
                                        passing: p.passing.clone(),
                                        is_optional: p.is_optional,
                                    })
                                } else {
                                    None
                                }
                            })
                            .collect(),
                        return_type: proc.return_type.clone(),
                    },
                    proc_scope: None,
                });
                for &param_id in &proc.params {
                    if let AstNode::Parameter(p) = &ast.nodes[param_id] {
                        symbols.push(Symbol {
                            name: p.name.clone(),
                            kind: SymbolKind::Parameter,
                            type_name: p.type_name.clone(),
                            visibility: Visibility::Private,
                            span: p.span,
                            detail: SymbolDetail::Parameter {
                                type_name: p.type_name.clone(),
                                passing: p.passing.clone(),
                                is_optional: p.is_optional,
                            },
                            proc_scope: Some(proc.name.clone()),
                        });
                    }
                }
                for &stmt_id in &proc.body {
                    if let AstNode::Statement(StatementNode::LocalDeclaration(decl)) =
                        &ast.nodes[stmt_id]
                    {
                        let is_const = decl.kind == DeclKind::Const;
                        let is_static = decl.kind == DeclKind::Static;
                        for (name, type_name, name_span) in &decl.names {
                            symbols.push(Symbol {
                                name: name.clone(),
                                kind: if is_const {
                                    SymbolKind::Constant
                                } else {
                                    SymbolKind::Variable
                                },
                                type_name: type_name.clone(),
                                visibility: Visibility::Private,
                                span: *name_span,
                                detail: SymbolDetail::Variable { is_static },
                                proc_scope: Some(proc.name.clone()),
                            });
                        }
                    }
                }
            }
            AstNode::Variable(var) => {
                let kind = if var.is_const {
                    SymbolKind::Constant
                } else {
                    SymbolKind::Variable
                };
                symbols.push(Symbol {
                    name: var.name.clone(),
                    kind,
                    type_name: var.type_name.clone(),
                    visibility: var.visibility.clone(),
                    span: var.span,
                    detail: SymbolDetail::Variable {
                        is_static: var.is_static,
                    },
                    proc_scope: None,
                });
            }
            AstNode::TypeDef(td) => {
                symbols.push(Symbol {
                    name: td.name.clone(),
                    kind: SymbolKind::TypeDef,
                    type_name: None,
                    visibility: td.visibility.clone(),
                    span: td.span,
                    detail: SymbolDetail::TypeDef {
                        members: td
                            .members
                            .iter()
                            .map(|(n, t, _)| (n.clone(), t.clone()))
                            .collect(),
                    },
                    proc_scope: None,
                });
                for (member_name, member_type, member_span) in &td.members {
                    symbols.push(Symbol {
                        name: member_name.clone(),
                        kind: SymbolKind::UdtMember,
                        type_name: member_type.clone(),
                        visibility: td.visibility.clone(),
                        span: *member_span,
                        detail: SymbolDetail::UdtMember {
                            parent_type: td.name.clone(),
                            type_name: member_type
                                .clone()
                                .unwrap_or_else(|| SmolStr::new("Variant")),
                        },
                        proc_scope: None,
                    });
                }
            }
            AstNode::EnumDef(ed) => {
                symbols.push(Symbol {
                    name: ed.name.clone(),
                    kind: SymbolKind::EnumDef,
                    type_name: None,
                    visibility: ed.visibility.clone(),
                    span: ed.span,
                    detail: SymbolDetail::EnumDef {
                        members: ed.members.clone(),
                    },
                    proc_scope: None,
                });
                let mut next_value: i64 = 0;
                for (member_name, value) in &ed.members {
                    let resolved = match value {
                        Some(v) => {
                            next_value = v + 1;
                            *v
                        }
                        None => {
                            let implicit = next_value;
                            next_value += 1;
                            implicit
                        }
                    };
                    symbols.push(Symbol {
                        name: member_name.clone(),
                        kind: SymbolKind::EnumMember,
                        type_name: Some(ed.name.clone()),
                        visibility: ed.visibility.clone(),
                        span: ed.span,
                        detail: SymbolDetail::EnumMember {
                            parent_enum: ed.name.clone(),
                            value: resolved,
                        },
                        proc_scope: None,
                    });
                }
            }
            _ => {}
        }
    }

    // Compute block ranges within each procedure's body.
    for &node_id in &ast.root {
        if let AstNode::Procedure(proc) = &ast.nodes[node_id] {
            compute_block_ranges(proc, ast, &mut block_ranges);
        }
    }

    SymbolTable {
        symbols,
        option_explicit: ast.option_explicit,
        proc_ranges,
        block_ranges,
        implements: ast.implements.clone(),
    }
}

/// Scan a procedure's body statements to find block open/close pairs
/// and record their ranges.
fn compute_block_ranges(proc: &ProcedureNode, ast: &Ast, out: &mut Vec<BlockRange>) {
    use crate::parser::lexer::Token;

    // Build a flat list of (token, span_start, span_end) from all body statements.
    struct OpenBlock {
        kind: BlockKind,
        start: u32,
        data: Option<SmolStr>,
    }

    let mut stack: Vec<OpenBlock> = Vec::new();

    for &stmt_id in &proc.body {
        let stmt = match &ast.nodes[stmt_id] {
            AstNode::Statement(s) => s,
            _ => continue,
        };

        match stmt {
            StatementNode::With(w) => {
                // Extract the With target (first Identifier after With keyword).
                let data = w
                    .tokens
                    .iter()
                    .find(|t| t.token == Token::Identifier)
                    .map(|t| t.text.clone());
                stack.push(OpenBlock {
                    kind: BlockKind::With,
                    start: w.span.start,
                    data,
                });
            }
            StatementNode::If(i) => {
                stack.push(OpenBlock {
                    kind: BlockKind::If,
                    start: i.span.start,
                    data: None,
                });
            }
            StatementNode::For(f) => {
                // Extract loop variable (first Identifier after For).
                let data = f
                    .tokens
                    .iter()
                    .find(|t| t.token == Token::Identifier)
                    .map(|t| t.text.clone());
                stack.push(OpenBlock {
                    kind: BlockKind::For,
                    start: f.span.start,
                    data,
                });
            }
            StatementNode::Select(s) => {
                stack.push(OpenBlock {
                    kind: BlockKind::Select,
                    start: s.span.start,
                    data: None,
                });
            }
            StatementNode::Do(d) => {
                stack.push(OpenBlock {
                    kind: BlockKind::Do,
                    start: d.span.start,
                    data: None,
                });
            }
            StatementNode::While(wh) => {
                stack.push(OpenBlock {
                    kind: BlockKind::While,
                    start: wh.span.start,
                    data: None,
                });
            }
            StatementNode::Expression(expr) => {
                // Check for closing tokens (EndWith, EndIf, Next, EndSelect, Loop, Wend)
                if let Some(first) = expr.tokens.first() {
                    let close_kind = match first.token {
                        Token::EndWith => Some(BlockKind::With),
                        Token::EndIf => Some(BlockKind::If),
                        Token::Next => Some(BlockKind::For),
                        Token::EndSelect => Some(BlockKind::Select),
                        Token::Loop => Some(BlockKind::Do),
                        Token::Wend => Some(BlockKind::While),
                        _ => None,
                    };
                    if let Some(kind) = close_kind {
                        // Pop the matching open block from the stack.
                        if let Some(pos) = stack.iter().rposition(|b| b.kind == kind) {
                            let open = stack.remove(pos);
                            out.push(BlockRange {
                                kind: open.kind,
                                start: open.start,
                                end: expr.span.end,
                                proc_name: proc.name.clone(),
                                data: open.data,
                            });
                        }
                    }
                }
            }
            _ => {}
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::parser::parse;

    #[test]
    fn type_block_registers_udt_symbol() {
        let result = parse("Type PersonType\n    Name As String\n    Age As Long\nEnd Type\n");
        let symbols = build_symbol_table(&result.ast);
        let td = symbols
            .symbols
            .iter()
            .find(|s| s.name.as_str() == "PersonType");
        assert!(td.is_some(), "expected TypeDef symbol 'PersonType'");
        let td = td.unwrap();
        assert!(matches!(td.kind, SymbolKind::TypeDef));
        match &td.detail {
            SymbolDetail::TypeDef { members } => {
                assert_eq!(members.len(), 2, "expected 2 members in TypeDef detail");
            }
            other => panic!("expected TypeDef detail, got {:?}", other),
        }
    }

    #[test]
    fn udt_members_accessible_by_name() {
        let result = parse("Type PersonType\n    Name As String\n    Age As Long\nEnd Type\n");
        let symbols = build_symbol_table(&result.ast);
        let member = symbols
            .symbols
            .iter()
            .find(|s| s.name.as_str() == "Name" && matches!(s.kind, SymbolKind::UdtMember));
        assert!(member.is_some(), "expected UdtMember symbol 'Name'");
        let member = member.unwrap();
        match &member.detail {
            SymbolDetail::UdtMember { type_name, .. } => {
                assert_eq!(type_name.as_str(), "String");
            }
            other => panic!("expected UdtMember detail, got {:?}", other),
        }
    }

    #[test]
    fn build_symbol_table_registers_local_dim_variable() {
        let result = parse("Sub Foo()\n    Dim x As String\nEnd Sub\n");
        let symbols = build_symbol_table(&result.ast);
        let sym = symbols.symbols.iter().find(|s| s.name.as_str() == "x");
        assert!(sym.is_some(), "expected local 'x' in symbol table");
        let sym = sym.unwrap();
        assert_eq!(sym.type_name.as_deref(), Some("String"));
        assert!(matches!(sym.kind, SymbolKind::Variable));
    }

    #[test]
    fn parameter_symbol_has_parameter_detail() {
        let result = parse("Sub Foo(x As Integer)\nEnd Sub\n");
        let symbols = build_symbol_table(&result.ast);
        let sym = symbols
            .symbols
            .iter()
            .find(|s| s.name.as_str() == "x" && matches!(s.kind, SymbolKind::Parameter))
            .expect("expected parameter 'x' in symbol table");
        assert!(
            matches!(sym.detail, SymbolDetail::Parameter { .. }),
            "expected SymbolDetail::Parameter, got {:?}",
            sym.detail
        );
    }

    #[test]
    fn enum_member_symbol_has_enum_member_detail() {
        let result = parse("Enum Color\n    Red\n    Green\nEnd Enum\n");
        let symbols = build_symbol_table(&result.ast);
        let sym = symbols
            .symbols
            .iter()
            .find(|s| s.name.as_str() == "Red" && matches!(s.kind, SymbolKind::EnumMember))
            .expect("expected enum member 'Red' in symbol table");
        match &sym.detail {
            SymbolDetail::EnumMember { parent_enum, .. } => {
                assert_eq!(parent_enum.as_str(), "Color");
            }
            other => panic!("expected SymbolDetail::EnumMember, got {:?}", other),
        }
    }

    #[test]
    fn enum_member_with_explicit_value_captures_integer_literal() {
        let result = parse("Enum Color\n    Red = 1\n    Green = 2\nEnd Enum\n");
        let symbols = build_symbol_table(&result.ast);

        let red = symbols
            .symbols
            .iter()
            .find(|s| s.name.as_str() == "Red" && matches!(s.kind, SymbolKind::EnumMember))
            .expect("expected enum member 'Red' in symbol table");
        match &red.detail {
            SymbolDetail::EnumMember { value, .. } => {
                assert_eq!(*value, 1, "expected Red = 1, got {:?}", value);
            }
            other => panic!("expected SymbolDetail::EnumMember for Red, got {:?}", other),
        }

        let green = symbols
            .symbols
            .iter()
            .find(|s| s.name.as_str() == "Green" && matches!(s.kind, SymbolKind::EnumMember))
            .expect("expected enum member 'Green' in symbol table");
        match &green.detail {
            SymbolDetail::EnumMember { value, .. } => {
                assert_eq!(*value, 2, "expected Green = 2, got {:?}", value);
            }
            other => panic!(
                "expected SymbolDetail::EnumMember for Green, got {:?}",
                other
            ),
        }
    }

    #[test]
    fn enum_member_with_negative_value_captures_negative_integer() {
        let result = parse("Enum Sign\n    Minus = -1\nEnd Enum\n");
        let symbols = build_symbol_table(&result.ast);
        let sym = symbols
            .symbols
            .iter()
            .find(|s| s.name.as_str() == "Minus" && matches!(s.kind, SymbolKind::EnumMember))
            .expect("expected enum member 'Minus' in symbol table");
        match &sym.detail {
            SymbolDetail::EnumMember { value, .. } => {
                assert_eq!(*value, -1, "expected Minus = -1, got {:?}", value);
            }
            other => panic!(
                "expected SymbolDetail::EnumMember for Minus, got {:?}",
                other
            ),
        }
    }

    #[test]
    fn enum_member_with_hex_literal_captures_integer_value() {
        let result = parse("Enum Flags\n    Ten = &H10\nEnd Enum\n");
        let symbols = build_symbol_table(&result.ast);
        let sym = symbols
            .symbols
            .iter()
            .find(|s| s.name.as_str() == "Ten" && matches!(s.kind, SymbolKind::EnumMember))
            .expect("expected enum member 'Ten' in symbol table");
        match &sym.detail {
            SymbolDetail::EnumMember { value, .. } => {
                assert_eq!(*value, 16, "expected Ten = &H10 (= 16), got {:?}", value);
            }
            other => panic!("expected SymbolDetail::EnumMember for Ten, got {:?}", other),
        }
    }

    #[test]
    fn enum_implicit_value_follows_previous_member() {
        let result = parse("Enum Foo\n    A\n    B = 10\n    C\nEnd Enum\n");
        let symbols = build_symbol_table(&result.ast);
        for (name, expected) in [("A", 0i64), ("B", 10i64), ("C", 11i64)] {
            let sym = symbols
                .symbols
                .iter()
                .find(|s| s.name.as_str() == name && matches!(s.kind, SymbolKind::EnumMember))
                .unwrap_or_else(|| panic!("expected enum member '{}' in symbol table", name));
            match &sym.detail {
                SymbolDetail::EnumMember { value, .. } => {
                    assert_eq!(
                        *value, expected,
                        "expected {} = {:?}, got {:?}",
                        name, expected, value
                    );
                }
                other => panic!(
                    "expected SymbolDetail::EnumMember for {}, got {:?}",
                    name, other
                ),
            }
        }
    }

    // ── PLAN-12: Block range tests ────────────────────────────────────

    #[test]
    fn with_block_range_is_computed() {
        let source = "Sub A()\n    With rng\n        .Value = 1\n    End With\nEnd Sub\n";
        let result = parse(source);
        let symbols = build_symbol_table(&result.ast);
        assert_eq!(symbols.block_ranges.len(), 1, "expected 1 block range");
        let br = &symbols.block_ranges[0];
        assert_eq!(br.kind, BlockKind::With);
        assert_eq!(br.proc_name.as_str(), "A");
        assert_eq!(br.data.as_deref(), Some("rng"));
    }

    #[test]
    fn nested_blocks_are_computed() {
        let source =
            "Sub A()\n    With rng\n        If True Then\n        End If\n    End With\nEnd Sub\n";
        let result = parse(source);
        let symbols = build_symbol_table(&result.ast);
        assert_eq!(
            symbols.block_ranges.len(),
            2,
            "expected 2 block ranges (With + If)"
        );
        assert!(symbols
            .block_ranges
            .iter()
            .any(|b| b.kind == BlockKind::With));
        assert!(symbols.block_ranges.iter().any(|b| b.kind == BlockKind::If));
    }

    #[test]
    fn for_block_range_captures_loop_variable() {
        let source = "Sub A()\n    For i = 1 To 10\n    Next i\nEnd Sub\n";
        let result = parse(source);
        let symbols = build_symbol_table(&result.ast);
        assert_eq!(symbols.block_ranges.len(), 1, "expected 1 block range");
        let br = &symbols.block_ranges[0];
        assert_eq!(br.kind, BlockKind::For);
        assert_eq!(br.data.as_deref(), Some("i"));
    }

    #[test]
    fn incomplete_block_does_not_panic() {
        let source = "Sub A()\n    With rng\n        .Value = 1\nEnd Sub\n";
        let result = parse(source);
        let symbols = build_symbol_table(&result.ast);
        // Incomplete With (no End With) — should not panic, just no block range
        assert!(
            symbols.block_ranges.is_empty(),
            "incomplete block should not produce a range"
        );
    }

    // ── PLAN-17: Implements in symbol table ──────────────────────────

    #[test]
    fn implements_registered_in_symbol_table() {
        let result = parse("Implements IFoo\nSub Bar()\nEnd Sub\n");
        let symbols = build_symbol_table(&result.ast);
        assert_eq!(symbols.implements.len(), 1);
        assert_eq!(symbols.implements[0].as_str(), "IFoo");
    }
}
