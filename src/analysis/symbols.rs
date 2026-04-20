use smol_str::SmolStr;

use crate::parser::ast::*;

#[derive(Debug, Clone)]
pub struct SymbolTable {
    pub symbols: Vec<Symbol>,
    pub option_explicit: bool,
    /// Full byte-range of each procedure in source order, used for cursor-scope detection.
    pub proc_ranges: Vec<(SmolStr, TextRange)>,
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
                        members: Vec::new(),
                    },
                    proc_scope: None,
                });
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

    SymbolTable {
        symbols,
        option_explicit: ast.option_explicit,
        proc_ranges,
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::parser::parse;

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
}
