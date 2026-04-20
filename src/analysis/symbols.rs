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
    Variable {
        is_static: bool,
    },
    TypeDef {
        members: Vec<(SmolStr, Option<SmolStr>)>,
    },
    EnumDef {
        members: Vec<(SmolStr, Option<i64>)>,
    },
    None,
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
                            detail: SymbolDetail::None,
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
                for (member_name, _value) in &ed.members {
                    symbols.push(Symbol {
                        name: member_name.clone(),
                        kind: SymbolKind::EnumMember,
                        type_name: Some(ed.name.clone()),
                        visibility: ed.visibility.clone(),
                        span: ed.span,
                        detail: SymbolDetail::None,
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
}
