use smol_str::SmolStr;

use crate::parser::ast::*;

#[derive(Debug, Clone)]
pub struct SymbolTable {
    pub symbols: Vec<Symbol>,
    pub option_explicit: bool,
}

#[derive(Debug, Clone)]
pub struct Symbol {
    pub name: SmolStr,
    pub kind: SymbolKind,
    pub type_name: Option<SmolStr>,
    pub visibility: Visibility,
    pub span: TextRange,
    pub detail: SymbolDetail,
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

    for &node_id in &ast.root {
        match &ast.nodes[node_id] {
            AstNode::Procedure(proc) => {
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
                });
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
                });
                for (member_name, value) in &ed.members {
                    symbols.push(Symbol {
                        name: member_name.clone(),
                        kind: SymbolKind::EnumMember,
                        type_name: Some(ed.name.clone()),
                        visibility: ed.visibility.clone(),
                        span: ed.span,
                        detail: SymbolDetail::None,
                    });
                    let _ = value;
                }
            }
            _ => {}
        }
    }

    SymbolTable {
        symbols,
        option_explicit: ast.option_explicit,
    }
}
