use la_arena::{Arena, Idx};
use smol_str::SmolStr;

pub type NodeId = Idx<AstNode>;

#[derive(Debug, Clone)]
pub struct Ast {
    pub nodes: Arena<AstNode>,
    pub root: Vec<NodeId>,
    /// Whether the module begins with `Option Explicit`; when true,
    /// undeclared variable usages produce diagnostics.
    pub option_explicit: bool,
}

impl Ast {
    pub fn new() -> Self {
        Self {
            nodes: Arena::new(),
            root: Vec::new(),
            option_explicit: false,
        }
    }
}

#[derive(Debug, Clone)]
pub enum AstNode {
    Module(ModuleNode),
    Procedure(ProcedureNode),
    Variable(VariableNode),
    Parameter(ParameterNode),
    TypeDef(TypeDefNode),
    EnumDef(EnumDefNode),
    Statement(StatementNode),
}

#[derive(Debug, Clone)]
pub struct ModuleNode {
    pub name: SmolStr,
    pub children: Vec<NodeId>,
}

#[derive(Debug, Clone)]
pub struct ProcedureNode {
    pub name: SmolStr,
    pub kind: ProcedureKind,
    pub visibility: Visibility,
    pub params: Vec<NodeId>,
    pub return_type: Option<SmolStr>,
    pub body: Vec<NodeId>,
    pub span: TextRange,
}

#[derive(Debug, Clone, PartialEq)]
pub enum ProcedureKind {
    Sub,
    Function,
    PropertyGet,
    PropertyLet,
    PropertySet,
}

#[derive(Debug, Clone, PartialEq)]
pub enum Visibility {
    Public,
    Private,
    Friend,
}

impl Default for Visibility {
    fn default() -> Self {
        Visibility::Public
    }
}

#[derive(Debug, Clone)]
pub struct VariableNode {
    pub name: SmolStr,
    pub type_name: Option<SmolStr>,
    pub visibility: Visibility,
    pub is_const: bool,
    pub is_static: bool,
    pub span: TextRange,
}

#[derive(Debug, Clone)]
pub struct ParameterNode {
    pub name: SmolStr,
    pub type_name: Option<SmolStr>,
    pub passing: ParameterPassing,
    pub is_optional: bool,
    pub is_param_array: bool,
    pub span: TextRange,
}

#[derive(Debug, Clone, PartialEq)]
pub enum ParameterPassing {
    ByRef,
    ByVal,
}

impl Default for ParameterPassing {
    fn default() -> Self {
        ParameterPassing::ByRef
    }
}

#[derive(Debug, Clone)]
pub struct TypeDefNode {
    pub name: SmolStr,
    pub visibility: Visibility,
    pub members: Vec<NodeId>,
    pub span: TextRange,
}

#[derive(Debug, Clone)]
pub struct EnumDefNode {
    pub name: SmolStr,
    pub visibility: Visibility,
    pub members: Vec<(SmolStr, Option<i64>)>,
    pub span: TextRange,
}

#[derive(Debug, Clone)]
pub enum StatementNode {
    Assignment {
        target: SmolStr,
        span: TextRange,
    },
    Call {
        name: SmolStr,
        span: TextRange,
    },
    Other {
        span: TextRange,
    },
}

#[derive(Debug, Clone, Copy, PartialEq, Eq)]
pub struct TextRange {
    pub start: u32,
    pub end: u32,
}

impl TextRange {
    pub fn new(start: usize, end: usize) -> Self {
        Self {
            start: start as u32,
            end: end as u32,
        }
    }
}
