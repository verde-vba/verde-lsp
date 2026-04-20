use la_arena::{Arena, Idx};
use smol_str::SmolStr;

use super::lexer::SpannedToken;

pub type NodeId = Idx<AstNode>;

#[derive(Debug, Clone)]
pub struct Ast {
    pub nodes: Arena<AstNode>,
    pub root: Vec<NodeId>,
    /// Whether the module begins with `Option Explicit`; when true,
    /// undeclared variable usages produce diagnostics.
    pub option_explicit: bool,
}

impl Default for Ast {
    fn default() -> Self {
        Self {
            nodes: Arena::new(),
            root: Vec::new(),
            option_explicit: false,
        }
    }
}

impl Ast {
    pub fn new() -> Self {
        Self::default()
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
    /// Byte span covering just the procedure body — from the character after
    /// the signature's terminating newline up to (but not including) the
    /// `End Sub`/`End Function`/`End Property` token. Consumers can slice
    /// `&source[body_range.start..body_range.end]` to scan body-only content
    /// without re-seeing the signature or surrounding module.
    pub body_range: TextRange,
}

#[derive(Debug, Clone, PartialEq)]
pub enum ProcedureKind {
    Sub,
    Function,
    PropertyGet,
    PropertyLet,
    PropertySet,
}

#[derive(Debug, Clone, PartialEq, Default)]
pub enum Visibility {
    #[default]
    Public,
    Private,
    Friend,
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

#[derive(Debug, Clone, PartialEq, Default)]
pub enum ParameterPassing {
    #[default]
    ByRef,
    ByVal,
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

#[derive(Debug, Clone, PartialEq)]
pub enum StatementNode {
    LocalDeclaration(LocalDeclarationNode),
    Expression(ExpressionStatementNode),
}

#[derive(Debug, Clone, PartialEq, Eq)]
pub enum DeclKind {
    Dim,
    Static,
    Const,
    ReDim,
}

#[derive(Debug, Clone, PartialEq)]
pub struct LocalDeclarationNode {
    pub kind: DeclKind,
    /// Names of locals introduced by this declaration. For `Dim a, b As String`
    /// this holds `[a, b]`. Types are intentionally not captured at this layer
    /// — diagnostics only need to know which identifiers are declared.
    pub names: Vec<SmolStr>,
    pub span: TextRange,
}

#[derive(Debug, Clone, PartialEq)]
pub struct ExpressionStatementNode {
    /// Raw tokens within the statement (excluding the terminating
    /// Newline/Colon). Preserves positional info so future AST walks can
    /// resolve identifier references without re-lexing the body.
    pub tokens: Vec<SpannedToken>,
    pub span: TextRange,
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
