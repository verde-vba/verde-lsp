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
    pub name_span: TextRange,
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
    If(IfStatementNode),
    For(ForStatementNode),
    With(WithStatementNode),
    Select(SelectStatementNode),
    Call(CallStatementNode),
    Set(SetStatementNode),
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
    /// Names and optional types introduced by this declaration.
    /// For `Dim a As Long, b As String, c` this holds
    /// `[(a, Some("Long")), (b, Some("String")), (c, None)]`.
    pub names: Vec<(SmolStr, Option<SmolStr>)>,
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

/// Header line of an `If ... Then` statement. Only the header tokens are
/// captured; the block body (then/else branches, End If) lands as subsequent
/// statements in the enclosing procedure body. Downstream passes will
/// reconstruct block structure when needed.
#[derive(Debug, Clone, PartialEq)]
pub struct IfStatementNode {
    pub tokens: Vec<SpannedToken>,
    pub span: TextRange,
}

/// Header line of a `For ... [To|Each] ...` loop. Body statements and the
/// matching `Next` land separately in the enclosing procedure body.
#[derive(Debug, Clone, PartialEq)]
pub struct ForStatementNode {
    pub tokens: Vec<SpannedToken>,
    pub span: TextRange,
}

/// Header line of a `With obj` block. The inner statements and the matching
/// `End With` land separately in the enclosing procedure body.
#[derive(Debug, Clone, PartialEq)]
pub struct WithStatementNode {
    pub tokens: Vec<SpannedToken>,
    pub span: TextRange,
}

/// Header line of a `Select Case` block. Case arms, default, and the
/// matching `End Select` land separately in the enclosing procedure body.
#[derive(Debug, Clone, PartialEq)]
pub struct SelectStatementNode {
    pub tokens: Vec<SpannedToken>,
    pub span: TextRange,
}

/// A `Call Foo(...)` statement. Captured as raw tokens for now; argument
/// parsing is deferred to a later sprint.
#[derive(Debug, Clone, PartialEq)]
pub struct CallStatementNode {
    pub tokens: Vec<SpannedToken>,
    pub span: TextRange,
}

/// A `Set lhs = rhs` object-reference assignment. Captured as raw tokens for
/// now; semantic splitting of lhs/rhs is deferred to a later sprint.
#[derive(Debug, Clone, PartialEq)]
pub struct SetStatementNode {
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
