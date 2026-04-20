use smol_str::SmolStr;

use super::ast::*;
use super::lexer::{self, Token};

pub struct ParseResult {
    pub ast: Ast,
    pub errors: Vec<ParseError>,
}

#[derive(Debug, Clone)]
pub struct ParseError {
    pub message: String,
    pub span: std::ops::Range<usize>,
}

pub fn parse(source: &str) -> ParseResult {
    let tokens = lexer::lex(source);
    let mut parser = Parser {
        tokens: &tokens,
        pos: 0,
        ast: Ast::new(),
        errors: Vec::new(),
    };

    parser.parse_module();

    ParseResult {
        ast: parser.ast,
        errors: parser.errors,
    }
}

struct Parser<'a> {
    tokens: &'a [lexer::SpannedToken],
    pos: usize,
    ast: Ast,
    errors: Vec<ParseError>,
}

impl<'a> Parser<'a> {
    fn peek(&self) -> Option<&lexer::SpannedToken> {
        self.tokens.get(self.pos)
    }

    fn skip_newlines(&mut self) {
        while matches!(self.peek(), Some(t) if t.token == Token::Newline || t.token == Token::Comment)
        {
            self.pos += 1;
        }
    }

    /// Advance past consecutive `LineContinuation` / `Newline` tokens. Used at
    /// sites where a VBA `_` line continuation is legal — most notably inside
    /// parameter lists, where a continuation may sit between a comma and the
    /// next parameter slot, or just before a closing paren.
    fn skip_line_continuations(&mut self) {
        while matches!(
            self.peek(),
            Some(t) if t.token == Token::LineContinuation || t.token == Token::Newline
        ) {
            self.pos += 1;
        }
    }

    /// Skip only `_` line-continuation tokens, leaving any plain `Newline`
    /// in place. Use at sites outside a parenthesized list where a standalone
    /// Newline actually terminates the construct we're parsing (e.g. the
    /// signature line of a procedure), but a `_` continuation should still
    /// transparently join the next physical line.
    fn skip_line_continuations_preserving_newline(&mut self) {
        while matches!(self.peek(), Some(t) if t.token == Token::LineContinuation) {
            self.pos += 1;
        }
    }

    /// Advance past tokens that separate statements inside a procedure body:
    /// Newline, Colon (VBA's `:` statement separator), LineContinuation, and
    /// Comment. Called between statement emissions in the body loop to land
    /// `pos` on the next meaningful token (or EOF / `End <kind>`).
    fn skip_statement_separators(&mut self) {
        while matches!(
            self.peek(),
            Some(t) if matches!(
                t.token,
                Token::Newline | Token::Colon | Token::LineContinuation | Token::Comment
            )
        ) {
            self.pos += 1;
        }
    }

    fn parse_module(&mut self) {
        self.skip_newlines();

        while self.pos < self.tokens.len() {
            self.skip_newlines();
            if self.pos >= self.tokens.len() {
                break;
            }

            let tok = &self.tokens[self.pos];
            match &tok.token {
                Token::Option => self.parse_option(),
                Token::Attribute => self.skip_attribute_line(),
                Token::Public | Token::Private | Token::Friend => {
                    self.parse_declaration_with_visibility()
                }
                Token::Sub => self.parse_procedure(Visibility::Public, ProcedureKind::Sub),
                Token::Function => {
                    self.parse_procedure(Visibility::Public, ProcedureKind::Function)
                }
                Token::Property => self.parse_property(Visibility::Public),
                Token::Dim | Token::Const => self.parse_variable(Visibility::Private),
                Token::Type => self.parse_type_def(Visibility::Public),
                Token::Enum => self.parse_enum_def(Visibility::Public),
                _ => {
                    self.pos += 1;
                }
            }
        }
    }

    fn parse_option(&mut self) {
        self.pos += 1; // skip "Option"
        if let Some(tok) = self.peek() {
            if tok.token == Token::Explicit {
                self.ast.option_explicit = true;
                self.pos += 1;
            }
        }
    }

    fn skip_attribute_line(&mut self) {
        while self.pos < self.tokens.len() {
            if self.tokens[self.pos].token == Token::Newline {
                break;
            }
            self.pos += 1;
        }
    }

    fn parse_declaration_with_visibility(&mut self) {
        let visibility = match self.tokens[self.pos].token {
            Token::Public => Visibility::Public,
            Token::Private => Visibility::Private,
            Token::Friend => Visibility::Friend,
            _ => unreachable!(),
        };
        self.pos += 1;
        self.skip_newlines();

        if let Some(tok) = self.peek() {
            match &tok.token {
                Token::Sub => self.parse_procedure(visibility, ProcedureKind::Sub),
                Token::Function => self.parse_procedure(visibility, ProcedureKind::Function),
                Token::Property => self.parse_property(visibility),
                Token::Type => self.parse_type_def(visibility),
                Token::Enum => self.parse_enum_def(visibility),
                Token::Const => self.parse_variable(visibility),
                _ => self.parse_variable(visibility),
            }
        }
    }

    fn parse_procedure(&mut self, visibility: Visibility, kind: ProcedureKind) {
        let start = self.tokens[self.pos].span.start;
        self.pos += 1; // skip Sub/Function

        let (name, name_span) = if let Some(tok) = self.peek() {
            if tok.token == Token::Identifier {
                let n = tok.text.clone();
                let s = TextRange::new(tok.span.start, tok.span.end);
                self.pos += 1;
                (n, s)
            } else {
                (SmolStr::new("?"), TextRange::new(start, start))
            }
        } else {
            return;
        };

        // Parse the parameter list `(...)` if present. VBA allows `Sub Foo`
        // without parentheses for a no-arg procedure, so we only descend into
        // parameter parsing when an opening paren is the next token.
        let params = if matches!(self.peek(), Some(t) if t.token == Token::LParen) {
            self.parse_parameter_list()
        } else {
            Vec::new()
        };

        // Optional `As <ReturnType>` clause for Function and Property Get.
        // `Sub`, `Property Let`, and `Property Set` do not carry a return
        // type; a missing `As` in a legacy Function is tolerated (stays None).
        // A `_` line continuation may sit between `)` and `As`, so we peek
        // across continuations — but must not consume the signature-terminating
        // Newline that the signature-advancing scan below relies on.
        let return_type = match kind {
            ProcedureKind::Function | ProcedureKind::PropertyGet => {
                self.skip_line_continuations_preserving_newline();
                if matches!(self.peek(), Some(t) if t.token == Token::As) {
                    self.pos += 1; // consume `As`
                    self.skip_line_continuations_preserving_newline();
                    self.parse_dotted_type_name().map(|(name, _end)| name)
                } else {
                    None
                }
            }
            ProcedureKind::Sub | ProcedureKind::PropertyLet | ProcedureKind::PropertySet => None,
        };

        // Skip to end of procedure
        let end_token = match kind {
            ProcedureKind::Sub => Token::EndSub,
            ProcedureKind::Function => Token::EndFunction,
            ProcedureKind::PropertyGet
            | ProcedureKind::PropertyLet
            | ProcedureKind::PropertySet => Token::EndProperty,
        };

        // Advance past the signature line so the body loop does not re-see
        // signature tokens. We look for either the signature-terminating
        // Newline or the `end_token` itself (single-line procedures are not
        // really a thing in VBA, but be defensive).
        while self.pos < self.tokens.len() {
            let tok = &self.tokens[self.pos];
            if tok.token == end_token {
                break;
            }
            if tok.token == Token::Newline {
                self.pos += 1;
                break;
            }
            self.pos += 1;
        }

        // Parse body statements until the matching End token (or EOF).
        let mut body: Vec<NodeId> = Vec::new();
        loop {
            self.skip_statement_separators();

            match self.peek() {
                None => break,
                Some(t) if t.token == end_token => {
                    let end = t.span.end;
                    self.pos += 1;

                    let node = AstNode::Procedure(ProcedureNode {
                        name,
                        name_span,
                        kind,
                        visibility,
                        params,
                        return_type,
                        body,
                        span: TextRange::new(start, end),
                    });
                    let id = self.ast.nodes.alloc(node);
                    self.ast.root.push(id);
                    return;
                }
                Some(_) => {
                    let stmt_node = self.classify_and_parse_statement();
                    let id = self.ast.nodes.alloc(stmt_node);
                    body.push(id);
                }
            }
        }

        let end = self.tokens.last().map(|t| t.span.end).unwrap_or(start);
        self.errors.push(ParseError {
            message: format!("Unterminated procedure: {}", name),
            span: start..end,
        });
    }

    /// Parse `(param, param, ...)` starting at the `LParen`. Returns arena
    /// indices of `AstNode::Parameter` entries. Leaves `pos` just past the
    /// matching `RParen` (or at end-of-stream if malformed).
    fn parse_parameter_list(&mut self) -> Vec<NodeId> {
        debug_assert!(matches!(self.peek(), Some(t) if t.token == Token::LParen));
        self.pos += 1; // consume `(`

        let mut params = Vec::new();
        loop {
            // Allow line continuations / whitespace tokens between parameters.
            self.skip_line_continuations();

            match self.peek() {
                None => return params,
                Some(t) if t.token == Token::RParen => {
                    self.pos += 1;
                    return params;
                }
                Some(t) if t.token == Token::Comma => {
                    // Stray leading/trailing comma; skip and continue.
                    self.pos += 1;
                    continue;
                }
                _ => {}
            }

            if let Some(id) = self.parse_one_parameter() {
                params.push(id);
            }

            // After a parameter, expect comma or closing paren. Skip anything
            // else defensively until we find one. Tolerate line continuations
            // that sit between a parameter and its trailing comma or between
            // the last parameter and the closing paren.
            loop {
                self.skip_line_continuations();
                match self.peek() {
                    None => return params,
                    Some(t) if t.token == Token::Comma => {
                        self.pos += 1;
                        break;
                    }
                    Some(t) if t.token == Token::RParen => {
                        self.pos += 1;
                        return params;
                    }
                    _ => {
                        self.pos += 1;
                    }
                }
            }
        }
    }

    /// Parse a single parameter slot: `[Optional] [ByVal|ByRef] [ParamArray]
    /// Ident [()] [As Type] [= default]`. Returns the arena index of the
    /// resulting `ParameterNode`, or `None` if no identifier was found.
    fn parse_one_parameter(&mut self) -> Option<NodeId> {
        let start_pos = self.pos;
        let start_offset = self.tokens.get(start_pos).map(|t| t.span.start)?;

        let mut is_optional = false;
        let mut is_param_array = false;
        let mut passing = ParameterPassing::ByRef;

        // Modifiers can appear in any reasonable order; accept a few permutations.
        loop {
            match self.peek().map(|t| &t.token) {
                Some(Token::Optional) => {
                    is_optional = true;
                    self.pos += 1;
                }
                Some(Token::ByVal) => {
                    passing = ParameterPassing::ByVal;
                    self.pos += 1;
                }
                Some(Token::ByRef) => {
                    passing = ParameterPassing::ByRef;
                    self.pos += 1;
                }
                Some(Token::ParamArray) => {
                    is_param_array = true;
                    self.pos += 1;
                }
                _ => break,
            }
        }

        // Parameter name.
        let name = match self.peek() {
            Some(t) if t.token == Token::Identifier => {
                let n = t.text.clone();
                self.pos += 1;
                n
            }
            _ => return None,
        };

        let mut end_offset = self
            .tokens
            .get(self.pos.saturating_sub(1))
            .map(|t| t.span.end)
            .unwrap_or(start_offset);

        // Optional trailing `()` for array parameters — skip, do not capture.
        if matches!(self.peek(), Some(t) if t.token == Token::LParen) {
            self.pos += 1;
            while let Some(t) = self.peek() {
                if t.token == Token::RParen {
                    end_offset = t.span.end;
                    self.pos += 1;
                    break;
                }
                self.pos += 1;
            }
        }

        // Optional `As <Type>` clause. Supports dotted UDT names such as
        // `ADODB.Connection` or `Microsoft.Office.Core.CommandBar`.
        let mut type_name = None;
        if matches!(self.peek(), Some(t) if t.token == Token::As) {
            self.pos += 1;
            if let Some((name, end)) = self.parse_dotted_type_name() {
                end_offset = end;
                type_name = Some(name);
            }
        }

        // Optional `= <default>` — skip tokens until the next top-level comma
        // or closing paren, without capturing the default expression.
        if matches!(self.peek(), Some(t) if t.token == Token::Eq) {
            self.pos += 1;
            let mut depth: i32 = 0;
            while let Some(t) = self.peek() {
                match t.token {
                    Token::LParen => {
                        depth += 1;
                        self.pos += 1;
                    }
                    Token::RParen if depth == 0 => break,
                    Token::RParen => {
                        depth -= 1;
                        self.pos += 1;
                    }
                    Token::Comma if depth == 0 => break,
                    _ => {
                        end_offset = t.span.end;
                        self.pos += 1;
                    }
                }
            }
        }

        let node = AstNode::Parameter(ParameterNode {
            name,
            type_name,
            passing,
            is_optional,
            is_param_array,
            span: TextRange::new(start_offset, end_offset),
        });
        Some(self.ast.nodes.alloc(node))
    }

    /// Parse a type name following an `As` keyword. Starts at the first token
    /// of the name (the `As` itself must already be consumed). Accepts any
    /// type-ish token (builtin type keyword or identifier) for the head and
    /// greedily extends with `.Ident` segments so dotted UDT names such as
    /// `ADODB.Connection` or `Microsoft.Office.Core.CommandBar` land in a
    /// single `SmolStr`. Returns the joined name and the end offset of the
    /// last consumed token, or `None` if no type-ish token was present.
    fn parse_dotted_type_name(&mut self) -> Option<(SmolStr, usize)> {
        let t = self.peek()?;
        let mut buf = String::from(t.text.as_str());
        let mut end_offset = t.span.end;
        self.pos += 1;

        while matches!(self.peek(), Some(t) if t.token == Token::Dot)
            && matches!(
                self.tokens.get(self.pos + 1),
                Some(t) if t.token == Token::Identifier
            )
        {
            self.pos += 1; // consume Dot
            let ident = &self.tokens[self.pos];
            buf.push('.');
            buf.push_str(ident.text.as_str());
            end_offset = ident.span.end;
            self.pos += 1;
        }

        Some((SmolStr::from(buf), end_offset))
    }

    /// Parse a local declaration starting at the Dim/Static/Const keyword.
    /// Collects declared identifier names and their optional `As <Type>` clauses,
    /// stopping at a statement terminator (Newline/Colon) or EOF.
    /// Leaves `pos` on the terminator (or EOF); the caller skips it.
    fn parse_local_declaration(&mut self, kind: DeclKind) -> LocalDeclarationNode {
        let start_offset = self.tokens.get(self.pos).map(|t| t.span.start).unwrap_or(0);
        self.pos += 1; // consume the kind keyword (Dim/Static/Const)

        // (ReDim is now routed to RedimStatementNode and never reaches here.)
        if matches!(self.peek(), Some(t) if t.token == Token::Preserve) {
            self.pos += 1;
        }

        let mut names: Vec<(SmolStr, Option<SmolStr>, TextRange)> = Vec::new();
        let mut end_offset = start_offset;
        let mut paren_depth: i32 = 0;
        let mut expect_name = true;

        while let Some(t) = self.peek() {
            match t.token {
                Token::Newline | Token::Colon if paren_depth == 0 => break,
                Token::LineContinuation => {
                    end_offset = t.span.end;
                    self.pos += 1;
                }
                Token::Identifier if expect_name && paren_depth == 0 => {
                    let name_span = TextRange::new(t.span.start, t.span.end);
                    names.push((t.text.clone(), None, name_span));
                    end_offset = t.span.end;
                    expect_name = false;
                    self.pos += 1;
                }
                Token::As if !expect_name && paren_depth == 0 => {
                    self.pos += 1; // consume `As`
                    if let Some((type_name, end)) = self.parse_dotted_type_name() {
                        end_offset = end;
                        if let Some(last) = names.last_mut() {
                            last.1 = Some(type_name);
                        }
                    }
                }
                Token::Comma if paren_depth == 0 => {
                    end_offset = t.span.end;
                    expect_name = true;
                    self.pos += 1;
                }
                Token::LParen => {
                    paren_depth += 1;
                    end_offset = t.span.end;
                    self.pos += 1;
                }
                Token::RParen => {
                    paren_depth = (paren_depth - 1).max(0);
                    end_offset = t.span.end;
                    self.pos += 1;
                }
                _ => {
                    // `= expr`, array bounds, etc.
                    end_offset = t.span.end;
                    expect_name = false;
                    self.pos += 1;
                }
            }
        }

        LocalDeclarationNode {
            kind,
            names,
            span: TextRange::new(start_offset, end_offset),
        }
    }

    /// Dispatch to the appropriate per-statement parser based on the current
    /// leading token. Declaration keywords (Dim/Static/Const) produce a
    /// `LocalDeclaration`; block-opening keywords (If/For/With/Select/While/Do)
    /// and array-resize keyword (ReDim) and top-level-statement keywords
    /// (Call/Set/Exit/GoTo/On) produce the matching `StatementNode` variant
    /// with the header-line tokens captured. Anything else falls back to an
    /// `Expression` statement. Callers must ensure `pos` is on the first token
    /// of the statement; this function returns a ready-to-alloc
    /// `AstNode::Statement`.
    fn classify_and_parse_statement(&mut self) -> AstNode {
        let head = self.peek().map(|t| t.token.clone());
        let decl_kind = match head {
            Some(Token::Dim) => Some(DeclKind::Dim),
            Some(Token::Static) => Some(DeclKind::Static),
            Some(Token::Const) => Some(DeclKind::Const),
            _ => None,
        };
        if let Some(kind) = decl_kind {
            let decl = self.parse_local_declaration(kind);
            return AstNode::Statement(StatementNode::LocalDeclaration(decl));
        }

        let stmt = match head {
            Some(Token::If) => {
                let (tokens, span) = self.collect_statement_tokens();
                StatementNode::If(IfStatementNode { tokens, span })
            }
            Some(Token::For) => {
                let (tokens, span) = self.collect_statement_tokens();
                StatementNode::For(ForStatementNode { tokens, span })
            }
            Some(Token::With) => {
                let (tokens, span) = self.collect_statement_tokens();
                StatementNode::With(WithStatementNode { tokens, span })
            }
            Some(Token::Select) => {
                let (tokens, span) = self.collect_statement_tokens();
                StatementNode::Select(SelectStatementNode { tokens, span })
            }
            Some(Token::Call) => {
                let (tokens, span) = self.collect_statement_tokens();
                StatementNode::Call(CallStatementNode { tokens, span })
            }
            Some(Token::Set) => {
                let (tokens, span) = self.collect_statement_tokens();
                StatementNode::Set(SetStatementNode { tokens, span })
            }
            Some(Token::While) => {
                let (tokens, span) = self.collect_statement_tokens();
                StatementNode::While(WhileStatementNode { tokens, span })
            }
            Some(Token::Do) => {
                let (tokens, span) = self.collect_statement_tokens();
                StatementNode::Do(DoStatementNode { tokens, span })
            }
            Some(Token::ReDim) => {
                let (tokens, span) = self.collect_statement_tokens();
                StatementNode::Redim(RedimStatementNode { tokens, span })
            }
            Some(Token::Exit) => {
                let (tokens, span) = self.collect_statement_tokens();
                StatementNode::Exit(ExitStatementNode { tokens, span })
            }
            Some(Token::GoTo) => {
                let (tokens, span) = self.collect_statement_tokens();
                StatementNode::GoTo(GoToStatementNode { tokens, span })
            }
            Some(Token::On) => {
                let (tokens, span) = self.collect_statement_tokens();
                StatementNode::OnError(OnErrorStatementNode { tokens, span })
            }
            _ => StatementNode::Expression(self.parse_expression_statement()),
        };
        AstNode::Statement(stmt)
    }

    /// Collect tokens for one statement up to the next Newline/Colon (outside
    /// parentheses). Line continuations are consumed as part of the current
    /// statement and are included in the captured token stream so downstream
    /// consumers retain positional fidelity. Leaves `pos` on the terminator
    /// (or EOF); the caller skips it. Returns the raw token vector and the
    /// covering span.
    fn collect_statement_tokens(&mut self) -> (Vec<lexer::SpannedToken>, TextRange) {
        let start_offset = self.tokens.get(self.pos).map(|t| t.span.start).unwrap_or(0);
        let mut end_offset = start_offset;
        let mut tokens: Vec<lexer::SpannedToken> = Vec::new();
        let mut paren_depth: i32 = 0;

        while let Some(t) = self.peek() {
            match t.token {
                Token::Newline | Token::Colon if paren_depth == 0 => break,
                Token::LParen => {
                    paren_depth += 1;
                    end_offset = t.span.end;
                    tokens.push(t.clone());
                    self.pos += 1;
                }
                Token::RParen => {
                    paren_depth = (paren_depth - 1).max(0);
                    end_offset = t.span.end;
                    tokens.push(t.clone());
                    self.pos += 1;
                }
                _ => {
                    end_offset = t.span.end;
                    tokens.push(t.clone());
                    self.pos += 1;
                }
            }
        }

        (tokens, TextRange::new(start_offset, end_offset))
    }

    fn parse_expression_statement(&mut self) -> ExpressionStatementNode {
        let (tokens, span) = self.collect_statement_tokens();
        ExpressionStatementNode { tokens, span }
    }

    fn parse_property(&mut self, visibility: Visibility) {
        self.pos += 1; // skip Property
        let kind = if let Some(tok) = self.peek() {
            match tok.token {
                Token::Get => {
                    self.pos += 1;
                    ProcedureKind::PropertyGet
                }
                Token::Let => {
                    self.pos += 1;
                    ProcedureKind::PropertyLet
                }
                _ => {
                    self.pos += 1;
                    ProcedureKind::PropertySet
                }
            }
        } else {
            return;
        };

        self.parse_procedure(visibility, kind);
    }

    fn parse_variable(&mut self, visibility: Visibility) {
        let is_const = matches!(self.peek(), Some(t) if t.token == Token::Const);
        if is_const {
            self.pos += 1;
        }

        if let Some(tok) = self.peek() {
            if tok.token == Token::Identifier {
                let name = tok.text.clone();
                let start = tok.span.start;
                let mut end = tok.span.end;
                self.pos += 1;

                let mut type_name = None;
                if matches!(self.peek(), Some(t) if t.token == Token::As) {
                    self.pos += 1;
                    if let Some(type_tok) = self.peek() {
                        type_name = Some(type_tok.text.clone());
                        end = type_tok.span.end;
                        self.pos += 1;
                    }
                }

                let node = AstNode::Variable(VariableNode {
                    name,
                    type_name,
                    visibility,
                    is_const,
                    is_static: false,
                    span: TextRange::new(start, end),
                });
                let id = self.ast.nodes.alloc(node);
                self.ast.root.push(id);
            }
        }
    }

    fn parse_type_def(&mut self, visibility: Visibility) {
        let start = self.tokens[self.pos].span.start;
        self.pos += 1;

        let name = if let Some(tok) = self.peek() {
            if tok.token == Token::Identifier {
                let n = tok.text.clone();
                self.pos += 1;
                n
            } else {
                SmolStr::new("?")
            }
        } else {
            return;
        };

        while self.pos < self.tokens.len() {
            if self.tokens[self.pos].token == Token::EndType {
                let end = self.tokens[self.pos].span.end;
                self.pos += 1;

                let node = AstNode::TypeDef(TypeDefNode {
                    name,
                    visibility,
                    members: Vec::new(),
                    span: TextRange::new(start, end),
                });
                let id = self.ast.nodes.alloc(node);
                self.ast.root.push(id);
                return;
            }
            self.pos += 1;
        }
    }

    fn parse_enum_def(&mut self, visibility: Visibility) {
        let start = self.tokens[self.pos].span.start;
        self.pos += 1;

        let name = if let Some(tok) = self.peek() {
            if tok.token == Token::Identifier {
                let n = tok.text.clone();
                self.pos += 1;
                n
            } else {
                SmolStr::new("?")
            }
        } else {
            return;
        };

        let mut members = Vec::new();
        while self.pos < self.tokens.len() {
            self.skip_newlines();
            if self.pos >= self.tokens.len() {
                break;
            }
            if self.tokens[self.pos].token == Token::EndEnum {
                let end = self.tokens[self.pos].span.end;
                self.pos += 1;

                let node = AstNode::EnumDef(EnumDefNode {
                    name,
                    visibility,
                    members,
                    span: TextRange::new(start, end),
                });
                let id = self.ast.nodes.alloc(node);
                self.ast.root.push(id);
                return;
            }
            if self.tokens[self.pos].token == Token::Identifier {
                members.push((self.tokens[self.pos].text.clone(), None));
            }
            self.pos += 1;
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    fn first_procedure(ast: &Ast) -> &ProcedureNode {
        ast.nodes
            .iter()
            .find_map(|(_, n)| match n {
                AstNode::Procedure(p) => Some(p),
                _ => None,
            })
            .expect("expected a procedure")
    }

    fn parameter(ast: &Ast, id: NodeId) -> &ParameterNode {
        match &ast.nodes[id] {
            AstNode::Parameter(p) => p,
            other => panic!("expected Parameter node, got {:?}", other),
        }
    }

    #[test]
    fn parse_procedure_captures_single_parameter() {
        let result = parse("Sub Foo(x As Long)\nEnd Sub\n");
        let proc = first_procedure(&result.ast);
        assert_eq!(proc.params.len(), 1, "expected 1 parameter");
        let p = parameter(&result.ast, proc.params[0]);
        assert_eq!(p.name.as_str(), "x");
        assert_eq!(p.type_name.as_ref().map(|s| s.as_str()), Some("Long"));
    }

    #[test]
    fn parse_procedure_captures_multiple_parameters() {
        let result = parse("Sub Foo(a, b As String, c As Long)\nEnd Sub\n");
        let proc = first_procedure(&result.ast);
        assert_eq!(proc.params.len(), 3, "expected 3 parameters");
        let names: Vec<&str> = proc
            .params
            .iter()
            .map(|id| parameter(&result.ast, *id).name.as_str())
            .collect();
        assert_eq!(names, vec!["a", "b", "c"]);
    }

    #[test]
    fn parse_procedure_captures_byval_optional_paramarray() {
        let result = parse(
            "Sub Foo(ByVal a As Long, Optional b As String = \"x\", ParamArray rest() As Variant)\nEnd Sub\n",
        );
        let proc = first_procedure(&result.ast);
        assert_eq!(proc.params.len(), 3, "expected 3 parameters");
        let a = parameter(&result.ast, proc.params[0]);
        let b = parameter(&result.ast, proc.params[1]);
        let rest = parameter(&result.ast, proc.params[2]);
        assert_eq!(a.passing, ParameterPassing::ByVal, "a should be ByVal");
        assert!(b.is_optional, "b should be optional");
        assert!(rest.is_param_array, "rest should be ParamArray");
        assert_eq!(a.name.as_str(), "a");
        assert_eq!(b.name.as_str(), "b");
        assert_eq!(rest.name.as_str(), "rest");
    }

    #[test]
    fn parse_procedure_with_empty_params() {
        let result = parse("Sub Foo()\nEnd Sub\n");
        let proc = first_procedure(&result.ast);
        assert!(
            proc.params.is_empty(),
            "expected no params, got {}",
            proc.params.len()
        );
    }

    fn statement(ast: &Ast, id: NodeId) -> &StatementNode {
        match &ast.nodes[id] {
            AstNode::Statement(s) => s,
            other => panic!("expected Statement node, got {:?}", other),
        }
    }

    #[test]
    fn parse_procedure_populates_body_with_expression_statement() {
        let result = parse("Sub Foo()\n    x = 1\nEnd Sub\n");
        let proc = first_procedure(&result.ast);
        assert_eq!(
            proc.body.len(),
            1,
            "expected 1 body statement, got {}",
            proc.body.len()
        );
        match statement(&result.ast, proc.body[0]) {
            StatementNode::Expression(expr) => {
                assert!(!expr.tokens.is_empty(), "expected non-empty tokens");
                assert!(
                    expr.tokens
                        .iter()
                        .any(|t| t.token == Token::Identifier && t.text.as_str() == "x"),
                    "expected 'x' identifier token in expression tokens"
                );
            }
            other => panic!("expected Expression statement, got {:?}", other),
        }
    }

    #[test]
    fn parse_procedure_captures_local_dim_declaration() {
        let result = parse("Sub Foo()\n    Dim x As Long\nEnd Sub\n");
        let proc = first_procedure(&result.ast);
        assert_eq!(proc.body.len(), 1, "expected 1 body statement");
        match statement(&result.ast, proc.body[0]) {
            StatementNode::LocalDeclaration(d) => {
                assert_eq!(d.kind, DeclKind::Dim);
                let names: Vec<&str> = d.names.iter().map(|(n, _, _)| n.as_str()).collect();
                assert_eq!(names, vec!["x"]);
                assert_eq!(d.names[0].1.as_deref(), Some("Long"));
            }
            other => panic!("expected LocalDeclaration, got {:?}", other),
        }
    }

    #[test]
    fn parse_procedure_captures_multiple_dim_names() {
        let result = parse("Sub Foo()\n    Dim a As Long, b As String, c\nEnd Sub\n");
        let proc = first_procedure(&result.ast);
        assert_eq!(proc.body.len(), 1, "expected 1 body statement");
        match statement(&result.ast, proc.body[0]) {
            StatementNode::LocalDeclaration(d) => {
                assert_eq!(d.kind, DeclKind::Dim);
                let names: Vec<&str> = d.names.iter().map(|(n, _, _)| n.as_str()).collect();
                assert_eq!(names, vec!["a", "b", "c"]);
                assert_eq!(d.names[0].1.as_deref(), Some("Long"));
                assert_eq!(d.names[1].1.as_deref(), Some("String"));
                assert_eq!(d.names[2].1, None);
            }
            other => panic!("expected LocalDeclaration, got {:?}", other),
        }
    }

    #[test]
    fn parse_procedure_captures_multiple_statements_in_order() {
        let result =
            parse("Sub Foo()\n    Dim x As Long\n    x = 1\n    Call DoStuff(x)\nEnd Sub\n");
        let proc = first_procedure(&result.ast);
        assert_eq!(
            proc.body.len(),
            3,
            "expected 3 body statements, got {}",
            proc.body.len()
        );
        match statement(&result.ast, proc.body[0]) {
            StatementNode::LocalDeclaration(d) => {
                let names: Vec<&str> = d.names.iter().map(|(n, _, _)| n.as_str()).collect();
                assert_eq!(names, vec!["x"]);
            }
            other => panic!("expected LocalDeclaration first, got {:?}", other),
        }
        assert!(matches!(
            statement(&result.ast, proc.body[1]),
            StatementNode::Expression(_)
        ));
        assert!(matches!(
            statement(&result.ast, proc.body[2]),
            StatementNode::Call(_)
        ));
    }

    #[test]
    fn parse_procedure_captures_const_declaration() {
        let result = parse("Sub Foo()\n    Const PI As Double = 3.14\nEnd Sub\n");
        let proc = first_procedure(&result.ast);
        assert_eq!(proc.body.len(), 1, "expected 1 body statement");
        match statement(&result.ast, proc.body[0]) {
            StatementNode::LocalDeclaration(d) => {
                assert_eq!(d.kind, DeclKind::Const);
                let names: Vec<&str> = d.names.iter().map(|(n, _, _)| n.as_str()).collect();
                assert_eq!(names, vec!["PI"]);
                assert_eq!(d.names[0].1.as_deref(), Some("Double"));
            }
            other => panic!("expected LocalDeclaration, got {:?}", other),
        }
    }

    #[test]
    fn parse_procedure_captures_dotted_udt_type() {
        let result = parse("Sub Foo(x As ADODB.Connection)\nEnd Sub\n");
        let proc = first_procedure(&result.ast);
        assert_eq!(proc.params.len(), 1, "expected 1 parameter");
        let p = parameter(&result.ast, proc.params[0]);
        assert_eq!(p.name.as_str(), "x");
        assert_eq!(
            p.type_name.as_ref().map(|s| s.as_str()),
            Some("ADODB.Connection")
        );
    }

    #[test]
    fn parse_procedure_captures_deeply_dotted_type() {
        let result = parse("Sub Foo(x As Microsoft.Office.Core.CommandBar)\nEnd Sub\n");
        let proc = first_procedure(&result.ast);
        assert_eq!(proc.params.len(), 1, "expected 1 parameter");
        let p = parameter(&result.ast, proc.params[0]);
        assert_eq!(
            p.type_name.as_ref().map(|s| s.as_str()),
            Some("Microsoft.Office.Core.CommandBar")
        );
    }

    #[test]
    fn parse_procedure_handles_multiline_parameter_list() {
        let source = "Sub Foo(ByVal a As Long, _\n        ByVal b As String, _\n        c As Variant)\nEnd Sub\n";
        let result = parse(source);
        let proc = first_procedure(&result.ast);
        assert_eq!(
            proc.params.len(),
            3,
            "expected 3 parameters, got {}",
            proc.params.len()
        );
        let a = parameter(&result.ast, proc.params[0]);
        let b = parameter(&result.ast, proc.params[1]);
        let c = parameter(&result.ast, proc.params[2]);
        assert_eq!(a.name.as_str(), "a");
        assert_eq!(b.name.as_str(), "b");
        assert_eq!(c.name.as_str(), "c");
        assert_eq!(a.type_name.as_ref().map(|s| s.as_str()), Some("Long"));
        assert_eq!(b.type_name.as_ref().map(|s| s.as_str()), Some("String"));
        assert_eq!(c.type_name.as_ref().map(|s| s.as_str()), Some("Variant"));
        assert_eq!(a.passing, ParameterPassing::ByVal, "a should be ByVal");
        assert_eq!(b.passing, ParameterPassing::ByVal, "b should be ByVal");
    }

    #[test]
    fn parse_procedure_handles_multiline_with_dotted_types() {
        let source =
            "Sub Foo(x As ADODB.Connection, _\n        y As Scripting.Dictionary)\nEnd Sub\n";
        let result = parse(source);
        let proc = first_procedure(&result.ast);
        assert_eq!(
            proc.params.len(),
            2,
            "expected 2 parameters, got {}",
            proc.params.len()
        );
        let x = parameter(&result.ast, proc.params[0]);
        let y = parameter(&result.ast, proc.params[1]);
        assert_eq!(
            x.type_name.as_ref().map(|s| s.as_str()),
            Some("ADODB.Connection")
        );
        assert_eq!(
            y.type_name.as_ref().map(|s| s.as_str()),
            Some("Scripting.Dictionary")
        );
    }

    #[test]
    fn parse_function_captures_simple_return_type() {
        let result = parse("Function Foo() As String\nEnd Function\n");
        let proc = first_procedure(&result.ast);
        assert_eq!(proc.kind, ProcedureKind::Function);
        assert_eq!(
            proc.return_type.as_ref().map(|s| s.as_str()),
            Some("String")
        );
    }

    #[test]
    fn parse_function_captures_dotted_return_type() {
        let result = parse("Function Bar() As ADODB.Connection\nEnd Function\n");
        let proc = first_procedure(&result.ast);
        assert_eq!(
            proc.return_type.as_ref().map(|s| s.as_str()),
            Some("ADODB.Connection")
        );
    }

    #[test]
    fn parse_function_captures_deeply_dotted_return_type() {
        let result = parse("Function Baz() As Microsoft.Office.Core.CommandBar\nEnd Function\n");
        let proc = first_procedure(&result.ast);
        assert_eq!(
            proc.return_type.as_ref().map(|s| s.as_str()),
            Some("Microsoft.Office.Core.CommandBar")
        );
    }

    #[test]
    fn parse_property_get_captures_return_type() {
        let result = parse("Property Get Value() As Long\nEnd Property\n");
        let proc = first_procedure(&result.ast);
        assert_eq!(proc.kind, ProcedureKind::PropertyGet);
        assert_eq!(proc.return_type.as_ref().map(|s| s.as_str()), Some("Long"));
    }

    #[test]
    fn parse_property_let_has_no_return_type() {
        let result = parse("Property Let Value(ByVal v As Long)\nEnd Property\n");
        let proc = first_procedure(&result.ast);
        assert_eq!(proc.kind, ProcedureKind::PropertyLet);
        assert!(
            proc.return_type.is_none(),
            "Property Let should not have a return type, got {:?}",
            proc.return_type
        );
    }

    #[test]
    fn parse_sub_has_no_return_type() {
        let result = parse("Sub Foo()\nEnd Sub\n");
        let proc = first_procedure(&result.ast);
        assert_eq!(proc.kind, ProcedureKind::Sub);
        assert!(
            proc.return_type.is_none(),
            "Sub should not have a return type, got {:?}",
            proc.return_type
        );
    }

    #[test]
    fn parse_function_without_return_type() {
        let result = parse("Function Foo()\nEnd Function\n");
        let proc = first_procedure(&result.ast);
        assert_eq!(proc.kind, ProcedureKind::Function);
        assert!(
            proc.return_type.is_none(),
            "legacy Function without As should have no return_type, got {:?}",
            proc.return_type
        );
    }

    #[test]
    fn parse_function_multiline_return_type() {
        let source = "Function Foo() _\n    As String\nEnd Function\n";
        let result = parse(source);
        let proc = first_procedure(&result.ast);
        assert_eq!(
            proc.return_type.as_ref().map(|s| s.as_str()),
            Some("String"),
            "expected return type captured across line continuation"
        );
    }

    #[test]
    fn parse_captures_if_statement_header() {
        let result = parse("Sub F()\n    If x > 0 Then\n    End If\nEnd Sub\n");
        let proc = first_procedure(&result.ast);
        assert!(!proc.body.is_empty(), "expected at least one statement");
        match statement(&result.ast, proc.body[0]) {
            StatementNode::If(node) => {
                assert!(!node.tokens.is_empty(), "expected If header tokens");
                assert!(
                    node.tokens.iter().any(|t| t.token == Token::If),
                    "expected the If keyword inside captured header tokens"
                );
            }
            other => panic!("expected If statement, got {:?}", other),
        }
    }

    #[test]
    fn parse_captures_for_statement_header() {
        let result = parse("Sub F()\n    For i = 1 To 10\n    Next i\nEnd Sub\n");
        let proc = first_procedure(&result.ast);
        assert!(!proc.body.is_empty(), "expected at least one statement");
        match statement(&result.ast, proc.body[0]) {
            StatementNode::For(node) => {
                assert!(
                    node.tokens.iter().any(|t| t.token == Token::For),
                    "expected the For keyword inside captured header tokens"
                );
            }
            other => panic!("expected For statement, got {:?}", other),
        }
    }

    #[test]
    fn parse_captures_with_statement_header() {
        let result = parse("Sub F()\n    With obj\n    End With\nEnd Sub\n");
        let proc = first_procedure(&result.ast);
        assert!(!proc.body.is_empty(), "expected at least one statement");
        match statement(&result.ast, proc.body[0]) {
            StatementNode::With(node) => {
                assert!(
                    node.tokens.iter().any(|t| t.token == Token::With),
                    "expected the With keyword inside captured header tokens"
                );
            }
            other => panic!("expected With statement, got {:?}", other),
        }
    }

    #[test]
    fn parse_captures_select_statement_header() {
        let result = parse("Sub F()\n    Select Case x\n    End Select\nEnd Sub\n");
        let proc = first_procedure(&result.ast);
        assert!(!proc.body.is_empty(), "expected at least one statement");
        match statement(&result.ast, proc.body[0]) {
            StatementNode::Select(node) => {
                assert!(
                    node.tokens.iter().any(|t| t.token == Token::Select),
                    "expected the Select keyword inside captured header tokens"
                );
            }
            other => panic!("expected Select statement, got {:?}", other),
        }
    }

    #[test]
    fn parse_captures_call_statement() {
        let result = parse("Sub F()\n    Call DoStuff(x)\nEnd Sub\n");
        let proc = first_procedure(&result.ast);
        assert_eq!(proc.body.len(), 1, "expected 1 body statement");
        match statement(&result.ast, proc.body[0]) {
            StatementNode::Call(node) => {
                assert!(
                    node.tokens.iter().any(|t| t.token == Token::Call),
                    "expected the Call keyword inside captured tokens"
                );
                assert!(
                    node.tokens
                        .iter()
                        .any(|t| t.token == Token::Identifier && t.text.as_str() == "DoStuff"),
                    "expected the called identifier inside captured tokens"
                );
            }
            other => panic!("expected Call statement, got {:?}", other),
        }
    }

    #[test]
    fn parse_captures_set_statement() {
        let result = parse("Sub F()\n    Set obj = New Foo\nEnd Sub\n");
        let proc = first_procedure(&result.ast);
        assert_eq!(proc.body.len(), 1, "expected 1 body statement");
        match statement(&result.ast, proc.body[0]) {
            StatementNode::Set(node) => {
                assert!(
                    node.tokens.iter().any(|t| t.token == Token::Set),
                    "expected the Set keyword inside captured tokens"
                );
                assert!(
                    node.tokens
                        .iter()
                        .any(|t| t.token == Token::Identifier && t.text.as_str() == "obj"),
                    "expected the target identifier inside captured tokens"
                );
            }
            other => panic!("expected Set statement, got {:?}", other),
        }
    }

    #[test]
    fn parse_captures_while_statement_header() {
        let result = parse("Sub F()\n    While x > 0\n    Wend\nEnd Sub\n");
        let proc = first_procedure(&result.ast);
        assert!(!proc.body.is_empty(), "expected at least one statement");
        match statement(&result.ast, proc.body[0]) {
            StatementNode::While(node) => {
                assert!(
                    node.tokens.iter().any(|t| t.token == Token::While),
                    "expected the While keyword inside captured header tokens"
                );
            }
            other => panic!("expected While statement, got {:?}", other),
        }
    }

    #[test]
    fn parse_captures_do_while_statement_header() {
        let result = parse("Sub F()\n    Do While x > 0\n    Loop\nEnd Sub\n");
        let proc = first_procedure(&result.ast);
        assert!(!proc.body.is_empty(), "expected at least one statement");
        match statement(&result.ast, proc.body[0]) {
            StatementNode::Do(node) => {
                assert!(
                    node.tokens.iter().any(|t| t.token == Token::Do),
                    "expected the Do keyword inside captured header tokens"
                );
                assert!(
                    node.tokens.iter().any(|t| t.token == Token::While),
                    "expected the While keyword inside captured Do header tokens"
                );
            }
            other => panic!("expected Do statement, got {:?}", other),
        }
    }

    #[test]
    fn parse_captures_do_until_statement_header() {
        let result = parse("Sub F()\n    Do Until x = 0\n    Loop\nEnd Sub\n");
        let proc = first_procedure(&result.ast);
        assert!(!proc.body.is_empty(), "expected at least one statement");
        match statement(&result.ast, proc.body[0]) {
            StatementNode::Do(node) => {
                assert!(
                    node.tokens.iter().any(|t| t.token == Token::Until),
                    "expected the Until keyword inside captured Do header tokens"
                );
            }
            other => panic!("expected Do statement, got {:?}", other),
        }
    }

    #[test]
    fn parse_captures_redim_statement_header() {
        let result = parse("Sub F()\n    ReDim arr(10)\nEnd Sub\n");
        let proc = first_procedure(&result.ast);
        assert!(!proc.body.is_empty(), "expected at least one statement");
        match statement(&result.ast, proc.body[0]) {
            StatementNode::Redim(node) => {
                assert!(
                    node.tokens.iter().any(|t| t.token == Token::ReDim),
                    "expected the ReDim keyword inside captured tokens"
                );
            }
            other => panic!("expected Redim statement, got {:?}", other),
        }
    }

    #[test]
    fn parse_captures_on_error_goto_as_on_error_statement() {
        let result = parse("Sub F()\n    On Error GoTo ErrHandler\nEnd Sub\n");
        let proc = first_procedure(&result.ast);
        assert_eq!(proc.body.len(), 1, "expected 1 body statement");
        match statement(&result.ast, proc.body[0]) {
            StatementNode::OnError(node) => {
                assert!(
                    node.tokens.iter().any(|t| t.token == Token::On),
                    "expected the On keyword inside captured tokens"
                );
            }
            other => panic!("expected OnError statement, got {:?}", other),
        }
    }

    #[test]
    fn parse_captures_goto_as_goto_statement() {
        let result = parse("Sub F()\n    GoTo MyLabel\nEnd Sub\n");
        let proc = first_procedure(&result.ast);
        assert_eq!(proc.body.len(), 1, "expected 1 body statement");
        match statement(&result.ast, proc.body[0]) {
            StatementNode::GoTo(node) => {
                assert!(
                    node.tokens.iter().any(|t| t.token == Token::GoTo),
                    "expected the GoTo keyword inside captured tokens"
                );
            }
            other => panic!("expected GoTo statement, got {:?}", other),
        }
    }

    #[test]
    fn parse_captures_exit_sub_as_exit_statement() {
        let result = parse("Sub F()\n    Exit Sub\nEnd Sub\n");
        let proc = first_procedure(&result.ast);
        assert_eq!(proc.body.len(), 1, "expected 1 body statement");
        match statement(&result.ast, proc.body[0]) {
            StatementNode::Exit(node) => {
                assert!(
                    node.tokens.iter().any(|t| t.token == Token::Exit),
                    "expected the Exit keyword inside captured tokens"
                );
            }
            other => panic!("expected Exit statement, got {:?}", other),
        }
    }
}
