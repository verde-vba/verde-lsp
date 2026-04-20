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

    fn advance(&mut self) -> Option<&lexer::SpannedToken> {
        let tok = self.tokens.get(self.pos);
        if tok.is_some() {
            self.pos += 1;
        }
        tok
    }

    fn skip_newlines(&mut self) {
        while matches!(self.peek(), Some(t) if t.token == Token::Newline || t.token == Token::Comment) {
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

        // Parse the parameter list `(...)` if present. VBA allows `Sub Foo`
        // without parentheses for a no-arg procedure, so we only descend into
        // parameter parsing when an opening paren is the next token.
        let params = if matches!(self.peek(), Some(t) if t.token == Token::LParen) {
            self.parse_parameter_list()
        } else {
            Vec::new()
        };

        // Skip to end of procedure
        let end_token = match kind {
            ProcedureKind::Sub => Token::EndSub,
            ProcedureKind::Function => Token::EndFunction,
            ProcedureKind::PropertyGet | ProcedureKind::PropertyLet | ProcedureKind::PropertySet => {
                Token::EndProperty
            }
        };

        while self.pos < self.tokens.len() {
            if self.tokens[self.pos].token == end_token {
                let end = self.tokens[self.pos].span.end;
                self.pos += 1;

                let node = AstNode::Procedure(ProcedureNode {
                    name,
                    kind,
                    visibility,
                    params,
                    return_type: None,
                    body: Vec::new(),
                    span: TextRange::new(start, end),
                });
                let id = self.ast.nodes.alloc(node);
                self.ast.root.push(id);
                return;
            }
            self.pos += 1;
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
            while matches!(
                self.peek(),
                Some(t) if t.token == Token::LineContinuation || t.token == Token::Newline
            ) {
                self.pos += 1;
            }

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
            // else defensively until we find one.
            loop {
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

        // Optional `As <Type>` clause.
        let mut type_name = None;
        if matches!(self.peek(), Some(t) if t.token == Token::As) {
            self.pos += 1;
            if let Some(t) = self.peek() {
                // Accept any type-ish token (builtin type keyword or identifier).
                // For the AST we only record identifier-like names; a builtin
                // type token's `text` still holds the source text.
                type_name = Some(t.text.clone());
                end_offset = t.span.end;
                self.pos += 1;
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

    fn parameter<'a>(ast: &'a Ast, id: NodeId) -> &'a ParameterNode {
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

    #[test]
    fn parse_procedure_records_body_range() {
        let source = "Sub Foo()\n    x = 1\nEnd Sub\n";
        let result = parse(source);
        let proc = first_procedure(&result.ast);
        let body = &source[proc.body_range.start as usize..proc.body_range.end as usize];
        assert!(
            body.starts_with("    x = 1"),
            "expected body to start with the body line, got {:?}",
            body
        );
        let end_idx = source.find("End Sub").expect("End Sub must exist");
        assert_eq!(
            proc.body_range.end as usize, end_idx,
            "body_range.end should point at the start of End Sub"
        );
    }

    #[test]
    fn parse_procedure_body_range_is_empty_when_no_body() {
        let source = "Sub Foo()\nEnd Sub\n";
        let result = parse(source);
        let proc = first_procedure(&result.ast);
        let body = &source[proc.body_range.start as usize..proc.body_range.end as usize];
        assert!(
            body.trim().is_empty(),
            "expected empty/whitespace-only body, got {:?}",
            body
        );
    }

    #[test]
    fn parse_procedure_body_range_excludes_signature() {
        let source = "Sub Foo(x As Long)\n    y = x\nEnd Sub\n";
        let result = parse(source);
        let proc = first_procedure(&result.ast);
        let body = &source[proc.body_range.start as usize..proc.body_range.end as usize];
        assert!(
            !body.contains("Sub Foo"),
            "body_range should not contain 'Sub Foo', got {:?}",
            body
        );
        assert!(
            !body.contains("x As Long"),
            "body_range should not contain 'x As Long', got {:?}",
            body
        );
    }
}
