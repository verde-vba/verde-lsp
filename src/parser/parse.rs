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
                    params: Vec::new(),
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
