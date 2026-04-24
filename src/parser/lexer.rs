use logos::Logos;

#[derive(Logos, Debug, PartialEq, Clone, Copy)]
#[logos(skip r"[ \t\r]+")]
pub enum Token {
    // Keywords
    #[token("Sub", ignore(ascii_case))]
    Sub,
    #[token("End Sub", ignore(ascii_case))]
    EndSub,
    #[token("Function", ignore(ascii_case))]
    Function,
    #[token("End Function", ignore(ascii_case))]
    EndFunction,
    #[token("Property", ignore(ascii_case))]
    Property,
    #[token("End Property", ignore(ascii_case))]
    EndProperty,
    #[token("Dim", ignore(ascii_case))]
    Dim,
    #[token("Public", ignore(ascii_case))]
    Public,
    #[token("Private", ignore(ascii_case))]
    Private,
    #[token("Friend", ignore(ascii_case))]
    Friend,
    #[token("Static", ignore(ascii_case))]
    Static,
    #[token("Const", ignore(ascii_case))]
    Const,
    #[token("As", ignore(ascii_case))]
    As,
    #[token("ByVal", ignore(ascii_case))]
    ByVal,
    #[token("ByRef", ignore(ascii_case))]
    ByRef,
    #[token("Optional", ignore(ascii_case))]
    Optional,
    #[token("ParamArray", ignore(ascii_case))]
    ParamArray,
    #[token("If", ignore(ascii_case))]
    If,
    #[token("Then", ignore(ascii_case))]
    Then,
    #[token("ElseIf", ignore(ascii_case))]
    ElseIf,
    #[token("Else", ignore(ascii_case))]
    Else,
    #[token("End If", ignore(ascii_case))]
    EndIf,
    #[token("For", ignore(ascii_case))]
    For,
    #[token("Each", ignore(ascii_case))]
    Each,
    #[token("To", ignore(ascii_case))]
    To,
    #[token("Step", ignore(ascii_case))]
    Step,
    #[token("Next", ignore(ascii_case))]
    Next,
    #[token("Do", ignore(ascii_case))]
    Do,
    #[token("While", ignore(ascii_case))]
    While,
    #[token("Until", ignore(ascii_case))]
    Until,
    #[token("Loop", ignore(ascii_case))]
    Loop,
    #[token("Wend", ignore(ascii_case))]
    Wend,
    #[token("Select", ignore(ascii_case))]
    Select,
    #[token("Case", ignore(ascii_case))]
    Case,
    #[token("End Select", ignore(ascii_case))]
    EndSelect,
    #[token("With", ignore(ascii_case))]
    With,
    #[token("End With", ignore(ascii_case))]
    EndWith,
    #[token("Set", ignore(ascii_case))]
    Set,
    #[token("Let", ignore(ascii_case))]
    Let,
    #[token("Call", ignore(ascii_case))]
    Call,
    #[token("New", ignore(ascii_case))]
    New,
    #[token("Exit", ignore(ascii_case))]
    Exit,
    #[token("GoTo", ignore(ascii_case))]
    GoTo,
    #[token("GoSub", ignore(ascii_case))]
    GoSub,
    #[token("Return", ignore(ascii_case))]
    Return,
    #[token("On", ignore(ascii_case))]
    On,
    #[token("Error", ignore(ascii_case))]
    Error,
    #[token("Resume", ignore(ascii_case))]
    Resume,
    #[token("ReDim", ignore(ascii_case))]
    ReDim,
    #[token("Preserve", ignore(ascii_case))]
    Preserve,
    #[token("Erase", ignore(ascii_case))]
    Erase,
    #[token("Type", ignore(ascii_case))]
    Type,
    #[token("End Type", ignore(ascii_case))]
    EndType,
    #[token("Enum", ignore(ascii_case))]
    Enum,
    #[token("End Enum", ignore(ascii_case))]
    EndEnum,
    #[token("Implements", ignore(ascii_case))]
    Implements,
    #[token("Event", ignore(ascii_case))]
    Event,
    #[token("RaiseEvent", ignore(ascii_case))]
    RaiseEvent,
    #[token("WithEvents", ignore(ascii_case))]
    WithEvents,
    #[token("Option", ignore(ascii_case))]
    Option,
    #[token("Explicit", ignore(ascii_case))]
    Explicit,
    #[token("Get", ignore(ascii_case))]
    Get,
    #[token("Declare", ignore(ascii_case))]
    Declare,

    // File number prefix (e.g. `#fileNum` in `Open … As #fileNum`)
    #[token("#")]
    Hash,

    // Conditional compilation
    #[token("#If", ignore(ascii_case))]
    HashIf,
    #[token("#ElseIf", ignore(ascii_case))]
    HashElseIf,
    #[token("#Else", ignore(ascii_case))]
    HashElse,
    #[token("#End If", ignore(ascii_case))]
    HashEndIf,
    #[token("#Const", ignore(ascii_case))]
    HashConst,

    // Literals
    #[token("True", ignore(ascii_case))]
    True,
    #[token("False", ignore(ascii_case))]
    False,
    #[token("Nothing", ignore(ascii_case))]
    Nothing,
    #[token("Null", ignore(ascii_case))]
    Null,
    #[token("Empty", ignore(ascii_case))]
    Empty,
    #[token("Me", ignore(ascii_case))]
    Me,

    // Type keywords
    #[token("Boolean", ignore(ascii_case))]
    BooleanType,
    #[token("Byte", ignore(ascii_case))]
    ByteType,
    #[token("Integer", ignore(ascii_case))]
    IntegerType,
    #[token("Long", ignore(ascii_case))]
    LongType,
    #[token("LongLong", ignore(ascii_case))]
    LongLongType,
    #[token("LongPtr", ignore(ascii_case))]
    LongPtrType,
    #[token("Single", ignore(ascii_case))]
    SingleType,
    #[token("Double", ignore(ascii_case))]
    DoubleType,
    #[token("Currency", ignore(ascii_case))]
    CurrencyType,
    #[token("Date", ignore(ascii_case))]
    DateType,
    #[token("String", ignore(ascii_case))]
    StringType,
    #[token("Variant", ignore(ascii_case))]
    VariantType,
    #[token("Object", ignore(ascii_case))]
    ObjectType,

    // Operators
    #[token("And", ignore(ascii_case))]
    And,
    #[token("Or", ignore(ascii_case))]
    Or,
    #[token("Not", ignore(ascii_case))]
    Not,
    #[token("Xor", ignore(ascii_case))]
    Xor,
    #[token("Mod", ignore(ascii_case))]
    Mod,
    #[token("Is", ignore(ascii_case))]
    Is,
    #[token("Like", ignore(ascii_case))]
    Like,
    #[token("TypeOf", ignore(ascii_case))]
    TypeOf,

    #[token("=")]
    Eq,
    #[token("<>")]
    Neq,
    #[token("<")]
    Lt,
    #[token(">")]
    Gt,
    #[token("<=")]
    Lte,
    #[token(">=")]
    Gte,
    #[token("+")]
    Plus,
    #[token("-")]
    Minus,
    #[token("*")]
    Star,
    #[token("/")]
    Slash,
    #[token("\\")]
    IntDiv,
    #[token("^")]
    Caret,
    #[token("&")]
    Ampersand,

    // Delimiters
    #[token("(")]
    LParen,
    #[token(")")]
    RParen,
    #[token(",")]
    Comma,
    #[token(".")]
    Dot,
    #[token(":")]
    Colon,
    #[token(";")]
    Semicolon,
    #[regex(r"_[ \t]*\r?\n")]
    LineContinuation,

    // Literals
    #[regex(r"[0-9]+(\.[0-9]+)?([eE][+-]?[0-9]+)?")]
    NumberLiteral,
    #[regex(r#""[^"]*""#)]
    StringLiteral,
    #[regex(r"#[0-9]{1,2}/[0-9]{1,2}/[0-9]{2,4}#")]
    DateLiteral,
    #[regex(r"&[hH][0-9a-fA-F]+&?")]
    HexLiteral,

    // Identifier
    #[regex(r"[a-zA-Z_][a-zA-Z0-9_]*")]
    Identifier,

    // Comment
    #[regex(r"'[^\n]*")]
    Comment,

    // Attribute (module-level)
    #[token("Attribute", ignore(ascii_case))]
    Attribute,

    // Newline
    #[token("\n")]
    Newline,
}

#[derive(Debug, Clone, PartialEq)]
pub struct SpannedToken {
    pub token: Token,
    pub span: std::ops::Range<usize>,
    pub text: smol_str::SmolStr,
}

#[derive(Debug, Clone)]
pub struct LexError {
    pub span: std::ops::Range<usize>,
}

pub fn lex(source: &str) -> (Vec<SpannedToken>, Vec<LexError>) {
    let mut tokens = Vec::new();
    let mut errors = Vec::new();
    let mut lexer = Token::lexer(source);

    while let Some(result) = lexer.next() {
        match result {
            Ok(token) => {
                tokens.push(SpannedToken {
                    token,
                    span: lexer.span(),
                    text: smol_str::SmolStr::new(lexer.slice()),
                });
            }
            Err(()) => {
                errors.push(LexError { span: lexer.span() });
            }
        }
    }

    (tokens, errors)
}

#[cfg(test)]
mod tests {
    use super::*;

    fn token_types(source: &str) -> Vec<Token> {
        let (tokens, _) = lex(source);
        tokens.into_iter().map(|t| t.token).collect()
    }

    // ── Keywords ─────────────────────────────────────────────────────

    #[test]
    fn lex_keywords_case_insensitive() {
        assert_eq!(token_types("sub"), vec![Token::Sub]);
        assert_eq!(token_types("SUB"), vec![Token::Sub]);
        assert_eq!(token_types("Sub"), vec![Token::Sub]);
        assert_eq!(token_types("DIM"), vec![Token::Dim]);
        assert_eq!(token_types("function"), vec![Token::Function]);
    }

    #[test]
    fn lex_multi_word_tokens() {
        assert_eq!(token_types("End Sub"), vec![Token::EndSub]);
        assert_eq!(token_types("end sub"), vec![Token::EndSub]);
        assert_eq!(token_types("End Function"), vec![Token::EndFunction]);
        assert_eq!(token_types("End If"), vec![Token::EndIf]);
        assert_eq!(token_types("End Select"), vec![Token::EndSelect]);
        assert_eq!(token_types("End With"), vec![Token::EndWith]);
        assert_eq!(token_types("End Type"), vec![Token::EndType]);
        assert_eq!(token_types("End Enum"), vec![Token::EndEnum]);
    }

    #[test]
    fn lex_conditional_compilation() {
        assert_eq!(token_types("#If"), vec![Token::HashIf]);
        assert_eq!(token_types("#ElseIf"), vec![Token::HashElseIf]);
        assert_eq!(token_types("#Else"), vec![Token::HashElse]);
        assert_eq!(token_types("#End If"), vec![Token::HashEndIf]);
        assert_eq!(token_types("#Const"), vec![Token::HashConst]);
    }

    // ── Operators ────────────────────────────────────────────────────

    #[test]
    fn lex_operators() {
        assert_eq!(token_types("="), vec![Token::Eq]);
        assert_eq!(token_types("<>"), vec![Token::Neq]);
        assert_eq!(token_types("<="), vec![Token::Lte]);
        assert_eq!(token_types(">="), vec![Token::Gte]);
        assert_eq!(token_types("<"), vec![Token::Lt]);
        assert_eq!(token_types(">"), vec![Token::Gt]);
        assert_eq!(
            token_types("+ - * / \\ ^ &"),
            vec![
                Token::Plus,
                Token::Minus,
                Token::Star,
                Token::Slash,
                Token::IntDiv,
                Token::Caret,
                Token::Ampersand
            ]
        );
    }

    #[test]
    fn lex_word_operators() {
        assert_eq!(token_types("And"), vec![Token::And]);
        assert_eq!(token_types("Or"), vec![Token::Or]);
        assert_eq!(token_types("Not"), vec![Token::Not]);
        assert_eq!(token_types("Mod"), vec![Token::Mod]);
    }

    // ── Literals ─────────────────────────────────────────────────────

    #[test]
    fn lex_number_literals() {
        let types = token_types("42 3.14 1e10 2.5E-3");
        assert_eq!(
            types,
            vec![
                Token::NumberLiteral,
                Token::NumberLiteral,
                Token::NumberLiteral,
                Token::NumberLiteral,
            ]
        );
    }

    #[test]
    fn lex_string_literal() {
        let (tokens, _) = lex("\"hello world\"");
        assert_eq!(tokens.len(), 1);
        assert_eq!(tokens[0].token, Token::StringLiteral);
        assert_eq!(tokens[0].text.as_str(), "\"hello world\"");
    }

    #[test]
    fn lex_hex_literal() {
        let (tokens, _) = lex("&HFF &h10&");
        assert_eq!(tokens.len(), 2);
        assert_eq!(tokens[0].token, Token::HexLiteral);
        assert_eq!(tokens[1].token, Token::HexLiteral);
    }

    #[test]
    fn lex_date_literal() {
        let (tokens, _) = lex("#1/15/2024#");
        assert_eq!(tokens.len(), 1);
        assert_eq!(tokens[0].token, Token::DateLiteral);
    }

    // ── Identifiers ──────────────────────────────────────────────────

    #[test]
    fn lex_identifier() {
        let (tokens, _) = lex("myVar _private foo123");
        assert_eq!(tokens.len(), 3);
        assert!(tokens.iter().all(|t| t.token == Token::Identifier));
        assert_eq!(tokens[0].text.as_str(), "myVar");
        assert_eq!(tokens[1].text.as_str(), "_private");
        assert_eq!(tokens[2].text.as_str(), "foo123");
    }

    // ── Comments ─────────────────────────────────────────────────────

    #[test]
    fn lex_comment() {
        let (tokens, _) = lex("' this is a comment\n");
        assert_eq!(tokens.len(), 2);
        assert_eq!(tokens[0].token, Token::Comment);
        assert_eq!(tokens[1].token, Token::Newline);
    }

    #[test]
    fn lex_comment_preserves_text() {
        let (tokens, _) = lex("' hello world");
        assert_eq!(tokens[0].text.as_str(), "' hello world");
    }

    // ── Line continuation ────────────────────────────────────────────

    #[test]
    fn lex_line_continuation() {
        let (tokens, _) = lex("x _\ny");
        assert_eq!(tokens.len(), 3);
        assert_eq!(tokens[0].token, Token::Identifier);
        assert_eq!(tokens[1].token, Token::LineContinuation);
        assert_eq!(tokens[2].token, Token::Identifier);
    }

    #[test]
    fn lex_line_continuation_crlf() {
        let (tokens, errors) = lex("x _\r\ny");
        assert!(errors.is_empty(), "expected no lex errors, got: {:?}", errors);
        assert_eq!(tokens.len(), 3);
        assert_eq!(tokens[0].token, Token::Identifier);
        assert_eq!(tokens[1].token, Token::LineContinuation);
        assert_eq!(tokens[2].token, Token::Identifier);
    }

    // ── Delimiters ───────────────────────────────────────────────────

    #[test]
    fn lex_delimiters() {
        assert_eq!(
            token_types("( ) , . : ;"),
            vec![
                Token::LParen,
                Token::RParen,
                Token::Comma,
                Token::Dot,
                Token::Colon,
                Token::Semicolon
            ]
        );
    }

    // ── Error handling ───────────────────────────────────────────────

    #[test]
    fn lex_error_on_invalid_character() {
        let (tokens, errors) = lex("x @ y");
        // @ is not a valid VBA token
        assert!(!errors.is_empty(), "expected a lex error for '@'");
        // x and y should still be lexed
        assert_eq!(tokens.len(), 2);
        assert_eq!(tokens[0].text.as_str(), "x");
        assert_eq!(tokens[1].text.as_str(), "y");
    }

    #[test]
    fn lex_error_does_not_prevent_subsequent_tokens() {
        let (tokens, errors) = lex("Dim @ x As Long");
        assert!(!errors.is_empty());
        assert!(tokens.iter().any(|t| t.token == Token::Dim));
        assert!(tokens.iter().any(|t| t.token == Token::As));
        assert!(tokens.iter().any(|t| t.token == Token::LongType));
    }

    // ── Type keywords ────────────────────────────────────────────────

    #[test]
    fn lex_type_keywords() {
        assert_eq!(token_types("Boolean"), vec![Token::BooleanType]);
        assert_eq!(token_types("Integer"), vec![Token::IntegerType]);
        assert_eq!(token_types("Long"), vec![Token::LongType]);
        assert_eq!(token_types("String"), vec![Token::StringType]);
        assert_eq!(token_types("Variant"), vec![Token::VariantType]);
        assert_eq!(token_types("Object"), vec![Token::ObjectType]);
    }

    // ── Span correctness ─────────────────────────────────────────────

    #[test]
    fn lex_hash_file_number() {
        let (tokens, errors) = lex("Open f For Append As #fileNum");
        assert!(errors.is_empty(), "expected no lex errors, got: {:?}", errors);
        assert!(
            tokens.iter().any(|t| t.token == Token::Hash && t.text.as_str() == "#"),
            "expected a Hash token for '#'"
        );
    }

    #[test]
    fn lex_span_covers_token_text() {
        let source = "Dim x As Long";
        let (tokens, _) = lex(source);
        for t in &tokens {
            assert_eq!(
                &source[t.span.clone()],
                t.text.as_str(),
                "span mismatch for {:?}",
                t.token
            );
        }
    }

    // ── Full statement ───────────────────────────────────────────────

    #[test]
    fn lex_full_dim_statement() {
        let types = token_types("Dim x As Long\n");
        assert_eq!(
            types,
            vec![
                Token::Dim,
                Token::Identifier,
                Token::As,
                Token::LongType,
                Token::Newline
            ]
        );
    }

    #[test]
    fn lex_sub_declaration() {
        let types = token_types("Sub Foo()\n");
        assert_eq!(
            types,
            vec![
                Token::Sub,
                Token::Identifier,
                Token::LParen,
                Token::RParen,
                Token::Newline
            ]
        );
    }
}
