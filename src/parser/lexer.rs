use logos::Logos;

#[derive(Logos, Debug, PartialEq, Clone)]
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
    #[regex(r"_[ \t]*\n")]
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

pub fn lex(source: &str) -> Vec<SpannedToken> {
    let mut tokens = Vec::new();
    let mut lexer = Token::lexer(source);

    while let Some(result) = lexer.next() {
        if let Ok(token) = result {
            tokens.push(SpannedToken {
                token,
                span: lexer.span(),
                text: smol_str::SmolStr::new(lexer.slice()),
            });
        }
    }

    tokens
}
