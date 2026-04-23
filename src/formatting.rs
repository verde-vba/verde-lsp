use crate::parser::lexer::{lex, SpannedToken, Token};

pub fn apply_formatting(src: &str) -> String {
    // Phase 1: keyword case normalization (token gap preservation)
    let (tokens, _lex_errors) = lex(src);
    let mut keyword_normalized = String::with_capacity(src.len());
    let mut pos = 0usize;

    for st in &tokens {
        keyword_normalized.push_str(&src[pos..st.span.start]);
        match keyword_canonical(&st.token) {
            Some(canonical) => keyword_normalized.push_str(canonical),
            None => keyword_normalized.push_str(st.text.as_str()),
        }
        pos = st.span.end;
    }
    keyword_normalized.push_str(&src[pos..]);

    // Phase 2: indent normalization + trailing whitespace removal
    let indents = calculate_line_indents(&keyword_normalized);
    let trailing_newline = keyword_normalized.ends_with('\n');

    let lines: Vec<&str> = keyword_normalized.lines().collect();
    let mut formatted = String::with_capacity(src.len());

    for (i, line) in lines.iter().enumerate() {
        let depth = indents.get(i).copied().unwrap_or(0);
        let trimmed = line.trim();

        if i > 0 {
            formatted.push('\n');
        }

        if !trimmed.is_empty() {
            for _ in 0..depth * 4 {
                formatted.push(' ');
            }
            formatted.push_str(trimmed);
        }
    }

    if trailing_newline {
        formatted.push('\n');
    }

    formatted
}

/// Computes the indent depth (in units of 4 spaces) for each line in `src`.
///
/// Returns one `usize` per line (matching `src.lines()` iteration order).
/// Blank lines always get depth 0.
fn calculate_line_indents(src: &str) -> Vec<usize> {
    let mut depths = Vec::new();
    let mut depth: i32 = 0;

    for line in src.lines() {
        let trimmed = line.trim();

        if trimmed.is_empty() {
            depths.push(0);
            continue;
        }

        let (tokens, _) = lex(trimmed);
        let first = tokens.first().map(|st| &st.token);

        let line_depth = match first {
            Some(t) if is_close_token(t) => {
                depth = (depth - 1).max(0);
                depth as usize
            }
            Some(t) if is_midblock_token(t) => (depth - 1).max(0) as usize,
            _ => depth as usize,
        };

        depths.push(line_depth);

        // Increment depth AFTER printing for open tokens (skip access modifiers)
        if let Some(t) = first_block_token(&tokens) {
            if is_open_token(t) {
                depth += 1;
            }
        }
    }

    depths
}

/// Returns the first token that is not an access modifier.
///
/// Handles `Public Sub Foo()` / `Private Function Bar()` correctly by
/// skipping `Public`/`Private`/`Friend`/`Static` before the block keyword.
fn first_block_token(tokens: &[SpannedToken]) -> Option<&Token> {
    for st in tokens {
        match &st.token {
            Token::Public | Token::Private | Token::Friend | Token::Static => continue,
            t => return Some(t),
        }
    }
    None
}

fn is_open_token(token: &Token) -> bool {
    matches!(
        token,
        Token::Sub
            | Token::Function
            | Token::Property
            | Token::If
            | Token::For
            | Token::Do
            | Token::While
            | Token::Select
            | Token::With
            | Token::Type
            | Token::Enum
    )
}

fn is_close_token(token: &Token) -> bool {
    matches!(
        token,
        Token::EndSub
            | Token::EndFunction
            | Token::EndProperty
            | Token::EndIf
            | Token::Next
            | Token::Loop
            | Token::Wend
            | Token::EndSelect
            | Token::EndWith
            | Token::EndType
            | Token::EndEnum
    )
}

/// Midblock tokens are printed at `depth - 1` but do NOT change `depth`.
///
/// VBA convention: `ElseIf`/`Else` align with `If`; `Case` aligns with `Select Case`.
fn is_midblock_token(token: &Token) -> bool {
    matches!(token, Token::ElseIf | Token::Else | Token::Case)
}

fn keyword_canonical(token: &Token) -> Option<&'static str> {
    match token {
        Token::Sub => Some("Sub"),
        Token::EndSub => Some("End Sub"),
        Token::Function => Some("Function"),
        Token::EndFunction => Some("End Function"),
        Token::Property => Some("Property"),
        Token::EndProperty => Some("End Property"),
        Token::Dim => Some("Dim"),
        Token::Public => Some("Public"),
        Token::Private => Some("Private"),
        Token::Friend => Some("Friend"),
        Token::Static => Some("Static"),
        Token::Const => Some("Const"),
        Token::As => Some("As"),
        Token::ByVal => Some("ByVal"),
        Token::ByRef => Some("ByRef"),
        Token::Optional => Some("Optional"),
        Token::ParamArray => Some("ParamArray"),
        Token::If => Some("If"),
        Token::Then => Some("Then"),
        Token::ElseIf => Some("ElseIf"),
        Token::Else => Some("Else"),
        Token::EndIf => Some("End If"),
        Token::For => Some("For"),
        Token::Each => Some("Each"),
        Token::To => Some("To"),
        Token::Step => Some("Step"),
        Token::Next => Some("Next"),
        Token::Do => Some("Do"),
        Token::While => Some("While"),
        Token::Until => Some("Until"),
        Token::Loop => Some("Loop"),
        Token::Wend => Some("Wend"),
        Token::Select => Some("Select"),
        Token::Case => Some("Case"),
        Token::EndSelect => Some("End Select"),
        Token::With => Some("With"),
        Token::EndWith => Some("End With"),
        Token::Set => Some("Set"),
        Token::Let => Some("Let"),
        Token::Call => Some("Call"),
        Token::New => Some("New"),
        Token::Exit => Some("Exit"),
        Token::GoTo => Some("GoTo"),
        Token::GoSub => Some("GoSub"),
        Token::Return => Some("Return"),
        Token::On => Some("On"),
        Token::Error => Some("Error"),
        Token::Resume => Some("Resume"),
        Token::ReDim => Some("ReDim"),
        Token::Preserve => Some("Preserve"),
        Token::Erase => Some("Erase"),
        Token::Type => Some("Type"),
        Token::EndType => Some("End Type"),
        Token::Enum => Some("Enum"),
        Token::EndEnum => Some("End Enum"),
        Token::Implements => Some("Implements"),
        Token::Event => Some("Event"),
        Token::RaiseEvent => Some("RaiseEvent"),
        Token::WithEvents => Some("WithEvents"),
        Token::Option => Some("Option"),
        Token::Explicit => Some("Explicit"),
        Token::Get => Some("Get"),
        Token::True => Some("True"),
        Token::False => Some("False"),
        Token::Nothing => Some("Nothing"),
        Token::Null => Some("Null"),
        Token::Empty => Some("Empty"),
        Token::Me => Some("Me"),
        Token::BooleanType => Some("Boolean"),
        Token::ByteType => Some("Byte"),
        Token::IntegerType => Some("Integer"),
        Token::LongType => Some("Long"),
        Token::LongLongType => Some("LongLong"),
        Token::LongPtrType => Some("LongPtr"),
        Token::SingleType => Some("Single"),
        Token::DoubleType => Some("Double"),
        Token::CurrencyType => Some("Currency"),
        Token::DateType => Some("Date"),
        Token::StringType => Some("String"),
        Token::VariantType => Some("Variant"),
        Token::ObjectType => Some("Object"),
        Token::And => Some("And"),
        Token::Or => Some("Or"),
        Token::Not => Some("Not"),
        Token::Xor => Some("Xor"),
        Token::Mod => Some("Mod"),
        Token::Is => Some("Is"),
        Token::Like => Some("Like"),
        Token::TypeOf => Some("TypeOf"),
        Token::Attribute => Some("Attribute"),
        _ => None,
    }
}
