use crate::parser::lexer::{lex, Token};

pub fn apply_formatting(src: &str) -> String {
    let tokens = lex(src);
    let mut result = String::with_capacity(src.len());
    let mut pos = 0usize;

    for st in &tokens {
        result.push_str(&src[pos..st.span.start]);
        match keyword_canonical(&st.token) {
            Some(canonical) => result.push_str(canonical),
            None => result.push_str(st.text.as_str()),
        }
        pos = st.span.end;
    }
    result.push_str(&src[pos..]);

    let trailing_newline = result.ends_with('\n');
    let mut formatted: String = result
        .lines()
        .map(str::trim_end)
        .collect::<Vec<_>>()
        .join("\n");
    if trailing_newline {
        formatted.push('\n');
    }
    formatted
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
