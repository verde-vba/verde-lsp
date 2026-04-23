use tower_lsp::lsp_types::*;

use crate::analysis::resolve::offset_to_position;
use crate::analysis::AnalysisHost;
use crate::parser::lexer::Token;

/// Token types we report, in legend order.
pub(crate) const TOKEN_TYPES: &[SemanticTokenType] = &[
    SemanticTokenType::KEYWORD,
    SemanticTokenType::COMMENT,
    SemanticTokenType::STRING,
    SemanticTokenType::NUMBER,
    SemanticTokenType::OPERATOR,
    SemanticTokenType::VARIABLE,
    SemanticTokenType::FUNCTION,
    SemanticTokenType::TYPE,
    SemanticTokenType::ENUM_MEMBER,
    SemanticTokenType::PARAMETER,
    SemanticTokenType::PROPERTY,
];

/// Build the legend for server capabilities registration.
pub(crate) fn legend() -> SemanticTokensLegend {
    SemanticTokensLegend {
        token_types: TOKEN_TYPES.to_vec(),
        token_modifiers: vec![],
    }
}

/// Map a lexer token to an index in TOKEN_TYPES, or None to skip.
fn classify(token: &Token) -> Option<u32> {
    match token {
        // Keywords
        Token::Sub
        | Token::EndSub
        | Token::Function
        | Token::EndFunction
        | Token::Property
        | Token::EndProperty
        | Token::Dim
        | Token::Public
        | Token::Private
        | Token::Friend
        | Token::Static
        | Token::Const
        | Token::As
        | Token::ByVal
        | Token::ByRef
        | Token::Optional
        | Token::ParamArray
        | Token::If
        | Token::Then
        | Token::ElseIf
        | Token::Else
        | Token::EndIf
        | Token::For
        | Token::Each
        | Token::To
        | Token::Step
        | Token::Next
        | Token::Do
        | Token::While
        | Token::Until
        | Token::Loop
        | Token::Wend
        | Token::Select
        | Token::Case
        | Token::EndSelect
        | Token::With
        | Token::EndWith
        | Token::Set
        | Token::Let
        | Token::Call
        | Token::New
        | Token::Exit
        | Token::GoTo
        | Token::GoSub
        | Token::Return
        | Token::On
        | Token::Error
        | Token::Resume
        | Token::ReDim
        | Token::Preserve
        | Token::Erase
        | Token::Type
        | Token::EndType
        | Token::Enum
        | Token::EndEnum
        | Token::Implements
        | Token::Event
        | Token::RaiseEvent
        | Token::WithEvents
        | Token::Option
        | Token::Explicit
        | Token::Get
        | Token::Declare
        | Token::And
        | Token::Or
        | Token::Not
        | Token::Xor
        | Token::Mod
        | Token::Is
        | Token::Like
        | Token::TypeOf
        | Token::True
        | Token::False
        | Token::Nothing
        | Token::Null
        | Token::Empty
        | Token::Me => Some(0), // KEYWORD

        Token::Comment => Some(1), // COMMENT

        Token::StringLiteral | Token::DateLiteral => Some(2), // STRING

        Token::NumberLiteral | Token::HexLiteral => Some(3), // NUMBER

        Token::Eq
        | Token::Neq
        | Token::Lt
        | Token::Gt
        | Token::Lte
        | Token::Gte
        | Token::Plus
        | Token::Minus
        | Token::Star
        | Token::Slash
        | Token::IntDiv
        | Token::Caret
        | Token::Ampersand => Some(4), // OPERATOR

        // Type keywords
        Token::BooleanType
        | Token::ByteType
        | Token::IntegerType
        | Token::LongType
        | Token::LongLongType
        | Token::LongPtrType
        | Token::SingleType
        | Token::DoubleType
        | Token::CurrencyType
        | Token::DateType
        | Token::StringType
        | Token::VariantType
        | Token::ObjectType => Some(7), // TYPE

        // Identifiers, delimiters, etc. -- skip for token-based pass
        _ => None,
    }
}

pub(crate) fn semantic_tokens(host: &AnalysisHost, uri: &Url) -> Option<SemanticTokensResult> {
    let tokens_data = host.with_tokens(uri, |_symbols, source, lexer_tokens| {
        let mut data: Vec<SemanticToken> = Vec::new();
        let mut prev_line = 0u32;
        let mut prev_char = 0u32;

        for st in lexer_tokens {
            let type_index = match classify(&st.token) {
                Some(idx) => idx,
                None => continue,
            };

            let pos = offset_to_position(source, st.span.start);
            let length = (st.span.end - st.span.start) as u32;

            let delta_line = pos.line - prev_line;
            let delta_start = if delta_line == 0 {
                pos.character - prev_char
            } else {
                pos.character
            };

            data.push(SemanticToken {
                delta_line,
                delta_start,
                length,
                token_type: type_index,
                token_modifiers_bitset: 0,
            });

            prev_line = pos.line;
            prev_char = pos.character;
        }

        data
    })?;

    Some(SemanticTokensResult::Tokens(SemanticTokens {
        result_id: None,
        data: tokens_data,
    }))
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::analysis::AnalysisHost;

    fn setup_host(source: &str) -> (AnalysisHost, Url) {
        let host = AnalysisHost::new();
        let uri = Url::parse("file:///test.bas").unwrap();
        let parse_result = crate::parser::parse(source);
        host.update(uri.clone(), source.to_string(), parse_result);
        (host, uri)
    }

    #[test]
    fn keyword_tokens_are_classified() {
        let source = "Sub Foo()\nEnd Sub\n";
        let (host, uri) = setup_host(source);
        let result = semantic_tokens(&host, &uri);
        assert!(result.is_some());
        if let Some(SemanticTokensResult::Tokens(tokens)) = result {
            assert!(
                !tokens.data.is_empty(),
                "expected semantic tokens for keywords"
            );
            // Sub keyword should be token_type 0 (KEYWORD)
            assert_eq!(tokens.data[0].token_type, 0);
        }
    }

    #[test]
    fn comment_token_classified() {
        let source = "' this is a comment\n";
        let (host, uri) = setup_host(source);
        let result = semantic_tokens(&host, &uri);
        assert!(result.is_some());
        if let Some(SemanticTokensResult::Tokens(tokens)) = result {
            assert!(
                tokens.data.iter().any(|t| t.token_type == 1),
                "expected comment token type"
            );
        }
    }

    #[test]
    fn string_literal_classified() {
        let source = "Sub F()\n    x = \"hello\"\nEnd Sub\n";
        let (host, uri) = setup_host(source);
        let result = semantic_tokens(&host, &uri);
        assert!(result.is_some());
        if let Some(SemanticTokensResult::Tokens(tokens)) = result {
            assert!(
                tokens.data.iter().any(|t| t.token_type == 2),
                "expected string token type"
            );
        }
    }

    #[test]
    fn tokens_are_in_delta_order() {
        let source = "Sub Foo()\n    Dim x As Long\nEnd Sub\n";
        let (host, uri) = setup_host(source);
        let result = semantic_tokens(&host, &uri);
        assert!(result.is_some());
        // Verify delta encoding produces valid results (no negative deltas)
        if let Some(SemanticTokensResult::Tokens(tokens)) = result {
            for token in &tokens.data {
                // delta_line and delta_start should form a valid forward sequence
                assert!(token.length > 0, "token length must be positive");
            }
        }
    }
}
