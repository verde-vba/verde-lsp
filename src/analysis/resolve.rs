use smol_str::SmolStr;
use tower_lsp::lsp_types::*;

use super::symbols::{Symbol, SymbolTable};
use crate::parser::ast::TextRange;

pub fn find_symbol_at_position<'a>(
    symbols: &'a SymbolTable,
    source: &str,
    position: Position,
) -> Option<&'a Symbol> {
    let offset = position_to_offset(source, position)?;
    symbols
        .symbols
        .iter()
        .find(|s| offset >= s.span.start as usize && offset <= s.span.end as usize)
}

pub fn find_symbol_by_name<'a>(symbols: &'a SymbolTable, name: &str) -> Vec<&'a Symbol> {
    symbols
        .symbols
        .iter()
        .filter(|s| s.name.eq_ignore_ascii_case(name))
        .collect()
}

/// Find every byte range in `source` where `word` appears as a standalone
/// identifier (word-boundary on both sides, case-insensitive). Used by rename
/// to collect all reference sites in addition to declaration sites.
pub fn find_all_word_occurrences(source: &str, word: &str) -> Vec<TextRange> {
    let bytes = source.as_bytes();
    let word_len = word.len();
    let mut result = Vec::new();
    let mut i = 0;

    while i + word_len <= bytes.len() {
        let prev_is_ident = i > 0 && is_ident_char(bytes[i - 1]);
        if !prev_is_ident && source[i..i + word_len].eq_ignore_ascii_case(word) {
            let next_is_ident = (i + word_len) < bytes.len() && is_ident_char(bytes[i + word_len]);
            if !next_is_ident {
                result.push(TextRange::new(i, i + word_len));
            }
        }
        i += 1;
    }

    result
}

pub fn find_word_at_position(source: &str, position: Position) -> Option<SmolStr> {
    let offset = position_to_offset(source, position)?;
    let bytes = source.as_bytes();

    let mut start = offset;
    while start > 0 && is_ident_char(bytes[start - 1]) {
        start -= 1;
    }

    let mut end = offset;
    while end < bytes.len() && is_ident_char(bytes[end]) {
        end += 1;
    }

    if start == end {
        return None;
    }

    Some(SmolStr::new(&source[start..end]))
}

pub fn text_range_to_lsp_range(source: &str, range: TextRange) -> Range {
    let start = offset_to_position(source, range.start as usize);
    let end = offset_to_position(source, range.end as usize);
    Range::new(start, end)
}

pub fn position_to_offset(source: &str, position: Position) -> Option<usize> {
    let mut line = 0u32;
    let mut col = 0u32;

    for (i, ch) in source.char_indices() {
        if line == position.line && col == position.character {
            return Some(i);
        }
        if ch == '\n' {
            if line == position.line {
                return Some(i);
            }
            line += 1;
            col = 0;
        } else {
            col += ch.len_utf16() as u32;
        }
    }

    if line == position.line {
        Some(source.len())
    } else {
        None
    }
}

pub fn offset_to_position(source: &str, offset: usize) -> Position {
    let mut line = 0u32;
    let mut col = 0u32;

    for (i, ch) in source.char_indices() {
        if i == offset {
            return Position::new(line, col);
        }
        if ch == '\n' {
            line += 1;
            col = 0;
        } else {
            col += ch.len_utf16() as u32;
        }
    }

    Position::new(line, col)
}

fn is_ident_char(b: u8) -> bool {
    b.is_ascii_alphanumeric() || b == b'_'
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn position_to_offset_counts_ascii_as_one_utf16_unit() {
        let source = "abc\ndef";
        assert_eq!(position_to_offset(source, Position::new(0, 0)), Some(0));
        assert_eq!(position_to_offset(source, Position::new(0, 2)), Some(2));
        assert_eq!(position_to_offset(source, Position::new(1, 1)), Some(5));
    }

    #[test]
    fn position_to_offset_counts_bmp_char_as_one_utf16_unit() {
        // 'あ' (U+3042) occupies 3 UTF-8 bytes and 1 UTF-16 code unit.
        let source = "あabc";
        assert_eq!(position_to_offset(source, Position::new(0, 1)), Some(3));
        assert_eq!(position_to_offset(source, Position::new(0, 2)), Some(4));
    }

    #[test]
    fn position_to_offset_counts_astral_char_as_two_utf16_units() {
        // '𝕏' (U+1D54F) occupies 4 UTF-8 bytes and 2 UTF-16 code units
        // (surrogate pair). LSP client sends UTF-16 offsets, so the column
        // for the following ASCII char must be 2, not 1.
        let source = "𝕏abc";
        assert_eq!(position_to_offset(source, Position::new(0, 2)), Some(4));
        assert_eq!(position_to_offset(source, Position::new(0, 3)), Some(5));
    }

    #[test]
    fn text_range_to_lsp_range_emits_utf16_columns_for_astral_char() {
        let source = "𝕏abc";
        let range = text_range_to_lsp_range(source, TextRange::new(4, 5));
        assert_eq!(range.start, Position::new(0, 2));
        assert_eq!(range.end, Position::new(0, 3));
    }
}
