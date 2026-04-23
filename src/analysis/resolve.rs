use smol_str::SmolStr;
use tower_lsp::lsp_types::*;

use super::symbols::{Symbol, SymbolTable};
use crate::parser::ast::TextRange;

pub(crate) fn find_symbol_at_position<'a>(
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

pub(crate) fn find_symbol_by_name<'a>(symbols: &'a SymbolTable, name: &str) -> Vec<&'a Symbol> {
    symbols
        .symbols
        .iter()
        .filter(|s| s.name.eq_ignore_ascii_case(name))
        .collect()
}

/// Find every byte range in `source` where `word` appears as a standalone
/// identifier (word-boundary on both sides, case-insensitive). Used by rename
/// to collect all reference sites in addition to declaration sites.
pub(crate) fn find_all_word_occurrences(source: &str, word: &str) -> Vec<TextRange> {
    let bytes = source.as_bytes();
    let word_len = word.len();
    let mut result = Vec::new();
    let mut i = 0;

    while i + word_len <= bytes.len() {
        let prev_is_ident = i > 0 && is_ident_char(bytes[i - 1]);
        if !prev_is_ident && bytes[i..i + word_len].eq_ignore_ascii_case(word.as_bytes()) {
            let next_is_ident = (i + word_len) < bytes.len() && is_ident_char(bytes[i + word_len]);
            if !next_is_ident {
                result.push(TextRange::new(i, i + word_len));
            }
        }
        i += 1;
    }

    result
}

pub(crate) fn find_word_at_position(source: &str, position: Position) -> Option<SmolStr> {
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

pub(crate) fn text_range_to_lsp_range(source: &str, range: TextRange) -> Range {
    let start = offset_to_position(source, range.start as usize);
    let end = offset_to_position(source, range.end as usize);
    Range::new(start, end)
}

pub(crate) fn position_to_offset(source: &str, position: Position) -> Option<usize> {
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

pub(crate) fn offset_to_position(source: &str, offset: usize) -> Position {
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

/// Detect dot-access at cursor position.
///
/// Returns `(var_name, member_partial)` where `member_partial` is the
/// identifier text from its start up to the cursor offset (may be empty
/// directly after the dot, e.g. `f.` with cursor right after the dot).
pub(crate) fn parse_dot_access_at(
    source: &str,
    cursor_offset: usize,
) -> Option<(SmolStr, SmolStr)> {
    let bytes = source.as_bytes();

    // Walk backward from cursor to the start of the current identifier token.
    let mut member_start = cursor_offset;
    while member_start > 0 && is_ident_char(bytes[member_start - 1]) {
        member_start -= 1;
    }

    // The character immediately before the identifier must be `.`.
    if member_start == 0 || bytes[member_start - 1] != b'.' {
        return None;
    }

    let dot_pos = member_start - 1;
    let member_partial = SmolStr::new(&source[member_start..cursor_offset]);

    // Walk backward from the dot to find the variable identifier.
    let mut var_start = dot_pos;
    while var_start > 0 && is_ident_char(bytes[var_start - 1]) {
        var_start -= 1;
    }

    if var_start == dot_pos {
        return None; // nothing before dot
    }

    let var_name = SmolStr::new(&source[var_start..dot_pos]);
    Some((var_name, member_partial))
}

/// Detect a leading dot (`.member` without preceding identifier).
/// Returns the member partial text if cursor is at a leading-dot position.
pub(crate) fn parse_leading_dot_at(source: &str, cursor_offset: usize) -> Option<SmolStr> {
    let bytes = source.as_bytes();
    let mut member_start = cursor_offset;
    while member_start > 0 && is_ident_char(bytes[member_start - 1]) {
        member_start -= 1;
    }
    if member_start == 0 || bytes[member_start - 1] != b'.' {
        return None;
    }
    let dot_pos = member_start - 1;
    // Check there's no identifier immediately before the dot
    if dot_pos > 0 && is_ident_char(bytes[dot_pos - 1]) {
        return None;
    }
    Some(SmolStr::new(&source[member_start..cursor_offset]))
}

/// Parse a dot chain backwards from cursor. For `a.b.c.|` returns
/// `(["a", "b", "c"], member_partial)`. For single `a.|` returns `(["a"], "")`.
pub(crate) fn parse_dot_chain_at(
    source: &str,
    cursor_offset: usize,
) -> Option<(Vec<SmolStr>, SmolStr)> {
    let bytes = source.as_bytes();

    let mut member_start = cursor_offset;
    while member_start > 0 && is_ident_char(bytes[member_start - 1]) {
        member_start -= 1;
    }
    if member_start == 0 || bytes[member_start - 1] != b'.' {
        return None;
    }
    let member_partial = SmolStr::new(&source[member_start..cursor_offset]);

    let mut chain = Vec::new();
    let mut pos = member_start - 1; // at the dot

    loop {
        let ident_end = pos;
        let mut ident_start = ident_end;
        while ident_start > 0 && is_ident_char(bytes[ident_start - 1]) {
            ident_start -= 1;
        }
        if ident_start == ident_end {
            break; // leading dot or no identifier
        }
        chain.push(SmolStr::new(&source[ident_start..ident_end]));
        if ident_start > 0 && bytes[ident_start - 1] == b'.' {
            pos = ident_start - 1;
        } else {
            break;
        }
    }

    chain.reverse();
    if chain.is_empty() {
        return None;
    }
    Some((chain, member_partial))
}

/// Detect `FuncName().` pattern. Returns the function name and member partial
/// if cursor is after a closing paren followed by a dot.
pub(crate) fn parse_func_call_dot_at(
    source: &str,
    cursor_offset: usize,
) -> Option<(SmolStr, SmolStr)> {
    let bytes = source.as_bytes();

    let mut member_start = cursor_offset;
    while member_start > 0 && is_ident_char(bytes[member_start - 1]) {
        member_start -= 1;
    }
    if member_start == 0 || bytes[member_start - 1] != b'.' {
        return None;
    }
    let member_partial = SmolStr::new(&source[member_start..cursor_offset]);
    let dot_pos = member_start - 1;

    // Check if `)` is before the dot
    if dot_pos == 0 || bytes[dot_pos - 1] != b')' {
        return None;
    }

    // Find matching `(`
    let mut paren_pos = dot_pos - 1; // at `)`
    let mut depth = 1i32;
    while paren_pos > 0 && depth > 0 {
        paren_pos -= 1;
        match bytes[paren_pos] {
            b')' => depth += 1,
            b'(' => depth -= 1,
            _ => {}
        }
    }
    if depth != 0 {
        return None;
    }

    // paren_pos is at `(`. Read identifier before it.
    let mut func_end = paren_pos;
    while func_end > 0 && bytes[func_end - 1] == b' ' {
        func_end -= 1;
    }
    let mut func_start = func_end;
    while func_start > 0 && is_ident_char(bytes[func_start - 1]) {
        func_start -= 1;
    }
    if func_start == func_end {
        return None;
    }

    let func_name = SmolStr::new(&source[func_start..func_end]);
    Some((func_name, member_partial))
}

/// Find the procedure name that contains the given byte offset.
pub(crate) fn find_containing_proc(
    proc_ranges: &[(SmolStr, TextRange)],
    offset: usize,
) -> Option<&SmolStr> {
    proc_ranges
        .iter()
        .find(|(_, r)| offset >= r.start as usize && offset <= r.end as usize)
        .map(|(name, _)| name)
}

/// Resolve the type of a variable at a given offset. Checks proc-scoped symbols first,
/// then module-level symbols.
pub(crate) fn resolve_var_type_at(
    symbols: &SymbolTable,
    offset: usize,
    var_name: &str,
) -> Option<SmolStr> {
    let cursor_proc = find_containing_proc(&symbols.proc_ranges, offset);

    cursor_proc
        .and_then(|proc_name| {
            symbols.symbols.iter().find(|s| {
                s.name.eq_ignore_ascii_case(var_name)
                    && matches!(
                        s.kind,
                        super::symbols::SymbolKind::Variable
                            | super::symbols::SymbolKind::Parameter
                    )
                    && s.proc_scope
                        .as_ref()
                        .is_some_and(|p| p.eq_ignore_ascii_case(proc_name))
            })
        })
        .or_else(|| {
            symbols.symbols.iter().find(|s| {
                s.name.eq_ignore_ascii_case(var_name)
                    && matches!(
                        s.kind,
                        super::symbols::SymbolKind::Variable
                            | super::symbols::SymbolKind::Parameter
                    )
                    && s.proc_scope.is_none()
            })
        })
        .and_then(|s| s.type_name.clone())
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

    #[test]
    fn find_all_word_occurrences_handles_multibyte_source() {
        // Japanese comments must not cause a panic when scanning for ASCII words.
        let source = "' ワークブック\nDim x As Long\nx = 1\n' ユーティリティ\nx = x + 1";
        let hits = find_all_word_occurrences(source, "x");
        assert_eq!(hits.len(), 4);
    }
}
