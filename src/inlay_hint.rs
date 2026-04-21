use tower_lsp::lsp_types::*;

use crate::analysis::resolve::offset_to_position;
use crate::analysis::symbols::{Symbol, SymbolKind, SymbolTable};

/// Compute inlay hints for all Variable and Constant symbols in the symbol table.
/// Hints are placed at the end of the symbol name span, showing the type.
pub fn inlay_hints(src: &str, table: &SymbolTable) -> Vec<InlayHint> {
    table
        .symbols
        .iter()
        .filter(|s| matches!(s.kind, SymbolKind::Variable | SymbolKind::Constant))
        .map(|s| build_hint(src, s))
        .collect()
}

fn build_hint(src: &str, symbol: &Symbol) -> InlayHint {
    let type_label = symbol.type_name.as_deref().unwrap_or("Variant");
    let label = format!(": {}", type_label);
    let pos = offset_to_position(src, symbol.span.end as usize);
    InlayHint {
        position: pos,
        label: InlayHintLabel::String(label),
        kind: Some(InlayHintKind::TYPE),
        text_edits: None,
        tooltip: None,
        padding_left: Some(true),
        padding_right: None,
        data: None,
    }
}
