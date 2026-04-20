use tower_lsp::lsp_types::*;

use crate::analysis::resolve::position_to_offset;
use crate::analysis::symbols::{SymbolKind, SymbolTable};
use crate::analysis::AnalysisHost;
use crate::vba_builtins;

fn symbol_kind_to_completion_kind(kind: &SymbolKind) -> CompletionItemKind {
    match kind {
        SymbolKind::Procedure => CompletionItemKind::METHOD,
        SymbolKind::Function => CompletionItemKind::FUNCTION,
        SymbolKind::Property => CompletionItemKind::PROPERTY,
        SymbolKind::Variable => CompletionItemKind::VARIABLE,
        SymbolKind::Constant => CompletionItemKind::CONSTANT,
        SymbolKind::Parameter => CompletionItemKind::VARIABLE,
        SymbolKind::TypeDef => CompletionItemKind::STRUCT,
        SymbolKind::EnumDef => CompletionItemKind::ENUM,
        SymbolKind::EnumMember => CompletionItemKind::ENUM_MEMBER,
    }
}

pub fn complete(host: &AnalysisHost, uri: &Url, position: Position) -> Vec<CompletionItem> {
    let mut items = Vec::new();

    // VBA keywords
    for kw in vba_builtins::KEYWORDS {
        items.push(CompletionItem {
            label: kw.to_string(),
            kind: Some(CompletionItemKind::KEYWORD),
            ..Default::default()
        });
    }

    // VBA built-in functions
    for func in vba_builtins::BUILTIN_FUNCTIONS {
        items.push(CompletionItem {
            label: func.to_string(),
            kind: Some(CompletionItemKind::FUNCTION),
            ..Default::default()
        });
    }

    // Symbols from current file, filtered by cursor scope
    if let Some(sym_items) = host.with_source(uri, |symbols, source| {
        let cursor_scope = proc_at_position(symbols, source, position);
        symbols
            .symbols
            .iter()
            .filter(|sym| match &sym.proc_scope {
                None => true,
                Some(scope) => cursor_scope
                    .as_deref()
                    .is_some_and(|cs| cs.eq_ignore_ascii_case(scope.as_str())),
            })
            .map(|sym| {
                let kind = symbol_kind_to_completion_kind(&sym.kind);
                CompletionItem {
                    label: sym.name.to_string(),
                    kind: Some(kind),
                    detail: sym.type_name.as_ref().map(|t| t.to_string()),
                    ..Default::default()
                }
            })
            .collect::<Vec<_>>()
    }) {
        items.extend(sym_items);
    }

    // Public symbols from other files in the workspace (cross-module completion)
    for sym in host.all_public_symbols_from_other_files(uri) {
        let kind = symbol_kind_to_completion_kind(&sym.kind);
        items.push(CompletionItem {
            label: sym.name.to_string(),
            kind: Some(kind),
            detail: sym.type_name.as_ref().map(|t| t.to_string()),
            ..Default::default()
        });
    }

    // Workbook sheet names from workbook-context.json
    for sheet in host.workbook_sheets() {
        items.push(CompletionItem {
            label: sheet,
            kind: Some(CompletionItemKind::MODULE),
            detail: Some("Worksheet".to_string()),
            ..Default::default()
        });
    }

    for table in host.workbook_tables() {
        items.push(CompletionItem {
            label: table,
            kind: Some(CompletionItemKind::STRUCT),
            detail: Some("Table".to_string()),
            ..Default::default()
        });
    }

    for named_range in host.workbook_named_ranges() {
        items.push(CompletionItem {
            label: named_range,
            kind: Some(CompletionItemKind::CONSTANT),
            detail: Some("Named Range".to_string()),
            ..Default::default()
        });
    }

    items
}

fn proc_at_position(symbols: &SymbolTable, source: &str, position: Position) -> Option<smol_str::SmolStr> {
    let offset = position_to_offset(source, position)?;
    symbols
        .proc_ranges
        .iter()
        .find(|(_, range)| (range.start as usize) <= offset && offset < (range.end as usize))
        .map(|(name, _)| name.clone())
}
