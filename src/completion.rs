use tower_lsp::lsp_types::*;

use crate::analysis::AnalysisHost;
use crate::analysis::symbols::SymbolKind;
use crate::vba_builtins;

pub fn complete(host: &AnalysisHost, uri: &Url, _position: Position) -> Vec<CompletionItem> {
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

    // Symbols from current file
    if let Some(symbols) = host.symbol_table(uri) {
        for sym in &symbols.symbols {
            let kind = match sym.kind {
                SymbolKind::Procedure => CompletionItemKind::METHOD,
                SymbolKind::Function => CompletionItemKind::FUNCTION,
                SymbolKind::Property => CompletionItemKind::PROPERTY,
                SymbolKind::Variable => CompletionItemKind::VARIABLE,
                SymbolKind::Constant => CompletionItemKind::CONSTANT,
                SymbolKind::Parameter => CompletionItemKind::VARIABLE,
                SymbolKind::TypeDef => CompletionItemKind::STRUCT,
                SymbolKind::EnumDef => CompletionItemKind::ENUM,
                SymbolKind::EnumMember => CompletionItemKind::ENUM_MEMBER,
            };

            items.push(CompletionItem {
                label: sym.name.to_string(),
                kind: Some(kind),
                detail: sym.type_name.as_ref().map(|t| t.to_string()),
                ..Default::default()
            });
        }
    }

    items
}
