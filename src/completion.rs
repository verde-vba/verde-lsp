use tower_lsp::lsp_types::*;

use crate::analysis::resolve::{parse_dot_access_at, position_to_offset};
use crate::analysis::symbols::{SymbolDetail, SymbolKind, SymbolTable};
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
        SymbolKind::UdtMember => CompletionItemKind::FIELD,
    }
}

pub fn complete(host: &AnalysisHost, uri: &Url, position: Position) -> Vec<CompletionItem> {
    // Dot-access short-circuit: return only UDT members when cursor follows `identifier.`
    if let Some(items) = host
        .with_source(uri, |symbols, source| {
            complete_dot_access(symbols, source, position)
        })
        .flatten()
    {
        return items;
    }

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

    // Workbook names from workbook-context.json
    push_named_items(
        &mut items,
        host.workbook_sheets(),
        CompletionItemKind::MODULE,
        "Worksheet",
    );
    push_named_items(
        &mut items,
        host.workbook_tables(),
        CompletionItemKind::STRUCT,
        "Table",
    );
    push_named_items(
        &mut items,
        host.workbook_named_ranges(),
        CompletionItemKind::CONSTANT,
        "Named Range",
    );

    items
}

/// Returns `Some(items)` when cursor is in a dot-access context (`var.`),
/// in which case only contextually relevant members are offered.
/// Returns `None` when not in dot-access context (caller uses normal completion).
fn complete_dot_access(
    symbols: &SymbolTable,
    source: &str,
    position: Position,
) -> Option<Vec<CompletionItem>> {
    let offset = position_to_offset(source, position)?;
    let (var_name, _) = parse_dot_access_at(source, offset)?;

    // `Me.` offers the class module's own procedures and module-level variables.
    if var_name.eq_ignore_ascii_case("Me") {
        let items: Vec<CompletionItem> = symbols
            .symbols
            .iter()
            .filter(|s| {
                s.proc_scope.is_none()
                    && matches!(
                        s.kind,
                        SymbolKind::Procedure
                            | SymbolKind::Function
                            | SymbolKind::Property
                            | SymbolKind::Variable
                            | SymbolKind::Constant
                    )
            })
            .map(|s| CompletionItem {
                label: s.name.to_string(),
                kind: Some(symbol_kind_to_completion_kind(&s.kind)),
                detail: s.type_name.as_ref().map(|t| t.to_string()),
                ..Default::default()
            })
            .collect();
        return Some(items);
    }

    // Prefer proc-scoped variable over module-level when cursor is inside a proc.
    let cursor_proc = symbols
        .proc_ranges
        .iter()
        .find(|(_, r)| offset >= r.start as usize && offset <= r.end as usize)
        .map(|(name, _)| name.clone());

    let type_name = cursor_proc
        .as_ref()
        .and_then(|proc_name| {
            symbols.symbols.iter().find(|s| {
                s.name.eq_ignore_ascii_case(&var_name)
                    && matches!(s.kind, SymbolKind::Variable | SymbolKind::Parameter)
                    && s.proc_scope
                        .as_ref()
                        .map(|p| p.eq_ignore_ascii_case(proc_name))
                        .unwrap_or(false)
            })
        })
        .or_else(|| {
            symbols.symbols.iter().find(|s| {
                s.name.eq_ignore_ascii_case(&var_name)
                    && matches!(s.kind, SymbolKind::Variable | SymbolKind::Parameter)
                    && s.proc_scope.is_none()
            })
        })
        .and_then(|s| s.type_name.clone());

    let type_name = type_name?;

    let members: Vec<CompletionItem> = symbols
        .symbols
        .iter()
        .filter(|s| {
            matches!(s.kind, SymbolKind::UdtMember)
                && match &s.detail {
                    SymbolDetail::UdtMember { parent_type, .. } => {
                        parent_type.eq_ignore_ascii_case(&type_name)
                    }
                    _ => false,
                }
        })
        .map(|s| CompletionItem {
            label: s.name.to_string(),
            kind: Some(CompletionItemKind::FIELD),
            detail: s.type_name.as_ref().map(|t| t.to_string()),
            ..Default::default()
        })
        .collect();

    Some(members)
}

fn push_named_items(
    items: &mut Vec<CompletionItem>,
    names: Vec<String>,
    kind: CompletionItemKind,
    detail: &str,
) {
    for name in names {
        items.push(CompletionItem {
            label: name,
            kind: Some(kind),
            detail: Some(detail.to_string()),
            ..Default::default()
        });
    }
}

fn proc_at_position(
    symbols: &SymbolTable,
    source: &str,
    position: Position,
) -> Option<smol_str::SmolStr> {
    let offset = position_to_offset(source, position)?;
    symbols
        .proc_ranges
        .iter()
        .find(|(_, range)| (range.start as usize) <= offset && offset < (range.end as usize))
        .map(|(name, _)| name.clone())
}
