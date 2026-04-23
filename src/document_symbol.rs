use tower_lsp::lsp_types::{DocumentSymbol, Url};

use crate::analysis::resolve::text_range_to_lsp_range;
use crate::analysis::symbols;
use crate::analysis::AnalysisHost;

pub fn document_symbols(host: &AnalysisHost, uri: &Url) -> Vec<DocumentSymbol> {
    host.with_source(uri, |symbol_table, source| {
        build_hierarchy(symbol_table, source)
    })
    .unwrap_or_default()
}

fn build_hierarchy(table: &symbols::SymbolTable, source: &str) -> Vec<DocumentSymbol> {
    let mut top_level: Vec<DocumentSymbol> = Vec::new();

    for sym in &table.symbols {
        if sym.proc_scope.is_some() {
            continue;
        }

        let selection_range = text_range_to_lsp_range(source, sym.span);

        let full_range = if matches!(
            sym.kind,
            symbols::SymbolKind::Procedure
                | symbols::SymbolKind::Function
                | symbols::SymbolKind::Property
        ) {
            table
                .proc_ranges
                .iter()
                .find(|(name, _)| name.eq_ignore_ascii_case(&sym.name))
                .map(|(_, r)| text_range_to_lsp_range(source, *r))
                .unwrap_or(selection_range)
        } else {
            selection_range
        };

        let children = collect_children(table, source, &sym.name);

        #[allow(deprecated)]
        top_level.push(DocumentSymbol {
            name: sym.name.to_string(),
            detail: sym.type_name.as_ref().map(|t| t.to_string()),
            kind: sym.kind.to_lsp_symbol_kind(),
            deprecated: None,
            range: full_range,
            selection_range,
            children: if children.is_empty() {
                None
            } else {
                Some(children)
            },
            tags: None,
        });
    }

    top_level
}

fn collect_children(
    table: &symbols::SymbolTable,
    source: &str,
    proc_name: &str,
) -> Vec<DocumentSymbol> {
    table
        .symbols
        .iter()
        .filter(|s| {
            s.proc_scope
                .as_ref()
                .is_some_and(|p| p.eq_ignore_ascii_case(proc_name))
        })
        .map(|s| {
            let range = text_range_to_lsp_range(source, s.span);
            #[allow(deprecated)]
            DocumentSymbol {
                name: s.name.to_string(),
                detail: s.type_name.as_ref().map(|t| t.to_string()),
                kind: s.kind.to_lsp_symbol_kind(),
                deprecated: None,
                range,
                selection_range: range,
                children: None,
                tags: None,
            }
        })
        .collect()
}
