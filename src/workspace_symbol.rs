use tower_lsp::lsp_types::*;

use crate::analysis::resolve::text_range_to_lsp_range;
use crate::analysis::symbols::SymbolKind as VbaSymbolKind;
use crate::analysis::AnalysisHost;

pub fn workspace_symbols(host: &AnalysisHost, query: &str) -> Vec<SymbolInformation> {
    let query_lower = query.to_ascii_lowercase();
    let mut results = Vec::new();

    for (uri, source) in host.all_file_sources() {
        let Some(table) = host.symbol_table(&uri) else {
            continue;
        };
        for sym in &table.symbols {
            if sym.proc_scope.is_some() {
                continue; // skip locals and parameters
            }
            let name_lower = sym.name.to_ascii_lowercase();
            if !query_lower.is_empty() && !name_lower.contains(&query_lower) {
                continue;
            }
            let lsp_kind = vba_kind_to_lsp(&sym.kind);
            let range = text_range_to_lsp_range(&source, sym.span);
            #[allow(deprecated)]
            results.push(SymbolInformation {
                name: sym.name.to_string(),
                kind: lsp_kind,
                location: Location {
                    uri: uri.clone(),
                    range,
                },
                tags: None,
                deprecated: None,
                container_name: None,
            });
        }
    }
    results
}

fn vba_kind_to_lsp(kind: &VbaSymbolKind) -> SymbolKind {
    match kind {
        VbaSymbolKind::Procedure | VbaSymbolKind::Property => SymbolKind::FUNCTION,
        VbaSymbolKind::Function => SymbolKind::FUNCTION,
        VbaSymbolKind::Variable | VbaSymbolKind::Constant => SymbolKind::VARIABLE,
        VbaSymbolKind::Parameter => SymbolKind::VARIABLE,
        VbaSymbolKind::TypeDef => SymbolKind::STRUCT,
        VbaSymbolKind::EnumDef => SymbolKind::ENUM,
        VbaSymbolKind::EnumMember => SymbolKind::ENUM_MEMBER,
    }
}
