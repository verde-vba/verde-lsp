use tower_lsp::lsp_types::*;

use crate::analysis::resolve::text_range_to_lsp_range;
use crate::analysis::AnalysisHost;

pub fn workspace_symbols(host: &AnalysisHost, query: &str) -> Vec<SymbolInformation> {
    let query_lower = query.to_ascii_lowercase();
    let mut results = Vec::new();

    for (uri, source) in host.all_file_sources() {
        if let Some(syms) = host.with_source(&uri, |table, _| {
            let mut items = Vec::new();
            for sym in &table.symbols {
                if sym.proc_scope.is_some() {
                    continue;
                }
                let name_lower = sym.name.to_ascii_lowercase();
                if !query_lower.is_empty() && !name_lower.contains(&query_lower) {
                    continue;
                }
                let lsp_kind = sym.kind.to_lsp_symbol_kind();
                let range = text_range_to_lsp_range(&source, sym.span);
                #[allow(deprecated)]
                items.push(SymbolInformation {
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
            items
        }) {
            results.extend(syms);
        }
    }
    results
}
