pub mod symbols;
pub mod resolve;
pub mod diagnostics;

use dashmap::DashMap;
use tower_lsp::lsp_types::*;

use crate::parser::ParseResult;
use symbols::SymbolTable;

/// Workspace-level context loaded from `workbook-context.json`.
/// All fields are optional; absent keys default to empty.
#[derive(Debug, Default, Clone, serde::Deserialize)]
pub struct WorkbookContext {
    #[serde(default)]
    pub sheets: Vec<String>,
}

pub struct AnalysisHost {
    files: DashMap<Url, FileAnalysis>,
    workbook_context: std::sync::RwLock<WorkbookContext>,
}

pub struct FileAnalysis {
    pub parse_result: ParseResult,
    pub symbols: SymbolTable,
    pub source: String,
}

impl Default for AnalysisHost {
    fn default() -> Self {
        Self::new()
    }
}

impl AnalysisHost {
    pub fn new() -> Self {
        Self {
            files: DashMap::new(),
            workbook_context: std::sync::RwLock::new(WorkbookContext::default()),
        }
    }

    pub fn set_workbook_context(&self, ctx: WorkbookContext) {
        *self.workbook_context.write().unwrap() = ctx;
    }

    pub fn workbook_sheets(&self) -> Vec<String> {
        self.workbook_context.read().unwrap().sheets.clone()
    }

    pub fn update(&self, uri: Url, source: String, parse_result: ParseResult) {
        let symbols = symbols::build_symbol_table(&parse_result.ast);
        self.files.insert(
            uri,
            FileAnalysis {
                parse_result,
                symbols,
                source,
            },
        );
    }

    pub fn remove(&self, uri: &Url) {
        self.files.remove(uri);
    }

    pub fn diagnostics(&self, uri: &Url) -> Vec<Diagnostic> {
        if let Some(file) = self.files.get(uri) {
            let mut cross_module_names: std::collections::HashSet<String> = self
                .all_public_symbols_from_other_files(uri)
                .into_iter()
                .map(|s| s.name.to_ascii_lowercase())
                .collect();
            cross_module_names.extend(self.collect_other_module_names(uri));
            diagnostics::compute(
                &file.parse_result,
                &file.symbols,
                &file.source,
                &cross_module_names,
            )
        } else {
            Vec::new()
        }
    }

    /// Extract bare module names (filename without extension, lowercased) from
    /// all registered files except `current_uri`. Used to allow `ModuleA.Foo`
    /// qualified calls without triggering undeclared-variable warnings.
    fn collect_other_module_names(&self, current_uri: &Url) -> Vec<String> {
        self.files
            .iter()
            .filter(|e| e.key() != current_uri)
            .filter_map(|e| {
                e.key()
                    .path_segments()
                    .and_then(|mut s| s.next_back())
                    .and_then(|f| f.split('.').next())
                    .map(|name| name.to_ascii_lowercase())
            })
            .collect()
    }

    pub fn with_source<T>(
        &self,
        uri: &Url,
        f: impl FnOnce(&SymbolTable, &str) -> T,
    ) -> Option<T> {
        self.files.get(uri).map(|file| f(&file.symbols, &file.source))
    }

    pub fn symbol_table(&self, uri: &Url) -> Option<SymbolTable> {
        self.files.get(uri).map(|f| f.symbols.clone())
    }

    /// Returns all Public, module-level symbols from every file EXCEPT `current_uri`.
    /// Used for cross-module completion.
    pub fn all_public_symbols_from_other_files(&self, current_uri: &Url) -> Vec<symbols::Symbol> {
        let mut result = Vec::new();
        for entry in self.files.iter() {
            if entry.key() == current_uri {
                continue;
            }
            for sym in &entry.symbols.symbols {
                if sym.visibility == crate::parser::ast::Visibility::Public
                    && sym.proc_scope.is_none()
                {
                    result.push(sym.clone());
                }
            }
        }
        result
    }

    /// Find the first Public module-level symbol matching `name` (case-insensitive)
    /// across all files except `current_uri`. Returns the source URI and symbol.
    pub fn find_public_symbol_in_other_files(
        &self,
        current_uri: &Url,
        name: &str,
    ) -> Option<(Url, symbols::Symbol)> {
        for entry in self.files.iter() {
            if entry.key() == current_uri {
                continue;
            }
            if let Some(sym) = entry.symbols.symbols.iter().find(|s| {
                s.visibility == crate::parser::ast::Visibility::Public
                    && s.proc_scope.is_none()
                    && s.name.eq_ignore_ascii_case(name)
            }) {
                return Some((entry.key().clone(), sym.clone()));
            }
        }
        None
    }
}
