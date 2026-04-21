pub mod diagnostics;
pub mod resolve;
pub mod symbols;

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
    #[serde(default)]
    pub tables: Vec<String>,
    #[serde(default)]
    pub named_ranges: Vec<String>,
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
        *self
            .workbook_context
            .write()
            .expect("workbook_context RwLock poisoned: prior panic in write context") = ctx;
    }

    /// Read `path`, parse as `WorkbookContext` JSON, and update the context.
    /// Returns `true` on success, `false` if the file is missing or invalid.
    pub fn reload_workbook_context_from_path(&self, path: &std::path::Path) -> bool {
        if let Ok(content) = std::fs::read_to_string(path) {
            if let Ok(ctx) = serde_json::from_str::<WorkbookContext>(&content) {
                self.set_workbook_context(ctx);
                return true;
            }
        }
        false
    }

    fn read_workbook_context(&self) -> std::sync::RwLockReadGuard<'_, WorkbookContext> {
        self.workbook_context
            .read()
            .expect("workbook_context RwLock poisoned")
    }

    pub fn workbook_sheets(&self) -> Vec<String> {
        self.read_workbook_context().sheets.clone()
    }

    pub fn workbook_tables(&self) -> Vec<String> {
        self.read_workbook_context().tables.clone()
    }

    pub fn workbook_named_ranges(&self) -> Vec<String> {
        self.read_workbook_context().named_ranges.clone()
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

    pub fn with_source<T>(&self, uri: &Url, f: impl FnOnce(&SymbolTable, &str) -> T) -> Option<T> {
        self.files
            .get(uri)
            .map(|file| f(&file.symbols, &file.source))
    }

    pub fn symbol_table(&self, uri: &Url) -> Option<SymbolTable> {
        self.files.get(uri).map(|f| f.symbols.clone())
    }

    /// Returns (uri, source) pairs for every registered file. Used by
    /// cross-file features (e.g. references) to search all workspace files.
    pub fn all_file_sources(&self) -> Vec<(Url, String)> {
        self.files
            .iter()
            .map(|e| (e.key().clone(), e.source.clone()))
            .collect()
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

    /// Return inlay hints (variable/constant type annotations) for the given file.
    /// `range` is accepted for API compatibility but currently ignored — all symbols
    /// in the file are returned regardless of position.
    pub fn inlay_hints(&self, uri: &Url, _range: Option<Range>) -> Vec<InlayHint> {
        if let Some(file) = self.files.get(uri) {
            crate::inlay_hint::inlay_hints(&file.source, &file.symbols)
        } else {
            Vec::new()
        }
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
