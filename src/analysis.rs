pub mod symbols;
pub mod resolve;
pub mod diagnostics;

use dashmap::DashMap;
use tower_lsp::lsp_types::*;

use crate::parser::ParseResult;
use symbols::SymbolTable;

pub struct AnalysisHost {
    files: DashMap<Url, FileAnalysis>,
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
        }
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
            diagnostics::compute(&file.parse_result, &file.symbols, &file.source)
        } else {
            Vec::new()
        }
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
