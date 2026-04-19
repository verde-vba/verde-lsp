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
}

impl AnalysisHost {
    pub fn new() -> Self {
        Self {
            files: DashMap::new(),
        }
    }

    pub fn update(&self, uri: Url, parse_result: ParseResult) {
        let symbols = symbols::build_symbol_table(&parse_result.ast);
        self.files.insert(uri, FileAnalysis { parse_result, symbols });
    }

    pub fn remove(&self, uri: &Url) {
        self.files.remove(uri);
    }

    pub fn diagnostics(&self, uri: &Url) -> Vec<Diagnostic> {
        if let Some(file) = self.files.get(uri) {
            diagnostics::compute(&file.parse_result, &file.symbols)
        } else {
            Vec::new()
        }
    }

    pub fn with_symbols<T>(&self, uri: &Url, f: impl FnOnce(&SymbolTable) -> T) -> Option<T> {
        self.files.get(uri).map(|file| f(&file.symbols))
    }

    pub fn symbol_table(&self, uri: &Url) -> Option<SymbolTable> {
        self.files.get(uri).map(|f| f.symbols.clone())
    }
}
