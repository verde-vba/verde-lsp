pub mod diagnostics;
pub mod resolve;
pub mod symbols;

use dashmap::DashMap;
use tower_lsp::lsp_types::*;

use crate::parser::lexer::SpannedToken;
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
    /// Lexed tokens retained from parse, for semantic tokens and formatting.
    pub tokens: Vec<SpannedToken>,
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
            .unwrap_or_else(|e| e.into_inner()) = ctx;
    }

    /// Read `path`, parse as `WorkbookContext` JSON, and update the context.
    /// Returns `true` on success, `false` if the file is missing or invalid.
    pub fn reload_workbook_context_from_path(&self, path: &std::path::Path) -> bool {
        log::info!("[analysis:reload_workbook_context] path={:?}", path);
        match std::fs::read_to_string(path) {
            Ok(content) => {
                log::debug!("[analysis:reload_workbook_context] read {} bytes", content.len());
                match serde_json::from_str::<WorkbookContext>(&content) {
                    Ok(ctx) => {
                        log::info!("[analysis:reload_workbook_context] loaded — sheets={} tables={} named_ranges={}",
                            ctx.sheets.len(), ctx.tables.len(), ctx.named_ranges.len());
                        self.set_workbook_context(ctx);
                        true
                    }
                    Err(e) => {
                        log::warn!("[analysis:reload_workbook_context] JSON parse error: {}", e);
                        false
                    }
                }
            }
            Err(e) => {
                log::debug!("[analysis:reload_workbook_context] file read failed: {}", e);
                false
            }
        }
    }

    fn read_workbook_context(&self) -> std::sync::RwLockReadGuard<'_, WorkbookContext> {
        self.workbook_context
            .read()
            .unwrap_or_else(|e| e.into_inner())
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

    /// Returns a snapshot of sheets, tables, and named ranges in a single lock acquisition.
    pub fn workbook_context_snapshot(&self) -> (Vec<String>, Vec<String>, Vec<String>) {
        let ctx = self.read_workbook_context();
        (
            ctx.sheets.clone(),
            ctx.tables.clone(),
            ctx.named_ranges.clone(),
        )
    }

    pub fn update(&self, uri: Url, source: String, parse_result: ParseResult) {
        let symbols = symbols::build_symbol_table(&parse_result.ast);
        log::debug!("[analysis:update] uri={} source_len={} errors={} symbols={} tokens={}",
            uri, source.len(), parse_result.errors.len(), symbols.symbols.len(), parse_result.tokens.len());
        let tokens = parse_result.tokens.clone();
        self.files.insert(
            uri,
            FileAnalysis {
                parse_result,
                symbols,
                source,
                tokens,
            },
        );
        log::debug!("[analysis:update] total files tracked: {}", self.files.len());
    }

    pub fn remove(&self, uri: &Url) {
        log::debug!("[analysis:remove] uri={}", uri);
        self.files.remove(uri);
    }

    pub fn diagnostics(&self, uri: &Url) -> Vec<Diagnostic> {
        let total_files = self.files.len();
        log::debug!("[analysis:diagnostics] uri={} file_exists={} total_files={}", uri, self.files.contains_key(uri), total_files);
        if let Some(file) = self.files.get(uri) {
            // Single iteration over all files to collect both public symbol names
            // and module names from other files.
            let mut cross_module_names: std::collections::HashSet<smol_str::SmolStr> =
                std::collections::HashSet::new();
            for entry in self.files.iter() {
                if entry.key() == uri {
                    continue;
                }
                // Collect public symbol names
                for sym in &entry.symbols.symbols {
                    if sym.visibility == crate::parser::ast::Visibility::Public
                        && sym.proc_scope.is_none()
                    {
                        cross_module_names
                            .insert(smol_str::SmolStr::new(sym.name.to_ascii_lowercase()));
                    }
                }
                // Collect module name (filename stem)
                if let Some(name) = entry
                    .key()
                    .path_segments()
                    .and_then(|mut s| s.next_back())
                    .and_then(|f| f.split('.').next())
                {
                    cross_module_names.insert(smol_str::SmolStr::new(name.to_ascii_lowercase()));
                }
            }
            log::debug!("[analysis:diagnostics] cross_module_names({})={:?}", cross_module_names.len(), cross_module_names);
            let diags = diagnostics::compute(
                &file.parse_result,
                &file.symbols,
                &file.source,
                &cross_module_names,
            );
            if !diags.is_empty() {
                for d in &diags {
                    log::debug!("[analysis:diagnostics] diag: {} @ {:?}", d.message, d.range);
                }
            }
            diags
        } else {
            Vec::new()
        }
    }

    pub fn with_source<T>(&self, uri: &Url, f: impl FnOnce(&SymbolTable, &str) -> T) -> Option<T> {
        self.files
            .get(uri)
            .map(|file| f(&file.symbols, &file.source))
    }

    pub fn with_tokens<T>(
        &self,
        uri: &Url,
        f: impl FnOnce(&SymbolTable, &str, &[SpannedToken]) -> T,
    ) -> Option<T> {
        self.files
            .get(uri)
            .map(|file| f(&file.symbols, &file.source, &file.tokens))
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

    /// Call hierarchy: prepare — returns the item at cursor if it is a procedure.
    pub fn prepare_call_hierarchy(
        &self,
        uri: &Url,
        position: Position,
    ) -> Option<Vec<tower_lsp::lsp_types::CallHierarchyItem>> {
        crate::call_hierarchy::prepare_call_hierarchy(self, uri, position)
    }

    /// Call hierarchy: incoming — callers of the given item across all files.
    pub fn incoming_calls(
        &self,
        item: &tower_lsp::lsp_types::CallHierarchyItem,
    ) -> Vec<tower_lsp::lsp_types::CallHierarchyIncomingCall> {
        crate::call_hierarchy::incoming_calls(self, item)
    }

    /// Call hierarchy: outgoing — callees invoked inside the given item's body.
    pub fn outgoing_calls(
        &self,
        item: &tower_lsp::lsp_types::CallHierarchyItem,
    ) -> Vec<tower_lsp::lsp_types::CallHierarchyOutgoingCall> {
        crate::call_hierarchy::outgoing_calls(self, item)
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

    /// Return Public, module-level symbols from the file whose module name
    /// (filename stem) matches `module_name` (case-insensitive). Excludes
    /// `current_uri` itself. Used for `Module1.Foo` dot-access completion.
    pub fn public_symbols_from_module(
        &self,
        current_uri: &Url,
        module_name: &str,
    ) -> Vec<symbols::Symbol> {
        for entry in self.files.iter() {
            if entry.key() == current_uri {
                continue;
            }
            let stem = entry
                .key()
                .path_segments()
                .and_then(|mut s| s.next_back())
                .and_then(|f| f.split('.').next())
                .unwrap_or("");
            if stem.eq_ignore_ascii_case(module_name) {
                return entry
                    .symbols
                    .symbols
                    .iter()
                    .filter(|s| {
                        s.visibility == crate::parser::ast::Visibility::Public
                            && s.proc_scope.is_none()
                    })
                    .cloned()
                    .collect();
            }
        }
        Vec::new()
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

    /// Find the file whose module name (filename stem) matches `name`
    /// (case-insensitive), excluding `current_uri`. Returns the URI and
    /// the public symbol list. Used for module-name hover / definition.
    pub fn find_module_by_name(
        &self,
        current_uri: &Url,
        name: &str,
    ) -> Option<(Url, Vec<symbols::Symbol>)> {
        for entry in self.files.iter() {
            if entry.key() == current_uri {
                continue;
            }
            let stem = entry
                .key()
                .path_segments()
                .and_then(|mut s| s.next_back())
                .and_then(|f| f.split('.').next())
                .unwrap_or("");
            if stem.eq_ignore_ascii_case(name) {
                let public_syms: Vec<symbols::Symbol> = entry
                    .symbols
                    .symbols
                    .iter()
                    .filter(|s| {
                        s.visibility == crate::parser::ast::Visibility::Public
                            && s.proc_scope.is_none()
                    })
                    .cloned()
                    .collect();
                return Some((entry.key().clone(), public_syms));
            }
        }
        None
    }
}
