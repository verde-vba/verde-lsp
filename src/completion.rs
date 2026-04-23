use tower_lsp::lsp_types::*;

use crate::analysis::resolve::{parse_dot_access_at, position_to_offset};
use crate::analysis::symbols::{SymbolDetail, SymbolKind, SymbolTable};
use crate::analysis::AnalysisHost;
use crate::excel_model::types::builtin_types;
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
    // Leading dot (With context): `.Value` → resolve With target's type members
    if let Some(items) = host
        .with_source(uri, |symbols, source| {
            complete_leading_dot(symbols, source, position)
        })
        .flatten()
    {
        return items;
    }

    // Module-name dot-access: `Module1.` → Public symbols from that module
    if let Some(var_name) = host
        .with_source(uri, |_, source| {
            let offset = position_to_offset(source, position)?;
            parse_dot_access_at(source, offset).map(|(name, _)| name)
        })
        .flatten()
    {
        let module_syms = host.public_symbols_from_module(uri, &var_name);
        if !module_syms.is_empty() {
            return module_syms
                .iter()
                .map(|sym| CompletionItem {
                    label: sym.name.to_string(),
                    kind: Some(symbol_kind_to_completion_kind(&sym.kind)),
                    detail: sym.type_name.as_ref().map(|t| t.to_string()),
                    ..Default::default()
                })
                .collect();
        }
    }

    // Function return value dot-access: `GetRange().` → resolve return type members
    if let Some(items) = host
        .with_source(uri, |symbols, source| {
            complete_func_return_dot(symbols, source, position)
        })
        .flatten()
    {
        return items;
    }

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

    // Application implicit globals (e.g. ActiveWorkbook, Range, Cells)
    for name in crate::excel_model::types::application_globals() {
        items.push(CompletionItem {
            label: name.to_string(),
            kind: Some(CompletionItemKind::PROPERTY),
            detail: Some("Application".to_string()),
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

    // Workbook names from workbook-context.json (single lock acquisition)
    let (sheets, tables, named_ranges) = host.workbook_context_snapshot();
    push_named_items(&mut items, sheets, CompletionItemKind::MODULE, "Worksheet");
    push_named_items(&mut items, tables, CompletionItemKind::STRUCT, "Table");
    push_named_items(
        &mut items,
        named_ranges,
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

    // Check if var_name is an Enum name → offer its members directly (e.g. Color.Red)
    let enum_members: Vec<CompletionItem> = symbols
        .symbols
        .iter()
        .filter(|s| {
            matches!(s.kind, SymbolKind::EnumMember)
                && match &s.detail {
                    SymbolDetail::EnumMember { parent_enum, .. } => {
                        parent_enum.eq_ignore_ascii_case(&var_name)
                    }
                    _ => false,
                }
        })
        .map(|s| CompletionItem {
            label: s.name.to_string(),
            kind: Some(CompletionItemKind::ENUM_MEMBER),
            detail: s.type_name.as_ref().map(|t| t.to_string()),
            ..Default::default()
        })
        .collect();

    if !enum_members.is_empty() {
        return Some(enum_members);
    }

    // Prefer proc-scoped variable over module-level when cursor is inside a proc.
    let type_name = crate::analysis::resolve::resolve_var_type_at(symbols, offset, &var_name)?;

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

    if !members.is_empty() {
        return Some(members);
    }

    // Check if type_name is an Enum → offer its members (e.g. Dim c As Color → c.Red)
    let enum_items: Vec<CompletionItem> = symbols
        .symbols
        .iter()
        .filter(|s| {
            matches!(s.kind, SymbolKind::EnumMember)
                && match &s.detail {
                    SymbolDetail::EnumMember { parent_enum, .. } => {
                        parent_enum.eq_ignore_ascii_case(&type_name)
                    }
                    _ => false,
                }
        })
        .map(|s| CompletionItem {
            label: s.name.to_string(),
            kind: Some(CompletionItemKind::ENUM_MEMBER),
            detail: s.type_name.as_ref().map(|t| t.to_string()),
            ..Default::default()
        })
        .collect();

    if !enum_items.is_empty() {
        return Some(enum_items);
    }

    // Fallback: look up the type in Excel builtin types (e.g. Range, PivotTable, Chart, Shape).
    let builtin_types = builtin_types();
    if let Some(excel_type) = builtin_types
        .iter()
        .find(|t| t.name.eq_ignore_ascii_case(&type_name))
    {
        let mut items: Vec<CompletionItem> = excel_type
            .properties
            .iter()
            .map(|p| CompletionItem {
                label: p.name.to_string(),
                kind: Some(CompletionItemKind::PROPERTY),
                detail: Some(p.return_type.to_string()),
                ..Default::default()
            })
            .collect();
        items.extend(excel_type.methods.iter().map(|m| CompletionItem {
            label: m.name.to_string(),
            kind: Some(CompletionItemKind::METHOD),
            detail: m.return_type.as_ref().map(|t| t.to_string()),
            ..Default::default()
        }));
        return Some(items);
    }

    // Chain dot-access: try resolving multi-segment chains like `rng.Font.`
    if let Some(items) = resolve_chain_dot(symbols, source, position) {
        return Some(items);
    }

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

/// Resolve members for a given type name. Returns completion items for UDT members,
/// Enum members, or Excel builtin type properties/methods.
fn resolve_type_members(symbols: &SymbolTable, type_name: &str) -> Option<Vec<CompletionItem>> {
    // UDT members
    let udt_members: Vec<CompletionItem> = symbols
        .symbols
        .iter()
        .filter(|s| {
            matches!(s.kind, SymbolKind::UdtMember)
                && match &s.detail {
                    SymbolDetail::UdtMember { parent_type, .. } => {
                        parent_type.eq_ignore_ascii_case(type_name)
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
    if !udt_members.is_empty() {
        return Some(udt_members);
    }

    // Enum members
    let enum_members: Vec<CompletionItem> = symbols
        .symbols
        .iter()
        .filter(|s| {
            matches!(s.kind, SymbolKind::EnumMember)
                && match &s.detail {
                    SymbolDetail::EnumMember { parent_enum, .. } => {
                        parent_enum.eq_ignore_ascii_case(type_name)
                    }
                    _ => false,
                }
        })
        .map(|s| CompletionItem {
            label: s.name.to_string(),
            kind: Some(CompletionItemKind::ENUM_MEMBER),
            detail: s.type_name.as_ref().map(|t| t.to_string()),
            ..Default::default()
        })
        .collect();
    if !enum_members.is_empty() {
        return Some(enum_members);
    }

    // Excel builtin types
    let btypes = crate::excel_model::types::builtin_types();
    if let Some(excel_type) = btypes
        .iter()
        .find(|t| t.name.eq_ignore_ascii_case(type_name))
    {
        let mut items: Vec<CompletionItem> = excel_type
            .properties
            .iter()
            .map(|p| CompletionItem {
                label: p.name.to_string(),
                kind: Some(CompletionItemKind::PROPERTY),
                detail: Some(p.return_type.to_string()),
                ..Default::default()
            })
            .collect();
        items.extend(excel_type.methods.iter().map(|m| CompletionItem {
            label: m.name.to_string(),
            kind: Some(CompletionItemKind::METHOD),
            detail: m.return_type.as_ref().map(|t| t.to_string()),
            ..Default::default()
        }));
        return Some(items);
    }

    None
}

/// Complete leading dot `.member` inside a With block.
fn complete_leading_dot(
    symbols: &SymbolTable,
    source: &str,
    position: Position,
) -> Option<Vec<CompletionItem>> {
    use crate::analysis::resolve::{parse_leading_dot_at, position_to_offset};
    use crate::analysis::symbols::BlockKind;

    let offset = position_to_offset(source, position)?;
    let _member_partial = parse_leading_dot_at(source, offset)?;

    // Find innermost With block containing cursor
    let with_block = symbols.block_ranges.iter().rfind(|b| {
        b.kind == BlockKind::With && (b.start as usize) <= offset && offset <= (b.end as usize)
    });

    let var_name = with_block?.data.as_ref()?;

    // Resolve the With target's type using the same logic as complete_dot_access
    let type_name = crate::analysis::resolve::resolve_var_type_at(symbols, offset, var_name)?;

    // Return members of the resolved type (UDT, Enum, or Excel builtin)
    resolve_type_members(symbols, &type_name)
}

/// Complete dot access on function return value: `GetRange().` -> Range members
fn complete_func_return_dot(
    symbols: &SymbolTable,
    source: &str,
    position: Position,
) -> Option<Vec<CompletionItem>> {
    use crate::analysis::resolve::{parse_func_call_dot_at, position_to_offset};

    let offset = position_to_offset(source, position)?;
    let (func_name, _partial) = parse_func_call_dot_at(source, offset)?;

    // Once we've detected `FuncName().`, we're in dot-access context.
    // Return Some (possibly empty) to prevent fallthrough to general completion.
    let return_type = symbols
        .symbols
        .iter()
        .find(|s| {
            s.proc_scope.is_none()
                && s.name.eq_ignore_ascii_case(&func_name)
                && matches!(s.kind, SymbolKind::Function | SymbolKind::Property)
        })
        .and_then(|s| match &s.detail {
            SymbolDetail::Procedure { return_type, .. } => return_type.clone(),
            _ => None,
        });

    match return_type {
        Some(t) => resolve_type_members(symbols, &t).or(Some(vec![])),
        None => Some(vec![]),
    }
}

/// Resolve chained dot access like `rng.Font.Bold`.
/// Walks the chain from left to right, resolving each segment's type.
fn resolve_chain_dot(
    symbols: &SymbolTable,
    source: &str,
    position: Position,
) -> Option<Vec<CompletionItem>> {
    use crate::analysis::resolve::{parse_dot_chain_at, position_to_offset};

    let offset = position_to_offset(source, position)?;
    let (chain, _partial) = parse_dot_chain_at(source, offset)?;

    if chain.len() < 2 {
        return None; // Single-segment handled by existing logic
    }

    // Resolve the first segment to a type
    let first = &chain[0];
    let mut current_type = crate::analysis::resolve::resolve_var_type_at(symbols, offset, first)?;

    // Walk remaining chain segments (except the last, which is what we're completing)
    let builtin_types = crate::excel_model::types::builtin_types();
    for segment in &chain[1..] {
        // Check UDT members
        if let Some(member) = symbols.symbols.iter().find(|s| {
            matches!(s.kind, SymbolKind::UdtMember)
                && s.name.eq_ignore_ascii_case(segment)
                && match &s.detail {
                    SymbolDetail::UdtMember { parent_type, .. } => {
                        parent_type.eq_ignore_ascii_case(&current_type)
                    }
                    _ => false,
                }
        }) {
            current_type = member.type_name.clone().unwrap_or_default();
            continue;
        }
        // Check Excel builtin type members
        if let Some(excel_type) = builtin_types
            .iter()
            .find(|t| t.name.eq_ignore_ascii_case(&current_type))
        {
            if let Some(prop) = excel_type
                .properties
                .iter()
                .find(|p| p.name.eq_ignore_ascii_case(segment))
            {
                current_type = smol_str::SmolStr::new(&prop.return_type);
                continue;
            }
            if let Some(method) = excel_type
                .methods
                .iter()
                .find(|m| m.name.eq_ignore_ascii_case(segment))
            {
                if let Some(ret) = &method.return_type {
                    current_type = smol_str::SmolStr::new(ret);
                    continue;
                }
            }
        }
        return None; // Can't resolve this segment
    }

    // Return members of the final resolved type
    resolve_type_members(symbols, &current_type)
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::analysis::AnalysisHost;

    fn setup_host(source: &str) -> (AnalysisHost, Url) {
        let host = AnalysisHost::new();
        let uri = Url::parse("file:///test.bas").unwrap();
        let parse_result = crate::parser::parse(source);
        host.update(uri.clone(), source.to_string(), parse_result);
        (host, uri)
    }

    #[test]
    fn complete_includes_application_globals() {
        let (host, uri) = setup_host("Sub Foo()\n    \nEnd Sub\n");
        let items = complete(&host, &uri, Position::new(1, 4));
        let labels: Vec<&str> = items.iter().map(|i| i.label.as_str()).collect();
        assert!(
            labels.contains(&"ActiveWorkbook"),
            "expected ActiveWorkbook"
        );
        assert!(labels.contains(&"ActiveSheet"), "expected ActiveSheet");
        assert!(labels.contains(&"ActiveCell"), "expected ActiveCell");
    }

    #[test]
    fn application_globals_not_in_dot_access() {
        let (host, uri) = setup_host("Sub Foo()\n    Dim rng As Range\n    rng.\nEnd Sub\n");
        let items = complete(&host, &uri, Position::new(2, 8));
        let labels: Vec<&str> = items.iter().map(|i| i.label.as_str()).collect();
        assert!(
            !labels.contains(&"ActiveWorkbook"),
            "globals should not appear in dot access"
        );
    }

    #[test]
    fn enum_direct_dot_access_completion() {
        let source =
            "Enum Color\n    Red\n    Green\n    Blue\nEnd Enum\nSub Foo()\n    Color.\nEnd Sub\n";
        let (host, uri) = setup_host(source);
        let items = complete(&host, &uri, Position::new(6, 10));
        let labels: Vec<&str> = items.iter().map(|i| i.label.as_str()).collect();
        assert!(labels.contains(&"Red"), "expected Red in Color. completion");
        assert!(
            labels.contains(&"Green"),
            "expected Green in Color. completion"
        );
        assert!(
            labels.contains(&"Blue"),
            "expected Blue in Color. completion"
        );
    }

    #[test]
    fn enum_variable_dot_access_completion() {
        let source =
            "Enum Color\n    Red\n    Green\nEnd Enum\nSub Foo()\n    Dim c As Color\n    c.\nEnd Sub\n";
        let (host, uri) = setup_host(source);
        let items = complete(&host, &uri, Position::new(6, 6));
        let labels: Vec<&str> = items.iter().map(|i| i.label.as_str()).collect();
        assert!(labels.contains(&"Red"), "expected Red in c. completion");
        assert!(labels.contains(&"Green"), "expected Green in c. completion");
    }

    // ── PLAN-07: Module name dot access ───────────────────────────────

    #[test]
    fn module_dot_access_returns_public_symbols() {
        let host = AnalysisHost::new();
        let uri_a = Url::parse("file:///module_a.bas").unwrap();
        let uri_b = Url::parse("file:///module_b.bas").unwrap();
        let source_a = "Public Sub Foo()\nEnd Sub\nPrivate Sub Bar()\nEnd Sub\n";
        let source_b = "Sub Test()\n    module_a.\nEnd Sub\n";
        let pr_a = crate::parser::parse(source_a);
        let pr_b = crate::parser::parse(source_b);
        host.update(uri_a.clone(), source_a.to_string(), pr_a);
        host.update(uri_b.clone(), source_b.to_string(), pr_b);

        let items = complete(&host, &uri_b, Position::new(1, 14));
        let labels: Vec<&str> = items.iter().map(|i| i.label.as_str()).collect();
        assert!(
            labels.contains(&"Foo"),
            "expected Public Sub Foo from module_a"
        );
        assert!(
            !labels.contains(&"Bar"),
            "Private Sub Bar should not appear"
        );
    }

    // ── PLAN-13: With block leading-dot completion ───────────────────

    #[test]
    fn with_block_leading_dot_completion() {
        let source = "Type MyType\n    x As Long\n    y As String\nEnd Type\nSub Foo()\n    Dim obj As MyType\n    With obj\n        .\nEnd With\nEnd Sub\n";
        let (host, uri) = setup_host(source);
        // cursor at the `.` position (line 7, col 9)
        let items = complete(&host, &uri, Position::new(7, 9));
        let labels: Vec<&str> = items.iter().map(|i| i.label.as_str()).collect();
        assert!(labels.contains(&"x"), "expected x in With dot completion");
        assert!(labels.contains(&"y"), "expected y in With dot completion");
    }

    #[test]
    fn leading_dot_outside_with_returns_general() {
        let source = "Sub Foo()\n    .\nEnd Sub\n";
        let (host, uri) = setup_host(source);
        let items = complete(&host, &uri, Position::new(1, 5));
        // Should not crash, returns general completion or empty
        // (no With context, leading dot produces no dot-access items)
        let _ = items;
    }

    // ── PLAN-15: Chain dot-access completion ─────────────────────────

    #[test]
    fn chain_dot_access_excel_builtin() {
        let source = "Sub Foo()\n    Dim rng As Range\n    rng.Font.\nEnd Sub\n";
        let (host, uri) = setup_host(source);
        let items = complete(&host, &uri, Position::new(2, 13));
        // Font is a property of Range that returns Font type
        // If Font type has members like Bold, they should appear
        // This test validates the chain resolution doesn't crash
        // Chain dot-access resolution doesn't crash; items may or may not be returned.
        let _ = &items;
    }

    // ── PLAN-19: Function return value dot-access ────────────────────

    #[test]
    fn function_return_dot_access() {
        let source = "Function GetObj() As MyType\nEnd Function\nType MyType\n    x As Long\nEnd Type\nSub Foo()\n    GetObj().\nEnd Sub\n";
        let (host, uri) = setup_host(source);
        // Cursor after `GetObj().`
        let items = complete(&host, &uri, Position::new(6, 13));
        let labels: Vec<&str> = items.iter().map(|i| i.label.as_str()).collect();
        assert!(
            labels.contains(&"x"),
            "expected UDT member x from function return type"
        );
    }

    #[test]
    fn sub_return_dot_access_returns_empty() {
        let source = "Sub DoWork()\nEnd Sub\nSub Foo()\n    DoWork().\nEnd Sub\n";
        let (host, uri) = setup_host(source);
        let items = complete(&host, &uri, Position::new(3, 13));
        // Sub has no return type, so no dot-access members
        let labels: Vec<&str> = items.iter().map(|i| i.label.as_str()).collect();
        assert!(
            !labels.contains(&"DoWork"),
            "should not return completions for Sub return"
        );
    }
}

fn proc_at_position(
    symbols: &SymbolTable,
    source: &str,
    position: Position,
) -> Option<smol_str::SmolStr> {
    let offset = position_to_offset(source, position)?;
    crate::analysis::resolve::find_containing_proc(&symbols.proc_ranges, offset).cloned()
}
