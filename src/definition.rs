use tower_lsp::lsp_types::*;

use crate::analysis::resolve;
use crate::analysis::symbols::{SymbolDetail, SymbolKind, SymbolTable};
use crate::analysis::AnalysisHost;

/// Resolve the member on the right side of a dot access for goto-definition.
fn goto_def_dot_member(
    symbols: &SymbolTable,
    source: &str,
    uri: &Url,
    position: Position,
) -> Option<GotoDefinitionResponse> {
    let offset = resolve::position_to_offset(source, position)?;
    let (var_name, _member_partial) = resolve::parse_dot_access_at(source, offset)?;
    let member_name = resolve::find_word_at_position(source, position)?;

    // `Me.member` — jump to definition of module-level symbol.
    if var_name.eq_ignore_ascii_case("Me") {
        let sym = symbols
            .symbols
            .iter()
            .find(|s| s.proc_scope.is_none() && s.name.eq_ignore_ascii_case(&member_name))?;
        let range = resolve::text_range_to_lsp_range(source, sym.span);
        return Some(GotoDefinitionResponse::Scalar(Location::new(
            uri.clone(),
            range,
        )));
    }

    // Resolve the variable's type.
    let type_name = resolve::resolve_var_type_at(symbols, offset, &var_name)?;

    // UDT member → jump to the member declaration inside the TypeDef.
    let sym = symbols.symbols.iter().find(|s| {
        matches!(s.kind, SymbolKind::UdtMember)
            && s.name.eq_ignore_ascii_case(&member_name)
            && match &s.detail {
                SymbolDetail::UdtMember { parent_type, .. } => {
                    parent_type.eq_ignore_ascii_case(&type_name)
                }
                _ => false,
            }
    })?;

    let range = resolve::text_range_to_lsp_range(source, sym.span);
    Some(GotoDefinitionResponse::Scalar(Location::new(
        uri.clone(),
        range,
    )))
}

/// Cross-module dot-access: `ModuleName.Member` where cursor is on `Member`.
/// Resolves `ModuleName` as a filename stem and jumps to `Member`'s definition
/// in that module.
fn goto_def_dot_module(
    host: &AnalysisHost,
    uri: &Url,
    position: Position,
) -> Option<GotoDefinitionResponse> {
    let (var_name, member_name) = host.with_source(uri, |_symbols, source| {
        let offset = resolve::position_to_offset(source, position)?;
        let (var, _partial) = resolve::parse_dot_access_at(source, offset)?;
        let word = resolve::find_word_at_position(source, position)?;
        Some((var, word))
    })??;

    let (mod_uri, public_syms) = host.find_module_by_name(uri, &var_name)?;
    let sym = public_syms
        .iter()
        .find(|s| s.name.eq_ignore_ascii_case(&member_name))?;
    let range = host.with_source(&mod_uri, |_, source| {
        resolve::text_range_to_lsp_range(source, sym.span)
    })?;
    Some(GotoDefinitionResponse::Scalar(Location::new(
        mod_uri, range,
    )))
}

pub fn goto_definition(
    host: &AnalysisHost,
    uri: &Url,
    position: Position,
) -> Option<GotoDefinitionResponse> {
    // Dot-access: if cursor is on the right side of `obj.Member`, jump to the member's def.
    if let Some(response) = host
        .with_source(uri, |symbols, source| {
            goto_def_dot_member(symbols, source, uri, position)
        })
        .flatten()
    {
        return Some(response);
    }
    // Cross-module dot-access: `ModuleName.Member` → jump to Member in that module.
    if let Some(response) = goto_def_dot_module(host, uri, position) {
        return Some(response);
    }

    // Try current file first, preferring symbols scoped to the cursor's procedure.
    let result = host.with_source(uri, |symbols, source| {
        let word = resolve::find_word_at_position(source, position)?;
        let matches = resolve::find_symbol_by_name(symbols, &word);

        // Among all matches, prefer a symbol whose proc_scope matches the
        // procedure that contains the cursor. This makes goto-def scope-aware
        // for local variables and parameters when two procs share a name.
        let sym = {
            let cursor_offset = resolve::position_to_offset(source, position);
            let containing_proc = cursor_offset
                .and_then(|off| resolve::find_containing_proc(&symbols.proc_ranges, off));

            if let Some(proc_name) = containing_proc {
                matches
                    .iter()
                    .find(|s| {
                        s.proc_scope
                            .as_ref()
                            .is_some_and(|p| p.eq_ignore_ascii_case(proc_name))
                    })
                    .copied()
                    .or_else(|| matches.first().copied())
            } else {
                matches.first().copied()
            }
        }?;

        let range = resolve::text_range_to_lsp_range(source, sym.span);
        Some(GotoDefinitionResponse::Scalar(Location::new(
            uri.clone(),
            range,
        )))
    });
    if let Some(Some(r)) = result {
        return Some(r);
    }

    // Fallback: cross-module public symbols
    let word = host.with_source(uri, |_, source| {
        resolve::find_word_at_position(source, position)
    })??;
    if let Some((other_uri, sym)) = host.find_public_symbol_in_other_files(uri, &word) {
        let range = host.with_source(&other_uri, |_, source| {
            resolve::text_range_to_lsp_range(source, sym.span)
        })?;
        return Some(GotoDefinitionResponse::Scalar(Location::new(
            other_uri, range,
        )));
    }

    // Fallback: module name (e.g. `Utils` → jump to Utils.bas top)
    if let Some((mod_uri, _public_syms)) = host.find_module_by_name(uri, &word) {
        let range = Range::new(Position::new(0, 0), Position::new(0, 0));
        return Some(GotoDefinitionResponse::Scalar(Location::new(
            mod_uri, range,
        )));
    }

    None
}
