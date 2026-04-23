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
    let cursor_proc = symbols
        .proc_ranges
        .iter()
        .find(|(_, r)| offset >= r.start as usize && offset <= r.end as usize)
        .map(|(name, _)| name.clone());

    let type_name = cursor_proc
        .as_ref()
        .and_then(|proc_name| {
            symbols.symbols.iter().find(|s| {
                s.name.eq_ignore_ascii_case(&var_name)
                    && matches!(s.kind, SymbolKind::Variable | SymbolKind::Parameter)
                    && s.proc_scope
                        .as_ref()
                        .is_some_and(|p| p.eq_ignore_ascii_case(proc_name))
            })
        })
        .or_else(|| {
            symbols.symbols.iter().find(|s| {
                s.name.eq_ignore_ascii_case(&var_name)
                    && matches!(s.kind, SymbolKind::Variable | SymbolKind::Parameter)
                    && s.proc_scope.is_none()
            })
        })
        .and_then(|s| s.type_name.clone())?;

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

pub fn goto_definition(
    host: &AnalysisHost,
    uri: &Url,
    position: Position,
) -> Option<GotoDefinitionResponse> {
    // Dot-access: if cursor is on the right side of `obj.Member`, jump to the member's def.
    if let Some(response) = host.with_source(uri, |symbols, source| {
        goto_def_dot_member(symbols, source, uri, position)
    }).flatten() {
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
            let containing_proc = cursor_offset.and_then(|off| {
                symbols
                    .proc_ranges
                    .iter()
                    .find(|(_, r)| off >= r.start as usize && off <= r.end as usize)
                    .map(|(name, _)| name.clone())
            });

            if let Some(ref proc_name) = containing_proc {
                matches
                    .iter()
                    .find(|s| {
                        s.proc_scope
                            .as_ref()
                            .map(|p| p.eq_ignore_ascii_case(proc_name))
                            .unwrap_or(false)
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
    let (other_uri, sym) = host.find_public_symbol_in_other_files(uri, &word)?;
    let range = host.with_source(&other_uri, |_, source| {
        resolve::text_range_to_lsp_range(source, sym.span)
    })?;
    Some(GotoDefinitionResponse::Scalar(Location::new(
        other_uri, range,
    )))
}
