use tower_lsp::lsp_types::*;

use crate::analysis::resolve;
use crate::analysis::AnalysisHost;

pub fn goto_definition(
    host: &AnalysisHost,
    uri: &Url,
    position: Position,
) -> Option<GotoDefinitionResponse> {
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
