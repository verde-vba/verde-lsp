use std::collections::HashMap;

use tower_lsp::lsp_types::*;

use crate::analysis::resolve::{
    find_all_word_occurrences, offset_to_position, text_range_to_lsp_range,
};
use crate::analysis::symbols::SymbolKind as AK;
use crate::analysis::AnalysisHost;

fn to_lsp_kind(kind: &AK) -> SymbolKind {
    match kind {
        AK::Function => SymbolKind::FUNCTION,
        AK::Procedure => SymbolKind::FUNCTION,
        AK::Property => SymbolKind::PROPERTY,
        _ => SymbolKind::FUNCTION,
    }
}

fn is_procedure_kind(kind: &AK) -> bool {
    matches!(kind, AK::Procedure | AK::Function | AK::Property)
}

pub fn prepare_call_hierarchy(
    host: &AnalysisHost,
    uri: &Url,
    position: Position,
) -> Option<Vec<CallHierarchyItem>> {
    host.with_source(uri, |symbols, source| {
        let sym = crate::analysis::resolve::find_symbol_at_position(symbols, source, position)?;
        if !is_procedure_kind(&sym.kind) {
            return None;
        }
        let range = text_range_to_lsp_range(source, sym.span);
        let item = CallHierarchyItem {
            name: sym.name.to_string(),
            kind: to_lsp_kind(&sym.kind),
            tags: None,
            detail: None,
            uri: uri.clone(),
            range,
            selection_range: range,
            data: Some(serde_json::json!({
                "name": sym.name.as_str(),
                "uri": uri.as_str()
            })),
        };
        Some(vec![item])
    })
    .flatten()
}

pub fn incoming_calls(
    host: &AnalysisHost,
    item: &CallHierarchyItem,
) -> Vec<CallHierarchyIncomingCall> {
    let target_name = &item.name;
    let mut result: HashMap<String, (CallHierarchyItem, Vec<Range>)> = HashMap::new();

    for (file_uri, source) in host.all_file_sources() {
        let occurrences = find_all_word_occurrences(&source, target_name);
        if occurrences.is_empty() {
            continue;
        }

        let Some((decl_spans, proc_ranges)) = host.with_source(&file_uri, |symbols, _| {
            let decl_spans: Vec<_> = symbols
                .symbols
                .iter()
                .filter(|s| s.name.eq_ignore_ascii_case(target_name) && is_procedure_kind(&s.kind))
                .map(|s| s.span)
                .collect();
            let proc_ranges = symbols.proc_ranges.clone();
            (decl_spans, proc_ranges)
        }) else {
            continue;
        };

        for occ in occurrences {
            // Skip declaration spans (where the name is defined)
            if decl_spans.iter().any(|ds| ds.start == occ.start) {
                continue;
            }
            // Find the enclosing procedure
            let Some((caller_name, caller_span)) = proc_ranges
                .iter()
                .find(|(_, span)| span.start <= occ.start && occ.end <= span.end)
                .map(|(n, s)| (n.clone(), *s))
            else {
                continue;
            };

            let call_range = text_range_to_lsp_range(&source, occ);
            let caller_range = text_range_to_lsp_range(&source, caller_span);
            let key = format!("{}#{}", file_uri, caller_name);

            let entry = result.entry(key).or_insert_with(|| {
                let sel_range = host
                    .with_source(&file_uri, |symbols, _| {
                        symbols
                            .symbols
                            .iter()
                            .find(|s| {
                                s.name.eq_ignore_ascii_case(&caller_name)
                                    && is_procedure_kind(&s.kind)
                            })
                            .map(|s| text_range_to_lsp_range(&source, s.span))
                            .unwrap_or(caller_range)
                    })
                    .unwrap_or(caller_range);

                let caller_lsp_kind = host
                    .with_source(&file_uri, |symbols, _| {
                        symbols
                            .symbols
                            .iter()
                            .find(|s| {
                                s.name.eq_ignore_ascii_case(&caller_name)
                                    && is_procedure_kind(&s.kind)
                            })
                            .map(|s| to_lsp_kind(&s.kind))
                            .unwrap_or(SymbolKind::FUNCTION)
                    })
                    .unwrap_or(SymbolKind::FUNCTION);

                let from = CallHierarchyItem {
                    name: caller_name.to_string(),
                    kind: caller_lsp_kind,
                    tags: None,
                    detail: None,
                    uri: file_uri.clone(),
                    range: caller_range,
                    selection_range: sel_range,
                    data: None,
                };
                (from, Vec::new())
            });
            entry.1.push(call_range);
        }
    }

    result
        .into_values()
        .map(|(from, from_ranges)| CallHierarchyIncomingCall { from, from_ranges })
        .collect()
}

pub fn outgoing_calls(
    host: &AnalysisHost,
    item: &CallHierarchyItem,
) -> Vec<CallHierarchyOutgoingCall> {
    let caller_name = &item.name;
    let caller_uri = &item.uri;

    // Get the caller's procedure body span
    let Some(body_span) = host
        .with_source(caller_uri, |symbols, _| {
            symbols
                .proc_ranges
                .iter()
                .find(|(n, _)| n.eq_ignore_ascii_case(caller_name))
                .map(|(_, span)| *span)
        })
        .flatten()
    else {
        return Vec::new();
    };

    let Some(caller_full_source) = host.with_source(caller_uri, |_, source| source.to_string())
    else {
        return Vec::new();
    };

    let body_start = body_span.start as usize;
    let body_end = (body_span.end as usize).min(caller_full_source.len());
    let body_source = &caller_full_source[body_start..body_end];

    // Collect all known procedure names across all files (excluding self)
    let mut known_procs: Vec<(Url, String, crate::parser::ast::TextRange, AK)> = Vec::new();
    for (file_uri, _) in host.all_file_sources() {
        if let Some(procs) = host.with_source(&file_uri, |symbols, _| {
            symbols
                .symbols
                .iter()
                .filter(|s| is_procedure_kind(&s.kind))
                .map(|s| (file_uri.clone(), s.name.to_string(), s.span, s.kind.clone()))
                .collect::<Vec<_>>()
        }) {
            known_procs.extend(procs);
        }
    }

    let mut result = Vec::new();
    for (proc_uri, proc_name, proc_span, proc_kind) in &known_procs {
        // Skip self-reference
        if proc_name.eq_ignore_ascii_case(caller_name) && proc_uri == caller_uri {
            continue;
        }
        let occs_in_body = find_all_word_occurrences(body_source, proc_name);
        if occs_in_body.is_empty() {
            continue;
        }

        let Some(proc_source) = host.with_source(proc_uri, |_, s| s.to_string()) else {
            continue;
        };
        let to_range = text_range_to_lsp_range(&proc_source, *proc_span);
        let to = CallHierarchyItem {
            name: proc_name.clone(),
            kind: to_lsp_kind(proc_kind),
            tags: None,
            detail: None,
            uri: proc_uri.clone(),
            range: to_range,
            selection_range: to_range,
            data: None,
        };

        // Translate body-relative offsets back to absolute source positions
        let from_ranges = occs_in_body
            .iter()
            .map(|occ| {
                let abs_start = body_start + occ.start as usize;
                let abs_end = body_start + occ.end as usize;
                let start_pos = offset_to_position(&caller_full_source, abs_start);
                let end_pos = offset_to_position(&caller_full_source, abs_end);
                Range::new(start_pos, end_pos)
            })
            .collect();

        result.push(CallHierarchyOutgoingCall { to, from_ranges });
    }
    result
}
