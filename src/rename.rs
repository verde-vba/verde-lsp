use std::collections::HashMap;
use tower_lsp::lsp_types::*;

use crate::analysis::resolve;
use crate::analysis::AnalysisHost;
use crate::parser::ast::{TextRange, Visibility};

pub fn rename(
    host: &AnalysisHost,
    uri: &Url,
    position: Position,
    new_name: &str,
) -> Option<WorkspaceEdit> {
    // Reject renaming to VBA keywords or builtin type names
    if crate::vba_builtins::KEYWORDS
        .iter()
        .any(|kw| kw.eq_ignore_ascii_case(new_name))
        || crate::vba_builtins::BUILTIN_TYPES
            .iter()
            .any(|t| t.eq_ignore_ascii_case(new_name))
    {
        return None;
    }

    // Determine the word at cursor, decide cross-file scope, and compute an
    // optional intra-file proc constraint for scope-aware local-variable rename.
    let (word, cross_file, proc_constraint) = host.with_source(uri, |symbols, source| {
        let word = resolve::find_word_at_position(source, position)?;

        // Determine the procedure byte-range that should bound rename edits.
        // Step 1: symbol directly at cursor (declaration site) → use its proc_scope.
        // Step 2: cursor on a use site → find the containing procedure, then verify
        //         the word is actually a local symbol there before constraining.
        let proc_constraint: Option<TextRange> = {
            let cursor_sym = resolve::find_symbol_at_position(symbols, source, position);
            if let Some(sym) = cursor_sym {
                sym.proc_scope.as_ref().and_then(|proc_name| {
                    symbols
                        .proc_ranges
                        .iter()
                        .find(|(n, _)| n.eq_ignore_ascii_case(proc_name))
                        .map(|(_, r)| *r)
                })
            } else {
                resolve::position_to_offset(source, position).and_then(|offset| {
                    let containing = symbols
                        .proc_ranges
                        .iter()
                        .find(|(_, r)| offset >= r.start as usize && offset <= r.end as usize);
                    containing.and_then(|(proc_name, proc_range)| {
                        let is_local_here = resolve::find_symbol_by_name(symbols, &word)
                            .iter()
                            .any(|s| {
                                s.proc_scope
                                    .as_ref()
                                    .map(|p| p.eq_ignore_ascii_case(proc_name))
                                    .unwrap_or(false)
                            });
                        if is_local_here {
                            Some(*proc_range)
                        } else {
                            None
                        }
                    })
                })
            }
        };

        let local_syms = resolve::find_symbol_by_name(symbols, &word);
        if !local_syms.is_empty() {
            let is_public_module_level = local_syms
                .iter()
                .any(|s| s.visibility == Visibility::Public && s.proc_scope.is_none());
            Some((word, is_public_module_level, proc_constraint))
        } else {
            let found_cross = host.find_public_symbol_in_other_files(uri, &word).is_some();
            if found_cross {
                Some((word, true, proc_constraint))
            } else {
                None
            }
        }
    })??;

    let files: Vec<(Url, String)> = if cross_file {
        host.all_file_sources()
    } else {
        host.with_source(uri, |_, source| vec![(uri.clone(), source.to_string())])
            .unwrap_or_default()
    };

    // Collect edits, filtering by proc_constraint for the cursor's file when set.
    let mut changes: HashMap<Url, Vec<TextEdit>> = HashMap::new();
    for (file_uri, source) in files {
        let edits: Vec<TextEdit> = resolve::find_all_word_occurrences(&source, &word)
            .into_iter()
            .filter(|r| match proc_constraint {
                Some(constraint) if file_uri == *uri => {
                    r.start >= constraint.start && r.end <= constraint.end
                }
                _ => true,
            })
            .map(|range| {
                TextEdit::new(
                    resolve::text_range_to_lsp_range(&source, range),
                    new_name.to_string(),
                )
            })
            .collect();
        if !edits.is_empty() {
            changes.insert(file_uri, edits);
        }
    }

    if changes.is_empty() {
        None
    } else {
        Some(WorkspaceEdit {
            changes: Some(changes),
            ..Default::default()
        })
    }
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
    fn rename_to_keyword_is_rejected() {
        let (host, uri) = setup_host("Sub Foo()\n    Dim x As Long\n    x = 1\nEnd Sub\n");
        let result = rename(&host, &uri, Position::new(1, 8), "Sub");
        assert!(
            result.is_none(),
            "renaming to keyword 'Sub' should be rejected"
        );
    }

    #[test]
    fn rename_to_builtin_type_is_rejected() {
        let (host, uri) = setup_host("Sub Foo()\n    Dim x As Long\n    x = 1\nEnd Sub\n");
        let result = rename(&host, &uri, Position::new(1, 8), "Long");
        assert!(
            result.is_none(),
            "renaming to builtin type should be rejected"
        );
    }

    #[test]
    fn rename_to_keyword_case_insensitive() {
        let (host, uri) = setup_host("Sub Foo()\n    Dim x As Long\n    x = 1\nEnd Sub\n");
        assert!(rename(&host, &uri, Position::new(1, 8), "sub").is_none());
        assert!(rename(&host, &uri, Position::new(1, 8), "DIM").is_none());
        assert!(rename(&host, &uri, Position::new(1, 8), "long").is_none());
    }

    #[test]
    fn rename_to_valid_name_is_allowed() {
        let (host, uri) = setup_host("Sub Foo()\n    Dim x As Long\n    x = 1\nEnd Sub\n");
        let result = rename(&host, &uri, Position::new(1, 8), "y");
        assert!(result.is_some(), "renaming to 'y' should be allowed");
    }
}
