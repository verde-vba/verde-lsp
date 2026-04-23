use tower_lsp::lsp_types::*;

use crate::analysis::resolve;
use crate::analysis::symbols::SymbolKind;
use crate::analysis::AnalysisHost;

pub fn goto_type_definition(
    host: &AnalysisHost,
    uri: &Url,
    position: Position,
) -> Option<GotoDefinitionResponse> {
    // Find the variable/parameter at cursor and resolve its type_name.
    let type_name = host.with_source(uri, |symbols, source| {
        let word = resolve::find_word_at_position(source, position)?;
        let matches = resolve::find_symbol_by_name(symbols, &word);

        // Prefer a proc-scoped match at the cursor position.
        let cursor_offset = resolve::position_to_offset(source, position);
        let containing_proc = cursor_offset.and_then(|off| {
            symbols
                .proc_ranges
                .iter()
                .find(|(_, r)| off >= r.start as usize && off <= r.end as usize)
                .map(|(name, _)| name.clone())
        });

        let sym = if let Some(ref proc_name) = containing_proc {
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
        }?;

        sym.type_name.clone()
    })??;

    // Search for TypeDef / EnumDef with that name in current file.
    let local_result = host.with_source(uri, |symbols, source| {
        symbols
            .symbols
            .iter()
            .find(|s| {
                matches!(s.kind, SymbolKind::TypeDef | SymbolKind::EnumDef)
                    && s.name.eq_ignore_ascii_case(&type_name)
            })
            .map(|s| {
                let range = resolve::text_range_to_lsp_range(source, s.span);
                GotoDefinitionResponse::Scalar(Location::new(uri.clone(), range))
            })
    });
    if let Some(Some(r)) = local_result {
        return Some(r);
    }

    // Cross-module fallback.
    let (other_uri, sym) = host.find_public_symbol_in_other_files(uri, &type_name)?;
    if !matches!(sym.kind, SymbolKind::TypeDef | SymbolKind::EnumDef) {
        return None;
    }
    let range = host.with_source(&other_uri, |_, source| {
        resolve::text_range_to_lsp_range(source, sym.span)
    })?;
    Some(GotoDefinitionResponse::Scalar(Location::new(
        other_uri, range,
    )))
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
    fn type_definition_jumps_to_typedef() {
        let source =
            "Type MyType\n    x As Long\nEnd Type\nSub Foo()\n    Dim f As MyType\nEnd Sub\n";
        let (host, uri) = setup_host(source);
        // Cursor on 'f' at line 4, col 8
        let result = goto_type_definition(&host, &uri, Position::new(4, 8));
        assert!(result.is_some(), "expected type definition jump for 'f'");
        if let Some(GotoDefinitionResponse::Scalar(loc)) = result {
            assert_eq!(loc.range.start.line, 0, "Type MyType starts at line 0");
        }
    }

    #[test]
    fn type_definition_jumps_to_enum() {
        let source =
            "Enum Color\n    Red\n    Green\nEnd Enum\nSub Foo()\n    Dim c As Color\nEnd Sub\n";
        let (host, uri) = setup_host(source);
        let result = goto_type_definition(&host, &uri, Position::new(5, 8));
        assert!(result.is_some(), "expected type definition jump for 'c'");
        if let Some(GotoDefinitionResponse::Scalar(loc)) = result {
            assert_eq!(loc.range.start.line, 0, "Enum Color starts at line 0");
        }
    }

    #[test]
    fn builtin_type_returns_none() {
        let source = "Sub Foo()\n    Dim x As Long\nEnd Sub\n";
        let (host, uri) = setup_host(source);
        let result = goto_type_definition(&host, &uri, Position::new(1, 8));
        assert!(result.is_none(), "builtin type Long should return None");
    }

    #[test]
    fn untyped_variable_returns_none() {
        let source = "Sub Foo()\n    Dim x\nEnd Sub\n";
        let (host, uri) = setup_host(source);
        let result = goto_type_definition(&host, &uri, Position::new(1, 8));
        assert!(result.is_none(), "untyped variable should return None");
    }
}
