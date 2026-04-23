use tower_lsp::lsp_types::*;

use crate::analysis::resolve;
use crate::analysis::symbols::{ParameterInfo, Symbol, SymbolDetail, SymbolKind};
use crate::analysis::AnalysisHost;
use crate::parser::ast::ParameterPassing;
use crate::parser::ast::ProcedureKind;

pub fn hover(host: &AnalysisHost, uri: &Url, position: Position) -> Option<Hover> {
    // Dot-access: if cursor is on the right side of `obj.Member`, resolve the member.
    if let Some(hover) = host
        .with_source(uri, |symbols, source| {
            hover_dot_member(symbols, source, position)
        })
        .flatten()
    {
        return Some(hover);
    }

    // Try current file first, preferring symbols scoped to the cursor's procedure.
    let result = host.with_source(uri, |symbols, source| {
        let word = resolve::find_word_at_position(source, position)?;
        let matches = resolve::find_symbol_by_name(symbols, &word);

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

        Some(symbol_to_hover(sym))
    });
    if let Some(Some(h)) = result {
        return Some(h);
    }

    // Fallback: cross-module public symbols
    let word = host.with_source(uri, |_, source| {
        resolve::find_word_at_position(source, position)
    })??;
    let (_other_uri, sym) = host.find_public_symbol_in_other_files(uri, &word)?;
    Some(symbol_to_hover(&sym))
}

/// Resolve the member on the right side of a dot access for hover.
fn hover_dot_member(
    symbols: &crate::analysis::symbols::SymbolTable,
    source: &str,
    position: Position,
) -> Option<Hover> {
    use crate::analysis::resolve::{parse_dot_access_at, position_to_offset};
    use crate::excel_model::types::builtin_types;

    let offset = position_to_offset(source, position)?;
    let (var_name, _member_partial) = parse_dot_access_at(source, offset)?;

    // The member_partial may be empty if cursor is right after dot.
    // Also need to check if cursor word extends beyond what parse_dot_access_at returned.
    let word = resolve::find_word_at_position(source, position)?;

    // We're in dot context; the word at cursor is the member name.
    let member_name = word;

    // `Me.member` — look for module-level symbols.
    if var_name.eq_ignore_ascii_case("Me") {
        let sym = symbols
            .symbols
            .iter()
            .find(|s| s.proc_scope.is_none() && s.name.eq_ignore_ascii_case(&member_name))?;
        return Some(symbol_to_hover(sym));
    }

    let type_name = resolve::resolve_var_type_at(symbols, offset, &var_name)?;

    // Check UDT members.
    if let Some(sym) = symbols.symbols.iter().find(|s| {
        matches!(s.kind, SymbolKind::UdtMember)
            && s.name.eq_ignore_ascii_case(&member_name)
            && match &s.detail {
                SymbolDetail::UdtMember { parent_type, .. } => {
                    parent_type.eq_ignore_ascii_case(&type_name)
                }
                _ => false,
            }
    }) {
        return Some(symbol_to_hover(sym));
    }

    // Check Excel builtin type members.
    if let Some(excel_type) = builtin_types()
        .iter()
        .find(|t| t.name.eq_ignore_ascii_case(&type_name))
    {
        if let Some(prop) = excel_type
            .properties
            .iter()
            .find(|p| p.name.eq_ignore_ascii_case(&member_name))
        {
            let sig = format!("{}: {}", prop.name, prop.return_type);
            return Some(Hover {
                contents: HoverContents::Markup(MarkupContent {
                    kind: MarkupKind::Markdown,
                    value: format!("```vba\n{sig}\n```"),
                }),
                range: None,
            });
        }
        if let Some(method) = excel_type
            .methods
            .iter()
            .find(|m| m.name.eq_ignore_ascii_case(&member_name))
        {
            let ret = method
                .return_type
                .as_ref()
                .map(|t| format!(" As {t}"))
                .unwrap_or_default();
            let sig = format!("Sub {}(){ret}", method.name);
            return Some(Hover {
                contents: HoverContents::Markup(MarkupContent {
                    kind: MarkupKind::Markdown,
                    value: format!("```vba\n{sig}\n```"),
                }),
                range: None,
            });
        }
    }

    None
}

fn symbol_to_hover(sym: &Symbol) -> Hover {
    let signature = match &sym.detail {
        SymbolDetail::Procedure {
            kind,
            params,
            return_type,
        } => {
            let kind_str = match kind {
                ProcedureKind::Sub => "Sub",
                ProcedureKind::Function => "Function",
                ProcedureKind::PropertyGet => "Property Get",
                ProcedureKind::PropertyLet => "Property Let",
                ProcedureKind::PropertySet => "Property Set",
            };
            let param_list = format_params(params);
            let ret = return_type
                .as_ref()
                .map(|t| format!(" As {}", t))
                .unwrap_or_default();
            format!("{} {}({}){}", kind_str, sym.name, param_list, ret)
        }
        SymbolDetail::Parameter {
            type_name,
            passing,
            is_optional,
        } => {
            let prefix = match passing {
                ParameterPassing::ByVal => "ByVal ",
                ParameterPassing::ByRef => "ByRef ",
            };
            let opt = if *is_optional { "Optional " } else { "" };
            let type_str = type_name
                .as_ref()
                .map(|t| format!(" As {}", t))
                .unwrap_or_default();
            format!("{}{}{}{}", opt, prefix, sym.name, type_str)
        }
        SymbolDetail::Variable { .. } => {
            let type_str = sym
                .type_name
                .as_ref()
                .map(|t| format!(" As {}", t))
                .unwrap_or_else(|| " As Variant".to_string());
            match sym.kind {
                SymbolKind::Constant => format!("Const {}{}", sym.name, type_str),
                _ => format!("Dim {}{}", sym.name, type_str),
            }
        }
        SymbolDetail::TypeDef { .. } => format!("Type {}", sym.name),
        SymbolDetail::EnumDef { members } => {
            format!("Enum {} ({} members)", sym.name, members.len())
        }
        SymbolDetail::EnumMember { parent_enum, value } => {
            format!("{}.{} = {}", parent_enum, sym.name, value)
        }
        SymbolDetail::UdtMember { type_name, .. } => {
            format!("{} As {}", sym.name, type_name)
        }
    };

    Hover {
        contents: HoverContents::Markup(MarkupContent {
            kind: MarkupKind::Markdown,
            value: format!("```vba\n{}\n```", signature),
        }),
        range: None,
    }
}

fn format_params(params: &[ParameterInfo]) -> String {
    params
        .iter()
        .map(|p| match &p.type_name {
            Some(t) => format!("{} As {}", p.name, t),
            None => p.name.to_string(),
        })
        .collect::<Vec<_>>()
        .join(", ")
}
