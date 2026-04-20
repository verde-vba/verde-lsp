use tower_lsp::lsp_types::*;

use crate::analysis::resolve;
use crate::analysis::symbols::{ParameterInfo, Symbol, SymbolDetail, SymbolKind};
use crate::analysis::AnalysisHost;
use crate::parser::ast::ParameterPassing;
use crate::parser::ast::ProcedureKind;

pub fn hover(host: &AnalysisHost, uri: &Url, position: Position) -> Option<Hover> {
    // Try current file first, preferring symbols scoped to the cursor's procedure.
    let result = host.with_source(uri, |symbols, source| {
        let word = resolve::find_word_at_position(source, position)?;
        let matches = resolve::find_symbol_by_name(symbols, &word);

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
