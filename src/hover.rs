use tower_lsp::lsp_types::*;

use crate::analysis::resolve;
use crate::analysis::symbols::{ParameterInfo, Symbol, SymbolDetail, SymbolKind};
use crate::analysis::AnalysisHost;
use crate::parser::ast::ProcedureKind;

pub fn hover(host: &AnalysisHost, uri: &Url, position: Position) -> Option<Hover> {
    // Try current file first
    let result = host.with_source(uri, |symbols, source| {
        let word = resolve::find_word_at_position(source, position)?;
        let matches = resolve::find_symbol_by_name(symbols, &word);
        let sym = matches.first()?;
        Some(symbol_to_hover(sym))
    });
    if let Some(Some(h)) = result {
        return Some(h);
    }

    // Fallback: cross-module public symbols
    let word = host.with_source(uri, |_, source| resolve::find_word_at_position(source, position))??;
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
        SymbolDetail::None => sym.name.to_string(),
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
