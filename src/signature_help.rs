use tower_lsp::lsp_types::*;

use crate::analysis::resolve;
use crate::analysis::symbols::{ParameterInfo, SymbolDetail};
use crate::analysis::AnalysisHost;
use crate::parser::ast::ProcedureKind;

/// Return `SignatureHelp` when the cursor is inside a function call's argument list.
///
/// Uses backward text scan from the cursor to find the enclosing `(` and the
/// function name preceding it, then resolves the function in the symbol table.
pub fn signature_help(host: &AnalysisHost, uri: &Url, position: Position) -> Option<SignatureHelp> {
    let (func_name, active_param) = host.with_source(uri, |_, source| {
        let offset = resolve::position_to_offset(source, position)?;
        find_call_context(source, offset)
    })??;

    // Try current file first.
    let sig = host
        .with_source(uri, |symbols, _| {
            symbols
                .symbols
                .iter()
                .find(|s| s.proc_scope.is_none() && s.name.eq_ignore_ascii_case(&func_name))
                .and_then(|sym| build_signature_info(sym.name.as_str(), &sym.detail))
        })
        .flatten()
        .or_else(|| {
            // Cross-module fallback.
            let (_, sym) = host.find_public_symbol_in_other_files(uri, &func_name)?;
            build_signature_info(sym.name.as_str(), &sym.detail)
        })
        .or_else(|| {
            // Builtin function fallback.
            build_builtin_signature_info(&func_name)
        })?;

    let param_count = sig.parameters.as_ref().map_or(0, |p| p.len());
    let clamped = if param_count > 0 {
        Some((active_param as u32).min(param_count as u32 - 1))
    } else {
        None
    };

    Some(SignatureHelp {
        signatures: vec![sig],
        active_signature: Some(0),
        active_parameter: clamped,
    })
}

/// Scan backward from `cursor_offset` to find the enclosing function call.
///
/// Returns `(function_name, active_parameter_index)` or `None` if the cursor
/// is not inside a call argument list.
fn find_call_context(source: &str, cursor_offset: usize) -> Option<(String, usize)> {
    let bytes = source.as_bytes();
    let mut depth: i32 = 0;
    let mut comma_count: usize = 0;
    let mut i = cursor_offset;

    while i > 0 {
        i -= 1;
        match bytes[i] {
            b')' => depth += 1,
            b'(' => {
                if depth == 0 {
                    let paren_pos = i;
                    // Skip spaces between identifier and `(`.
                    while i > 0 && bytes[i - 1] == b' ' {
                        i -= 1;
                    }
                    // Read identifier backward.
                    let name_end = i;
                    while i > 0 && is_ident_byte(bytes[i - 1]) {
                        i -= 1;
                    }
                    let name_start = i;
                    if name_start == name_end || name_start == paren_pos {
                        return None;
                    }
                    let name = source[name_start..name_end].to_string();
                    if name.is_empty() {
                        return None;
                    }
                    return Some((name, comma_count));
                } else {
                    depth -= 1;
                }
            }
            b',' if depth == 0 => comma_count += 1,
            b'\n' | b'\r' => {
                // Check if this newline is preceded by `_` (line continuation).
                // There may be spaces/tabs between `_` and the newline.
                let mut j = i;
                while j > 0 && matches!(bytes[j - 1], b' ' | b'\t' | b'\r') {
                    j -= 1;
                }
                if j > 0 && bytes[j - 1] == b'_' {
                    continue;
                }
                return None;
            }
            _ => {}
        }
    }
    None
}

fn is_ident_byte(b: u8) -> bool {
    b.is_ascii_alphanumeric() || b == b'_'
}

fn build_signature_info(name: &str, detail: &SymbolDetail) -> Option<SignatureInformation> {
    let SymbolDetail::Procedure {
        kind,
        params,
        return_type,
    } = detail
    else {
        return None;
    };

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
        .map(|t| format!(" As {t}"))
        .unwrap_or_default();
    let label = format!("{kind_str} {name}({param_list}){ret}");

    let parameters: Vec<ParameterInformation> = params
        .iter()
        .map(|p| ParameterInformation {
            label: ParameterLabel::Simple(format_param(p)),
            documentation: None,
        })
        .collect();

    Some(SignatureInformation {
        label,
        documentation: None,
        parameters: if parameters.is_empty() {
            None
        } else {
            Some(parameters)
        },
        active_parameter: None,
    })
}

fn build_builtin_signature_info(name: &str) -> Option<SignatureInformation> {
    use crate::vba_builtins::BUILTIN_SIGNATURES;

    let sig = BUILTIN_SIGNATURES
        .iter()
        .find(|s| s.name.eq_ignore_ascii_case(name))?;

    let param_strs: Vec<String> = sig
        .params
        .iter()
        .map(|(pname, ptype, optional)| {
            if *optional {
                format!("[{pname} As {ptype}]")
            } else {
                format!("{pname} As {ptype}")
            }
        })
        .collect();
    let param_list = param_strs.join(", ");

    let ret = sig
        .return_type
        .map(|t| format!(" As {t}"))
        .unwrap_or_default();
    let label = format!("Function {}({param_list}){ret}", sig.name);

    let parameters: Vec<ParameterInformation> = sig
        .params
        .iter()
        .map(|(pname, ptype, optional)| {
            let label_str = if *optional {
                format!("[{pname} As {ptype}]")
            } else {
                format!("{pname} As {ptype}")
            };
            ParameterInformation {
                label: ParameterLabel::Simple(label_str),
                documentation: None,
            }
        })
        .collect();

    Some(SignatureInformation {
        label,
        documentation: None,
        parameters: if parameters.is_empty() {
            None
        } else {
            Some(parameters)
        },
        active_parameter: None,
    })
}

fn format_params(params: &[ParameterInfo]) -> String {
    params
        .iter()
        .map(format_param)
        .collect::<Vec<_>>()
        .join(", ")
}

fn format_param(p: &ParameterInfo) -> String {
    match &p.type_name {
        Some(t) => format!("{} As {t}", p.name),
        None => p.name.to_string(),
    }
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn find_call_context_single_line() {
        let source = "Foo(a, b)";
        let result = find_call_context(source, 7);
        assert_eq!(result, Some(("Foo".to_string(), 1)));
    }

    #[test]
    fn find_call_context_returns_none_outside_parens() {
        let source = "x = 1";
        let result = find_call_context(source, 3);
        assert!(result.is_none());
    }

    #[test]
    fn find_call_context_with_line_continuation() {
        let source = "Foo(a, _\n    b, _\n    c)";
        let c_pos = source.find('c').unwrap();
        let result = find_call_context(source, c_pos);
        assert!(
            result.is_some(),
            "expected call context with line continuation"
        );
        let (name, param) = result.unwrap();
        assert_eq!(name, "Foo");
        assert_eq!(param, 2);
    }

    #[test]
    fn find_call_context_line_continuation_second_line() {
        let source = "Foo(a, _\n    b, _\n    c)";
        let b_pos = source.find('b').unwrap();
        let result = find_call_context(source, b_pos);
        assert!(result.is_some(), "expected call context on second line");
        let (name, param) = result.unwrap();
        assert_eq!(name, "Foo");
        assert_eq!(param, 1);
    }

    #[test]
    fn find_call_context_no_continuation_returns_none() {
        let source = "Foo(a,\n    b)";
        let b_pos = source.find('b').unwrap();
        let result = find_call_context(source, b_pos);
        assert!(result.is_none(), "newline without _ should return None");
    }

    #[test]
    fn find_call_context_continuation_with_spaces_before_underscore() {
        let source = "Foo(a,   _\n    b)";
        let b_pos = source.find('b').unwrap();
        let result = find_call_context(source, b_pos);
        assert!(
            result.is_some(),
            "expected call context with spaces before _"
        );
        let (name, param) = result.unwrap();
        assert_eq!(name, "Foo");
        assert_eq!(param, 1);
    }

    // ── PLAN-09: Builtin signature help ───────────────────────────────

    #[test]
    fn builtin_msgbox_signature() {
        let sig = build_builtin_signature_info("MsgBox");
        assert!(sig.is_some(), "expected MsgBox signature");
        let sig = sig.unwrap();
        assert!(sig.label.contains("MsgBox"), "label should contain MsgBox");
        assert!(sig.parameters.is_some());
        let params = sig.parameters.unwrap();
        assert_eq!(params.len(), 5, "MsgBox has 5 parameters");
    }

    #[test]
    fn builtin_instr_signature() {
        let sig = build_builtin_signature_info("InStr");
        assert!(sig.is_some(), "expected InStr signature");
        let sig = sig.unwrap();
        let params = sig.parameters.unwrap();
        assert_eq!(params.len(), 4, "InStr has 4 parameters");
    }

    #[test]
    fn builtin_case_insensitive() {
        assert!(build_builtin_signature_info("msgbox").is_some());
        assert!(build_builtin_signature_info("MSGBOX").is_some());
        assert!(build_builtin_signature_info("iif").is_some());
    }

    #[test]
    fn unknown_builtin_returns_none() {
        assert!(build_builtin_signature_info("NonExistentFunc").is_none());
    }
}
