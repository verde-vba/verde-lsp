use serde::{Deserialize, Serialize};
use smol_str::SmolStr;

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ExcelObjectType {
    pub name: SmolStr,
    pub properties: Vec<PropertyDef>,
    pub methods: Vec<MethodDef>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct PropertyDef {
    pub name: SmolStr,
    pub return_type: SmolStr,
    pub readonly: bool,
    pub description: String,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct MethodDef {
    pub name: SmolStr,
    pub return_type: Option<SmolStr>,
    pub params: Vec<ParamDef>,
    pub description: String,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct ParamDef {
    pub name: SmolStr,
    pub type_name: SmolStr,
    pub optional: bool,
}

pub fn load_builtin_types() -> Vec<ExcelObjectType> {
    // MVP: hardcoded core types. Will load from JSON in excel-types/ later.
    vec![
        ExcelObjectType {
            name: SmolStr::new("Range"),
            properties: vec![
                prop("Value", "Variant", false),
                prop("Text", "String", true),
                prop("Row", "Long", true),
                prop("Column", "Long", true),
                prop("Address", "String", true),
                prop("Count", "Long", true),
                prop("Cells", "Range", true),
                prop("Rows", "Range", true),
                prop("Columns", "Range", true),
                prop("Font", "Font", false),
                prop("Interior", "Interior", false),
                prop("NumberFormat", "Variant", false),
                prop("Formula", "Variant", false),
                prop("FormulaR1C1", "Variant", false),
            ],
            methods: vec![
                method("Select", None, &[]),
                method("Copy", None, &[("Destination", "Range", true)]),
                method("Clear", None, &[]),
                method("ClearContents", None, &[]),
                method("Delete", None, &[("Shift", "Variant", true)]),
                method("Find", Some("Range"), &[("What", "Variant", false)]),
                method("Sort", None, &[("Key1", "Range", true)]),
                method("AutoFill", None, &[("Destination", "Range", false)]),
            ],
        },
        ExcelObjectType {
            name: SmolStr::new("Worksheet"),
            properties: vec![
                prop("Name", "String", false),
                prop("CodeName", "String", true),
                prop("Index", "Long", true),
                prop("Cells", "Range", true),
                prop("Range", "Range", true),
                prop("Rows", "Range", true),
                prop("Columns", "Range", true),
                prop("UsedRange", "Range", true),
                prop("Visible", "XlSheetVisibility", false),
            ],
            methods: vec![
                method("Activate", None, &[]),
                method("Select", None, &[]),
                method("Copy", None, &[]),
                method("Delete", None, &[]),
                method("Protect", None, &[("Password", "String", true)]),
            ],
        },
        ExcelObjectType {
            name: SmolStr::new("Workbook"),
            properties: vec![
                prop("Name", "String", true),
                prop("Path", "String", true),
                prop("FullName", "String", true),
                prop("Sheets", "Sheets", true),
                prop("Worksheets", "Sheets", true),
                prop("ActiveSheet", "Object", true),
                prop("Saved", "Boolean", false),
            ],
            methods: vec![
                method("Save", None, &[]),
                method("SaveAs", None, &[("Filename", "String", false)]),
                method("Close", None, &[("SaveChanges", "Boolean", true)]),
                method("Activate", None, &[]),
            ],
        },
        ExcelObjectType {
            name: SmolStr::new("Application"),
            properties: vec![
                prop("ActiveWorkbook", "Workbook", true),
                prop("ActiveSheet", "Object", true),
                prop("ActiveCell", "Range", true),
                prop("Selection", "Object", true),
                prop("ThisWorkbook", "Workbook", true),
                prop("Workbooks", "Workbooks", true),
                prop("ScreenUpdating", "Boolean", false),
                prop("DisplayAlerts", "Boolean", false),
                prop("EnableEvents", "Boolean", false),
                prop("StatusBar", "Variant", false),
                prop("Calculation", "XlCalculation", false),
            ],
            methods: vec![
                method("Calculate", None, &[]),
                method("Run", Some("Variant"), &[("Macro", "String", false)]),
                method("InputBox", Some("Variant"), &[("Prompt", "String", false)]),
            ],
        },
    ]
}

/// Return the names of all Excel `Application` members (properties + methods).
///
/// These identifiers (e.g. `ActiveWorkbook`, `ActiveSheet`, `Range`, `Cells`,
/// `Worksheets`, `Selection`) are exposed as implicit globals in VBA's Excel
/// host, so they must be treated as "declared" under Option Explicit.
///
/// Returns an empty vector if the `Application` type is not present in the
/// builtin type set (defensive; should not happen in practice).
pub fn application_globals() -> Vec<String> {
    let types = load_builtin_types();
    let Some(app) = types.iter().find(|t| t.name == "Application") else {
        return Vec::new();
    };
    let mut names = Vec::with_capacity(app.properties.len() + app.methods.len());
    for p in &app.properties {
        names.push(p.name.to_string());
    }
    for m in &app.methods {
        names.push(m.name.to_string());
    }
    names
}

fn prop(name: &str, type_name: &str, readonly: bool) -> PropertyDef {
    PropertyDef {
        name: SmolStr::new(name),
        return_type: SmolStr::new(type_name),
        readonly,
        description: String::new(),
    }
}

fn method(name: &str, ret: Option<&str>, params: &[(&str, &str, bool)]) -> MethodDef {
    MethodDef {
        name: SmolStr::new(name),
        return_type: ret.map(SmolStr::new),
        params: params
            .iter()
            .map(|(n, t, opt)| ParamDef {
                name: SmolStr::new(n),
                type_name: SmolStr::new(t),
                optional: *opt,
            })
            .collect(),
        description: String::new(),
    }
}
