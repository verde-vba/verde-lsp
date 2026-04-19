use serde::{Deserialize, Serialize};

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct WorkbookContext {
    pub workbook_path: String,
    pub sheets: Vec<SheetInfo>,
    pub workbook_named_ranges: Vec<NamedRange>,
    pub references: Vec<Reference>,
    pub last_updated: String,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct SheetInfo {
    pub name: String,
    pub code_name: String,
    pub index: u32,
    pub tables: Vec<TableInfo>,
    pub named_ranges: Vec<NamedRange>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct TableInfo {
    pub name: String,
    pub range: String,
    pub columns: Vec<String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct NamedRange {
    pub name: String,
    #[serde(default)]
    pub value: Option<String>,
    pub refers_to: Option<String>,
    #[serde(default)]
    pub range: Option<String>,
}

#[derive(Debug, Clone, Serialize, Deserialize)]
pub struct Reference {
    pub name: String,
    pub guid: String,
}

impl WorkbookContext {
    pub fn load(path: &str) -> Result<Self, Box<dyn std::error::Error>> {
        let content = std::fs::read_to_string(path)?;
        Ok(serde_json::from_str(&content)?)
    }

    pub fn sheet_names(&self) -> Vec<&str> {
        self.sheets.iter().map(|s| s.name.as_str()).collect()
    }

    pub fn table_names(&self) -> Vec<&str> {
        self.sheets
            .iter()
            .flat_map(|s| s.tables.iter().map(|t| t.name.as_str()))
            .collect()
    }

    pub fn table_columns(&self, table_name: &str) -> Vec<&str> {
        self.sheets
            .iter()
            .flat_map(|s| &s.tables)
            .find(|t| t.name == table_name)
            .map(|t| t.columns.iter().map(|c| c.as_str()).collect())
            .unwrap_or_default()
    }

    pub fn named_range_names(&self) -> Vec<&str> {
        let mut names: Vec<&str> = self
            .workbook_named_ranges
            .iter()
            .map(|n| n.name.as_str())
            .collect();
        for sheet in &self.sheets {
            for nr in &sheet.named_ranges {
                names.push(&nr.name);
            }
        }
        names
    }
}
