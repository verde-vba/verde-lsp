use tower_lsp::lsp_types::*;

use crate::analysis::resolve::offset_to_position;
use crate::analysis::AnalysisHost;

pub fn folding_ranges(host: &AnalysisHost, uri: &Url) -> Vec<FoldingRange> {
    host.with_source(uri, |symbols, source| {
        symbols
            .proc_ranges
            .iter()
            .map(|(_, range)| {
                let start = offset_to_position(source, range.start as usize);
                let end = offset_to_position(source, range.end as usize);
                FoldingRange {
                    start_line: start.line,
                    start_character: None,
                    end_line: end.line,
                    end_character: None,
                    kind: Some(FoldingRangeKind::Region),
                    collapsed_text: None,
                }
            })
            .collect()
    })
    .unwrap_or_default()
}
