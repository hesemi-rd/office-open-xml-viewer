use crate::{parse_cell_ref, resolve_implicit_ordinal, resolve_sheet_path};
use crate::{CellRange, SharedString, SheetMeta, XlsxZip};
use ooxml_common::depth::parse_guarded;
use ooxml_common::ns::is_x_ns;
use ooxml_common::zip::read_zip_string;
use std::collections::HashMap;

pub(crate) const MAX_REFERENCE_CELLS: usize = 1_000_000;
const MAX_COL: u32 = 16_384;
const MAX_ROW: u32 = 1_048_576;

#[derive(Clone, Debug, PartialEq)]
pub(crate) enum ReferencedCellValue {
    Empty,
    Text(String),
    Number(f64),
}

pub(crate) fn split_sheet_ref(formula: &str) -> Option<(Option<String>, String)> {
    let formula = formula.trim();
    if formula.is_empty() {
        return None;
    }

    let (sheet_name, reference) = match formula.rfind('!') {
        Some(bang) => {
            let raw_sheet = formula[..bang].trim();
            let reference = formula[bang + 1..].trim();
            if raw_sheet.is_empty() || reference.is_empty() {
                return None;
            }
            let sheet = if raw_sheet.starts_with('\'') && raw_sheet.ends_with('\'') {
                if raw_sheet.len() < 2 {
                    return None;
                }
                raw_sheet[1..raw_sheet.len() - 1].replace("''", "'")
            } else {
                if raw_sheet.contains(['\'', '!', '[', ']']) {
                    return None;
                }
                raw_sheet.to_string()
            };
            (Some(sheet), reference)
        }
        None => (None, formula),
    };

    Some((sheet_name, reference.replace('$', "")))
}

fn parse_a1_cell(reference: &str) -> Option<(u32, u32)> {
    let split = reference
        .find(|c: char| !c.is_ascii_alphabetic())
        .unwrap_or(reference.len());
    let (col_text, row_text) = reference.split_at(split);
    if col_text.is_empty() || row_text.is_empty() || !row_text.chars().all(|c| c.is_ascii_digit()) {
        return None;
    }

    let col = col_text.chars().try_fold(0u32, |value, c| {
        value
            .checked_mul(26)?
            .checked_add(c.to_ascii_uppercase() as u32 - 'A' as u32 + 1)
    })?;
    let row = row_text.parse::<u32>().ok()?;
    if !(1..=MAX_COL).contains(&col) || !(1..=MAX_ROW).contains(&row) {
        return None;
    }
    Some((col, row))
}

pub(crate) fn parse_a1_range(reference: &str) -> Option<CellRange> {
    let mut parts = reference.trim().split(':');
    let first = parts.next()?;
    let second = parts.next();
    if parts.next().is_some() {
        return None;
    }
    let (col_a, row_a) = parse_a1_cell(first)?;
    let (col_b, row_b) = match second {
        Some(cell) => parse_a1_cell(cell)?,
        None => (col_a, row_a),
    };
    Some(CellRange {
        top: row_a.min(row_b),
        left: col_a.min(col_b),
        bottom: row_a.max(row_b),
        right: col_a.max(col_b),
    })
}

fn inline_string_text(cell: roxmltree::Node<'_, '_>) -> String {
    cell.descendants()
        .filter(|node| {
            node.is_element()
                && node.tag_name().name() == "t"
                && is_x_ns(node.tag_name().namespace())
        })
        .filter_map(|node| node.text())
        .collect()
}

fn cell_value(
    cell: roxmltree::Node<'_, '_>,
    shared_strings: &[SharedString],
) -> ReferencedCellValue {
    let cell_type = cell.attribute("t").unwrap_or("");
    if cell_type == "inlineStr" {
        return ReferencedCellValue::Text(inline_string_text(cell));
    }
    let value = cell
        .children()
        .find(|node| {
            node.is_element()
                && node.tag_name().name() == "v"
                && is_x_ns(node.tag_name().namespace())
        })
        .and_then(|node| node.text())
        .unwrap_or("");

    match cell_type {
        "s" => value
            .parse::<usize>()
            .ok()
            .and_then(|index| shared_strings.get(index))
            .map(|string| ReferencedCellValue::Text(string.text.clone()))
            .unwrap_or(ReferencedCellValue::Empty),
        "str" => ReferencedCellValue::Text(value.to_string()),
        "" | "n" => value
            .parse::<f64>()
            .ok()
            .filter(|number| number.is_finite())
            .map(ReferencedCellValue::Number)
            .unwrap_or(ReferencedCellValue::Empty),
        _ => ReferencedCellValue::Empty,
    }
}

pub(crate) fn extract_reference_cells(
    sheet_xml: &str,
    range: &CellRange,
    shared_strings: &[SharedString],
) -> Vec<ReferencedCellValue> {
    let Some(row_count) = range
        .bottom
        .checked_sub(range.top)
        .and_then(|value| value.checked_add(1))
    else {
        return Vec::new();
    };
    let Some(col_count) = range
        .right
        .checked_sub(range.left)
        .and_then(|value| value.checked_add(1))
    else {
        return Vec::new();
    };
    let Some(total) = (row_count as usize).checked_mul(col_count as usize) else {
        return Vec::new();
    };
    if total > MAX_REFERENCE_CELLS {
        return Vec::new();
    }

    let Ok(document) = parse_guarded(sheet_xml) else {
        return Vec::new();
    };
    let mut values = vec![ReferencedCellValue::Empty; total];
    let mut previous_row = 0;
    for row_node in document.descendants().filter(|node| {
        node.is_element() && node.tag_name().name() == "row" && is_x_ns(node.tag_name().namespace())
    }) {
        let row = resolve_implicit_ordinal(
            row_node
                .attribute("r")
                .and_then(|value| value.parse::<u32>().ok()),
            &mut previous_row,
        );
        let mut previous_col = 0;
        for cell in row_node.children().filter(|node| {
            node.is_element()
                && node.tag_name().name() == "c"
                && is_x_ns(node.tag_name().namespace())
        }) {
            let (col, cell_row) = match cell.attribute("r") {
                Some(reference) => {
                    let (col, row) = parse_cell_ref(reference);
                    previous_col = col;
                    (col, row)
                }
                None => (resolve_implicit_ordinal(None, &mut previous_col), row),
            };
            if cell_row < range.top
                || cell_row > range.bottom
                || col < range.left
                || col > range.right
            {
                continue;
            }
            let index =
                (cell_row - range.top) as usize * col_count as usize + (col - range.left) as usize;
            values[index] = cell_value(cell, shared_strings);
        }
    }
    values
}

#[allow(clippy::too_many_arguments)]
pub(crate) fn resolve_worksheet_reference(
    archive: &mut XlsxZip,
    formula: &str,
    current_sheet_xml: &str,
    current_sheet_name: &str,
    sheets: &[SheetMeta],
    workbook_rels: &roxmltree::Document<'_>,
    shared_strings: &[SharedString],
    xml_cache: &mut HashMap<String, Option<String>>,
) -> Vec<ReferencedCellValue> {
    let Some((source_sheet, reference)) = split_sheet_ref(formula) else {
        return Vec::new();
    };
    let Some(range) = parse_a1_range(&reference) else {
        return Vec::new();
    };
    let sheet_name = source_sheet.as_deref().unwrap_or(current_sheet_name);
    if sheet_name == current_sheet_name {
        return extract_reference_cells(current_sheet_xml, &range, shared_strings);
    }

    if !xml_cache.contains_key(sheet_name) {
        let xml = sheets
            .iter()
            .find(|sheet| sheet.name == sheet_name)
            .and_then(|sheet| resolve_sheet_path(workbook_rels, &sheet.r_id))
            .and_then(|path| read_zip_string(archive, &format!("xl/{path}")).ok());
        xml_cache.insert(sheet_name.to_string(), xml);
    }
    xml_cache
        .get(sheet_name)
        .and_then(Option::as_deref)
        .map(|xml| extract_reference_cells(xml, &range, shared_strings))
        .unwrap_or_default()
}

#[cfg(test)]
mod tests {
    use super::*;
    use crate::{CellRange, SharedString};

    #[test]
    fn quoted_unicode_sheet_reference_is_split_and_unescaped() {
        assert_eq!(
            split_sheet_ref("'التقرير'!$A$2:$A$5"),
            Some((Some("التقرير".into()), "A2:A5".into())),
        );
        assert_eq!(
            split_sheet_ref("'Bob''s data'!C1"),
            Some((Some("Bob's data".into()), "C1".into())),
        );
        assert_eq!(split_sheet_ref("A1:A3"), Some((None, "A1:A3".into())));
    }

    #[test]
    fn direct_a1_range_is_normalized() {
        let range = parse_a1_range("C5:A2").expect("direct A1 range parses");
        assert_eq!(
            (range.top, range.left, range.bottom, range.right),
            (2, 1, 5, 3),
        );
        assert!(parse_a1_range("Sales").is_none());
        assert!(parse_a1_range("A1+1").is_none());
    }

    #[test]
    fn worksheet_cells_resolve_inline_shared_formula_string_and_numbers() {
        let xml = r#"<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><sheetData><row r="1"><c r="A1" t="inlineStr"><is><t>Inline</t></is></c><c r="B1" t="s"><v>0</v></c><c r="C1" t="str"><f>UPPER(&quot;x&quot;)</f><v>X</v></c><c r="D1"><v>42</v></c></row></sheetData></worksheet>"#;
        let shared = vec![SharedString {
            text: "Shared".into(),
            ..Default::default()
        }];
        let range = CellRange {
            top: 1,
            left: 1,
            bottom: 1,
            right: 4,
        };
        assert_eq!(
            extract_reference_cells(xml, &range, &shared),
            vec![
                ReferencedCellValue::Text("Inline".into()),
                ReferencedCellValue::Text("Shared".into()),
                ReferencedCellValue::Text("X".into()),
                ReferencedCellValue::Number(42.0),
            ],
        );
    }

    #[test]
    fn oversized_range_is_rejected_before_allocation() {
        let range = CellRange {
            top: 1,
            left: 1,
            bottom: 1_048_576,
            right: 16_384,
        };
        assert!(extract_reference_cells("<worksheet/>", &range, &[]).is_empty());
    }
}
