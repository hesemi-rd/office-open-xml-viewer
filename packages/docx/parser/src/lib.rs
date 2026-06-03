use wasm_bindgen::prelude::*;

mod types;
mod xml_util;
mod styles;
mod numbering;
mod parser;
mod markdown;
mod math;

#[wasm_bindgen]
pub fn parse_docx(data: &[u8], max_zip_entry_bytes: Option<u64>) -> String {
    console_error_panic_hook::set_once();
    let _guard = ooxml_common::zip::scoped_max(max_zip_entry_bytes);
    match parser::parse(data) {
        Ok(doc) => serde_json::to_string(&doc).unwrap_or_else(|e| {
            format!("{{\"error\":\"{}\"}}", e)
        }),
        Err(e) => format!("{{\"error\":\"{}\"}}", e),
    }
}

/// WASM-callable markdown projection (mirrors `to_markdown_native`). Returns
/// GitHub-flavoured markdown of headings / paragraphs / tables / footnotes,
/// discarding positioning, section properties, fonts, and drawing shapes.
#[wasm_bindgen]
pub fn docx_to_markdown(data: &[u8], max_zip_entry_bytes: Option<u64>) -> Result<String, JsValue> {
    console_error_panic_hook::set_once();
    let _guard = ooxml_common::zip::scoped_max(max_zip_entry_bytes);
    let doc = parser::parse(data).map_err(|e| JsValue::from_str(&e))?;
    Ok(markdown::render_document(&doc))
}

/// Native equivalent of `parse_docx` for use from the MCP server.
#[cfg(not(target_arch = "wasm32"))]
pub fn parse_docx_native(data: &[u8]) -> Result<String, String> {
    parser::parse(data)
        .and_then(|doc| serde_json::to_string(&doc).map_err(|e| e.to_string()))
}

/// Parse a docx and project the result to GitHub-flavoured markdown:
/// headings (from outlineLevel), paragraphs with bullet/numbered lists,
/// tables, footnote references collated at the end, and rich-text
/// formatting (bold / italic / strikethrough / hyperlink). Designed for AI
/// agents that need to read content efficiently — discards positioning,
/// section properties, font metrics, drawing shapes.
#[cfg(not(target_arch = "wasm32"))]
pub fn to_markdown_native(data: &[u8]) -> Result<String, String> {
    let doc = parser::parse(data)?;
    Ok(markdown::render_document(&doc))
}
