use wasm_bindgen::prelude::*;

mod types;
mod xml_util;
mod styles;
mod numbering;
mod parser;

#[cfg(not(target_arch = "wasm32"))]
mod markdown;

#[wasm_bindgen]
pub fn parse_docx(data: &[u8]) -> String {
    console_error_panic_hook::set_once();
    match parser::parse(data) {
        Ok(doc) => serde_json::to_string(&doc).unwrap_or_else(|e| {
            format!("{{\"error\":\"{}\"}}", e)
        }),
        Err(e) => format!("{{\"error\":\"{}\"}}", e),
    }
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
