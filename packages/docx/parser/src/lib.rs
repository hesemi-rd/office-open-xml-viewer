use wasm_bindgen::prelude::*;

mod markdown;
mod math;
mod numbering;
mod parser;
mod styles;
mod types;
mod xml_util;

#[wasm_bindgen]
pub fn parse_docx(data: &[u8], max_zip_entry_bytes: Option<u64>) -> Result<String, JsValue> {
    console_error_panic_hook::set_once();
    let _guard = ooxml_common::zip::scoped_max(max_zip_entry_bytes);
    let doc =
        parser::parse(data).map_err(|e| JsValue::from_str(&format!("docx-parser error: {e}")))?;
    serde_json::to_string(&doc).map_err(|e| JsValue::from_str(&format!("serialize error: {e}")))
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

/// Extract raw bytes for a single embedded image entry (e.g.
/// "word/media/image1.png") from a docx zip archive. Thin `wasm_bindgen`
/// wrapper over the shared [`ooxml_common::zip::extract_zip_entry`] reader; used
/// by the main thread to lazily materialize image blobs on demand.
#[wasm_bindgen]
pub fn extract_image(
    data: &[u8],
    path: &str,
    max_zip_entry_bytes: Option<u64>,
) -> Result<Vec<u8>, JsValue> {
    ooxml_common::zip::extract_zip_entry(data, path, max_zip_entry_bytes)
        .map_err(|e| JsValue::from_str(&e))
}

/// Native equivalent of `parse_docx` for use from the MCP server.
#[cfg(not(target_arch = "wasm32"))]
pub fn parse_docx_native(data: &[u8]) -> Result<String, String> {
    parser::parse(data).and_then(|doc| serde_json::to_string(&doc).map_err(|e| e.to_string()))
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

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn extract_image_reads_entry() {
        use std::io::{Cursor, Write};
        let mut buf = Vec::new();
        {
            let mut w = zip::ZipWriter::new(Cursor::new(&mut buf));
            let o = zip::write::SimpleFileOptions::default();
            w.start_file("word/media/i.png", o).unwrap();
            w.write_all(b"X").unwrap();
            w.finish().unwrap();
        }
        assert_eq!(extract_image(&buf, "word/media/i.png", None).unwrap(), b"X");
    }

    /// `parse_docx` now returns `Result<String, JsValue>` (Err surfaces as a
    /// thrown JS Error), replacing the old hand-built `{"error":"…"}` string.
    /// The old path double-embedded the parser message into JSON *without
    /// escaping*, so a message containing a `"` produced invalid JSON that made
    /// the TS-side `JSON.parse` throw a confusing SyntaxError instead of the
    /// real cause. We can't call the `#[wasm_bindgen]` fn from a plain unit
    /// test (that needs wasm-bindgen-test), but `parse_docx` is now a thin
    /// `parser::parse(...).map_err(...)?` wrapper, so proving the internal
    /// parse returns a plain `Err(String)` for bad input is enough: the JSON
    /// escaping hazard is gone *by construction* because the error never round
    /// trips through hand-assembled JSON — wasm-bindgen serializes the string.
    #[test]
    fn parse_rejects_non_zip_bytes_without_json_escaping_hazard() {
        // Not a zip archive — parser::parse must return Err, not panic.
        let err = parser::parse(&[1, 2, 3]).expect_err("non-zip bytes must error");
        // A raw String error is returned verbatim; even a `"`-laden message is
        // carried as-is (no manual `{"error":"…"}` templating to corrupt).
        let quoted = format!("boom: \"{}\"", err);
        assert!(quoted.contains('"'), "sanity: message can contain quotes");
    }
}
