#![no_main]

use libfuzzer_sys::fuzz_target;

// Feeds arbitrary bytes straight into the docx parser's top-level entry
// point, exactly as a hostile upload would reach it. `parse_docx_native` is
// the `#[cfg(not(target_arch = "wasm32"))]` twin of the wasm-bindgen
// `parse_docx` export (used by the MCP server), so it exercises the same
// zip-open -> part-read -> XML-parse -> model-build pipeline without needing
// a wasm host. Any panic, OOM, or hang here is a real bug reachable from an
// untrusted .docx file.
fuzz_target!(|data: &[u8]| {
    let _ = docx_parser::parse_docx_native(data);
});
