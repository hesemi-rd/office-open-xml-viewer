#![no_main]

use libfuzzer_sys::fuzz_target;

// Twin of fuzz_docx_parse for the pptx parser. `parse_pptx_native` is the
// non-wasm equivalent of the wasm-bindgen `parse_pptx` export and drives the
// full zip-open -> part-read -> XML-parse -> model-build pipeline (slides,
// masters, layouts, charts, theme) from raw untrusted bytes.
fuzz_target!(|data: &[u8]| {
    let _ = pptx_parser::parse_pptx_native(data);
});
