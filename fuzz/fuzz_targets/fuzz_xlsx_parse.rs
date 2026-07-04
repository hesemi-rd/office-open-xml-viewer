#![no_main]

use libfuzzer_sys::fuzz_target;

// Twin of fuzz_docx_parse for the xlsx parser. `parse_workbook_native` is the
// non-wasm equivalent of the wasm-bindgen `parse_xlsx` export and drives the
// workbook-level zip-open -> part-read -> XML-parse pipeline (sheet metadata,
// styles, shared strings) from raw untrusted bytes.
fuzz_target!(|data: &[u8]| {
    let _ = xlsx_parser::parse_workbook_native(data);
});
