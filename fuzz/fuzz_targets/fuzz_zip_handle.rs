#![no_main]

use libfuzzer_sys::fuzz_target;

// Exercises the shared zip-part-open path used by every parser
// (`ooxml_common::zip::extract_zip_entry`): open a raw byte blob as a zip
// archive and read one entry by name. This covers the zip-bomb size guard
// (`scoped_max` / `current_max`) and the central-directory parsing that
// docx/pptx/xlsx all funnel through, without going through a full document
// parse. `data` is split into a path string (first line, up to the first
// newline) and the remaining bytes are the candidate zip archive, so the
// fuzzer can co-evolve plausible part names (e.g. "word/document.xml")
// alongside the archive bytes instead of only ever missing by name.
fuzz_target!(|data: &[u8]| {
    let split_at = data.iter().position(|&b| b == b'\n').unwrap_or(0);
    let (path_bytes, rest) = data.split_at(split_at);
    let zip_bytes = if rest.is_empty() { rest } else { &rest[1..] };
    let path = String::from_utf8_lossy(path_bytes);
    let _ = ooxml_common::zip::extract_zip_entry(zip_bytes, &path, None);
    // Also fuzz the "already declared a tighter cap" path, which is the one
    // real callers actually hit (max_zip_entry_bytes override from JS).
    let _ = ooxml_common::zip::extract_zip_entry(zip_bytes, &path, Some(4096));
});
