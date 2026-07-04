#![no_main]

use libfuzzer_sys::fuzz_target;

// Exercises the shared `.rels` relationship XML parser
// (`ooxml_common::rels::parse_rels`) used by docx/pptx/xlsx to resolve
// `r:id` / `r:embed` references. Also runs `resolve_target` over every
// parsed relationship's raw `Target` against a couple of representative
// base directories, covering the `../` path-normalization logic that had a
// real bug in docx's old private copy (see ooxml-common/src/rels.rs docs).
// Malformed/non-UTF-8 input must degrade to an empty map, never panic.
fuzz_target!(|data: &[u8]| {
    let Ok(xml) = std::str::from_utf8(data) else {
        return;
    };
    let rels = ooxml_common::rels::parse_rels(xml);
    for rel in rels.values() {
        let _ = rel.resolve("word/charts");
        let _ = rel.resolve("ppt/slides");
        let _ = rel.resolve("xl/worksheets");
        let _ = rel.resolve("");
    }
});
