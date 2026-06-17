//! Shared configuration for ZIP entry decompression limits.
//!
//! OOXML parsers must cap per-entry decompressed output to block zip-bomb DoS.
//! The cap defaults to 512 MiB — large enough for legitimate embedded video /
//! 4K media but small enough to refuse pathological archives — and can be
//! overridden per-parse via the `wasm_bindgen` entry points so library users
//! can tighten the budget (untrusted gateways) or loosen it (legitimate huge
//! decks) without forking.
//!
//! The current cap is stored in a thread-local so existing `read_zip_*`
//! helpers in each parser can consult it without threading a parameter
//! through ~70 call sites. Each WASM entry point installs a [`Guard`] for
//! its scope and the cap is restored on drop, so concurrent JS callers never
//! interfere (WASM is single-threaded; each invocation runs to completion).

use std::cell::Cell;

/// 512 MiB. OOXML legitimately reaches tens of MB (embedded video, 4K
/// images) but not hundreds, so this cap blocks zip-bomb DoS without
/// rejecting real files.
pub const DEFAULT_MAX_ZIP_ENTRY_BYTES: u64 = 512 * 1024 * 1024;

thread_local! {
    static MAX_ZIP_ENTRY_BYTES: Cell<u64> = const { Cell::new(DEFAULT_MAX_ZIP_ENTRY_BYTES) };
}

/// RAII guard that restores the previous cap when dropped. Created by
/// [`scoped_max`]; the caller should bind it to a `let _guard = …` for the
/// full duration of the parse call.
#[must_use = "binding the guard keeps the cap installed for this scope"]
pub struct Guard {
    previous: u64,
}

impl Drop for Guard {
    fn drop(&mut self) {
        MAX_ZIP_ENTRY_BYTES.with(|c| c.set(self.previous));
    }
}

/// Install a per-call ZIP entry size cap for the lifetime of the returned
/// guard. `None`, zero, or any non-positive value falls back to
/// [`DEFAULT_MAX_ZIP_ENTRY_BYTES`].
pub fn scoped_max(value: Option<u64>) -> Guard {
    let resolved = value
        .filter(|&n| n > 0)
        .unwrap_or(DEFAULT_MAX_ZIP_ENTRY_BYTES);
    let previous = MAX_ZIP_ENTRY_BYTES.with(|c| c.replace(resolved));
    Guard { previous }
}

/// Current cap in effect on this thread. Parsers consult this from their
/// `read_zip_*` helpers when validating entry sizes.
pub fn current_max() -> u64 {
    MAX_ZIP_ENTRY_BYTES.with(Cell::get)
}

/// Read one zip entry's bytes by path. Honors the scoped max-entry guard.
pub fn extract_zip_entry(
    data: &[u8],
    path: &str,
    max_zip_entry_bytes: Option<u64>,
) -> Result<Vec<u8>, String> {
    use std::io::{Cursor, Read};
    let _guard = scoped_max(max_zip_entry_bytes);
    let cursor = Cursor::new(data);
    let mut zip = zip::ZipArchive::new(cursor).map_err(|e| format!("zip open error: {e}"))?;
    let mut entry = zip
        .by_name(path)
        .map_err(|e| format!("entry not found: {path}: {e}"))?;
    let max = max_zip_entry_bytes.unwrap_or(u64::MAX);
    let mut buf = Vec::with_capacity(entry.size().min(max) as usize);
    entry
        .by_ref()
        .take(max)
        .read_to_end(&mut buf)
        .map_err(|e| format!("read error: {e}"))?;
    Ok(buf)
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn extract_zip_entry_reads_by_path() {
        use std::io::{Cursor, Write};
        let mut buf = Vec::new();
        {
            let mut w = zip::ZipWriter::new(Cursor::new(&mut buf));
            let opts = zip::write::SimpleFileOptions::default();
            w.start_file("ppt/media/image1.png", opts).unwrap();
            w.write_all(b"\x89PNGdata").unwrap();
            w.finish().unwrap();
        }
        let bytes = extract_zip_entry(&buf, "ppt/media/image1.png", None).unwrap();
        assert_eq!(bytes, b"\x89PNGdata");
        assert!(extract_zip_entry(&buf, "ppt/media/missing.png", None)
            .unwrap_err()
            .contains("not found"));
    }
}
