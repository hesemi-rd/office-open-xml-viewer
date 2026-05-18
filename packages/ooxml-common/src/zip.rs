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
