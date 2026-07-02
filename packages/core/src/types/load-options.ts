import type { MathRenderer } from '../math/mathjax';

/**
 * Common load-time options shared by the docx / pptx / xlsx
 * `Document.load` / `Presentation.load` / `Workbook.load` factories and their
 * viewer wrappers.
 *
 * This is the single source of truth — each package re-exports this exact type
 * as its `LoadOptions` so application code can pass one options object to any
 * of the three.
 */
export interface LoadOptions {
  /**
   * Opt in to loading webfont substitutes from Google Fonts
   * (`fonts.googleapis.com`). Default `false` — the canvas falls back to
   * locally available fonts.
   *
   * When enabled, end-user IP / User-Agent is sent to Google, which may
   * have privacy / GDPR implications for your application. To avoid the
   * third-party request, host the substitutes yourself and reference them
   * via `@font-face` in your application CSS.
   */
  useGoogleFonts?: boolean;
  /**
   * Override the URL the parser worker fetches the WebAssembly module from.
   *
   * By default each format resolves the `.wasm` asset that ships next to its
   * bundle (relative to the module URL), so no configuration is needed. Set
   * this to serve the parser WASM from a CDN or a self-hosted path instead — a
   * relative value is resolved against the current document URL. The same
   * dependency-injection contract across docx / pptx / xlsx.
   *
   * The referenced file must be the matching format's `*_parser_bg.wasm`
   * artifact (the one wasm-bindgen emitted for that parser); pointing it at a
   * mismatched or missing file makes `load()` reject when the worker
   * instantiates it.
   */
  wasmUrl?: string | URL;
  /**
   * Override the per-entry ZIP decompression cap (bytes) used by the zip-bomb
   * guard in the Rust parser. Defaults to 512 MiB. Raise it to load documents
   * with very large embedded media, or lower it to tighten the budget for
   * untrusted input. Zero / negative values fall back to the default.
   */
  maxZipEntryBytes?: number;
  /**
   * Reject the parse request if the parser worker does not answer within this
   * many milliseconds. Opt-in safety net for a wedged or crashed worker that
   * would otherwise leave `load()` pending forever. **Default: unlimited** —
   * parsing a large document with heavy embedded media can legitimately take
   * tens of seconds, so no timeout is imposed unless you set one. A worker that
   * throws or fails to load already rejects immediately regardless of this
   * value; this bound only covers the "silent, never-responds" case.
   */
  workerTimeoutMs?: number;
  /**
   * Opt-in OMML equation engine (MathJax + STIX Two Math, ~3 MB). Inject it
   * **once** here and every render of this document / presentation / workbook
   * uses it — the same dependency-injection contract across all three formats
   * and their viewers. Import it from the separate `@silurus/ooxml/math` entry
   * (`import { math } from '@silurus/ooxml/math'`). Omit it and equations are
   * skipped and the engine tree-shakes away entirely (no network, no bundle
   * cost).
   */
  math?: MathRenderer;
}
