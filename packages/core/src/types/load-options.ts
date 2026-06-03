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
   * Override the per-entry ZIP decompression cap (bytes) used by the zip-bomb
   * guard in the Rust parser. Defaults to 512 MiB. Raise it to load documents
   * with very large embedded media, or lower it to tighten the budget for
   * untrusted input. Zero / negative values fall back to the default.
   */
  maxZipEntryBytes?: number;
}
