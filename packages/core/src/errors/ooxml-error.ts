/**
 * Machine-readable code for a typed load-time failure.
 *
 * Kept intentionally small for now — the three container-level failures the
 * `load()` factories can detect on the main thread before handing bytes to the
 * parser worker (see `sniffCfb`). This is the seed of the broader typed-error
 * surface tracked as PD4 (OoxmlError typed errors); future codes such as
 * `'invalid-password'` (once Agile Encryption decryption lands, PD8) or
 * `'corrupt'` will extend this union. Add codes here rather than throwing bare
 * `Error(string)`, so callers can `switch` on `err.code` instead of matching
 * message text.
 */
export type OoxmlErrorCode = 'encrypted' | 'legacy-binary-format' | 'not-ooxml';

/**
 * Typed error thrown by the docx / pptx / xlsx `load()` factories for failures
 * that carry a stable, programmatic {@link OoxmlErrorCode} (e.g. a
 * password-protected or legacy-binary file detected from its container magic).
 *
 * Note on workers: `instanceof OoxmlError` does not survive a structured-clone
 * across the worker boundary. Detection that needs a typed error is therefore
 * done on the main thread (before the worker is involved) so a genuine
 * `OoxmlError` instance is thrown to the caller. Errors that must cross the
 * worker boundary should carry the `code` string and be reconstructed on the
 * main side.
 */
export class OoxmlError extends Error {
  readonly code: OoxmlErrorCode;

  constructor(code: OoxmlErrorCode, message: string) {
    super(message);
    this.name = 'OoxmlError';
    this.code = code;
    // Restore the prototype chain for environments that down-level `extends
    // Error` (e.g. older TS `target`), so `instanceof OoxmlError` holds.
    Object.setPrototypeOf(this, OoxmlError.prototype);
  }
}
