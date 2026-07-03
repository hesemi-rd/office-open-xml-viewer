import { sniffCfb } from './cfb-sniff';
import { OoxmlError } from './ooxml-error';

/**
 * Main-thread guard shared by the docx / pptx / xlsx `load()` factories.
 *
 * Given the raw file bytes (before they are handed to the ZIP-based parser
 * worker), throw a typed {@link OoxmlError} if the bytes are a CFB (OLE2)
 * container rather than an OOXML ZIP:
 *
 *   - password-protected OOXML          -> code `'encrypted'`
 *   - legacy binary .doc / .xls / .ppt  -> code `'legacy-binary-format'`
 *   - any other / corrupt CFB           -> code `'not-ooxml'`
 *
 * A ZIP-based OOXML file (or anything that is not a CFB) passes through
 * silently. Detection is done here, on the main thread, so a real `OoxmlError`
 * instance reaches the caller — `instanceof` would not survive the worker's
 * structured-clone boundary.
 *
 * The `'encrypted'` message is intentionally provisional: decryption
 * (Agile Encryption via `LoadOptions.password`) lands in the next PR and will
 * revise this wording.
 */
export function assertNotCfbContainer(bytes: Uint8Array | ArrayBuffer): void {
  const u8 = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
  const kind = sniffCfb(u8);
  if (kind === null) return; // not a CFB — a normal ZIP-based OOXML file

  switch (kind) {
    case 'encrypted':
      throw new OoxmlError(
        'encrypted',
        'This file is password-protected (MS-OFFCRYPTO). Decryption is not yet supported.',
      );
    case 'legacy-binary-format':
      throw new OoxmlError(
        'legacy-binary-format',
        'This is a legacy binary Office file (.doc/.xls/.ppt), not OOXML.',
      );
    case 'cfb-unknown':
      throw new OoxmlError(
        'not-ooxml',
        'This file is an OLE2/Compound File container, not an OOXML (ZIP) document.',
      );
    default:
      // Exhaustiveness guard: if `CfbKind` grows a new member, this branch
      // stops compiling (the assignment requires `never`) instead of letting
      // an unrecognised CFB fall through and reach the ZIP parser worker
      // un-rejected. Fail closed at runtime too, in case a future change
      // drops the compile-time check (e.g. a widened return type upstream).
      kind satisfies never;
      throw new OoxmlError(
        'not-ooxml',
        'This file is an OLE2/Compound File container of an unrecognised kind, not an OOXML (ZIP) document.',
      );
  }
}
