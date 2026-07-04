/**
 * Embedded-font registration shared by docx / pptx / xlsx viewers.
 *
 * OOXML documents may ship the fonts they use inside the package so the text
 * renders with the authored typeface even when it is absent from the host
 * system. This module turns those embedded font parts into `FontFace` objects
 * registered into the active FontFaceSet (`document.fonts` on the main thread,
 * `self.fonts` in a worker), so a subsequent `ctx.font = '… FamilyName'` selects
 * the real font instead of a substitute.
 *
 * Two obfuscation conventions exist across the formats, handled here:
 *
 * - **WordprocessingML (docx)** stores fonts as *obfuscated* parts
 *   (`.odttf`, content type `application/vnd.openxmlformats-officedocument.obfuscatedFont`).
 *   The first 32 bytes are XOR-masked with the `w:fontKey` GUID per ECMA-376
 *   §17.8.1 (Font Embedding). {@link deobfuscateOdttf} reverses it.
 * - **PresentationML (pptx)** stores fonts as `.fntdata` parts, content type
 *   `application/x-font-ttf` (raw sfnt) or `application/x-fontdata` (EOT). Per
 *   ECMA-376 §15 (the Font part), the §17.8.1 obfuscation is *only* permitted
 *   for WordprocessingML, so pptx font bytes are consumed as-is (no XOR).
 *
 * The registration mechanism (build `FontFace` from bytes, add to the set,
 * force `load()`, await `fonts.ready`) mirrors {@link preloadGoogleFonts}: the
 * two loaders share {@link activeFontSet} and the same first-paint determinism
 * contract (fonts must be ready before the caller measures/paints text).
 */
import { activeFontSet, withFontCeiling } from './preload.js';

/**
 * ECMA-376 §17.8.1 — de-obfuscate a WordprocessingML embedded font (`.odttf`).
 *
 * The algorithm (verbatim from the spec): reverse the order of the bytes in the
 * `fontKey` GUID (big-endian ordering), then XOR that 16-byte key against the
 * first 32 bytes of the font binary — once against bytes 0–15, once against
 * 16–31. The transform is its own inverse, so the same routine de-obfuscates.
 *
 * `fontKey` is the GUID string from `<w:embedRegular w:fontKey="{…}"/>` &c.,
 * e.g. `"{3EEE3167-E5B8-4798-AE48-EA6B71E31D4D}"`. Braces and hyphens are
 * optional; any 32 hex digits are accepted.
 *
 * Returns a fresh `Uint8Array` (the input is not mutated). Throws if the GUID
 * does not yield exactly 16 bytes — a malformed key must not silently produce a
 * garbage font.
 */
export function deobfuscateOdttf(bytes: Uint8Array, fontKey: string): Uint8Array {
  const key = fontKeyBytes(fontKey);
  const out = bytes.slice();
  const n = Math.min(32, out.length);
  for (let i = 0; i < n; i++) {
    // Bytes 0–15 XOR key[0..15]; bytes 16–31 XOR the SAME key again (i % 16).
    out[i] ^= key[i % 16];
  }
  return out;
}

/** Parse a `w:fontKey` GUID into its 16 bytes, reversed (big-endian) as §17.8.1
 *  requires. The reversal is over the full 16-byte string-order representation
 *  (the spec example `001B70DC-AA60-4AD5-90EC-18A0948E1EAE` →
 *  `AE1E8E94-A018-EC90-D54A-60AADC701B00`), NOT the .NET mixed-endian GUID
 *  layout. Throws on any input that is not exactly 32 hex digits. */
function fontKeyBytes(fontKey: string): Uint8Array {
  const hex = fontKey.replace(/[{}\-\s]/g, '');
  if (hex.length !== 32 || /[^0-9a-fA-F]/.test(hex)) {
    throw new Error(`invalid fontKey GUID: ${fontKey}`);
  }
  const raw = new Uint8Array(16);
  for (let i = 0; i < 16; i++) {
    raw[i] = parseInt(hex.slice(i * 2, i * 2 + 2), 16);
  }
  return raw.reverse();
}

/** One embedded font face to register: the family name it is drawn with, the
 *  raw part bytes, whether the part is `.odttf`-obfuscated (docx) and, if so,
 *  the `w:fontKey` GUID, plus the CSS bold/italic descriptors that map the
 *  style slot (regular / bold / italic / boldItalic). */
export interface EmbeddedFontFace {
  /** The FontFaceSet family name — the document's font name (e.g. "Ubuntu").
   *  The renderer sets `ctx.font` with this exact name so the browser selects
   *  the embedded face. */
  family: string;
  /** Raw bytes of the embedded font part (still obfuscated when `odttf`). */
  bytes: Uint8Array;
  /** `true` for docx `.odttf` parts (needs §17.8.1 de-obfuscation with
   *  `fontKey`); `false` for pptx `.fntdata` (raw sfnt / EOT). */
  odttf: boolean;
  /** The `w:fontKey` GUID; required when `odttf` is true, ignored otherwise. */
  fontKey?: string;
  /** CSS `font-weight` for this style slot ('normal' | 'bold'). */
  weight: 'normal' | 'bold';
  /** CSS `font-style` for this style slot ('normal' | 'italic'). */
  style: 'normal' | 'italic';
}

/**
 * Register a set of embedded font faces into the active FontFaceSet and await
 * their load, so the caller can measure/paint text with the real typefaces.
 *
 * De-obfuscation (docx `.odttf`) is applied per face. A face whose bytes are
 * corrupt, whose `fontKey` is malformed, or that exceeds {@link maxBytes} is
 * skipped with a `console.warn` — one bad embedded font must never abort the
 * whole document. Sizes are capped as a zip-bomb / memory safety net.
 *
 * No-ops (resolves immediately) when there is no FontFaceSet or `FontFace` in
 * the current context (e.g. Node without a shim), matching
 * {@link preloadGoogleFonts}.
 *
 * @param faces      the embedded faces to register
 * @param maxBytes   per-face size ceiling; faces larger than this are skipped
 *                   (default 30 MB — comfortably above a full CJK font, well
 *                   below a memory hazard)
 */
export async function registerEmbeddedFonts(
  faces: Iterable<EmbeddedFontFace>,
  maxBytes = 30 * 1024 * 1024,
): Promise<void> {
  const set = activeFontSet();
  if (!set || typeof FontFace === 'undefined') return;

  const added: FontFace[] = [];
  const failed: string[] = [];
  for (const face of faces) {
    try {
      if (face.bytes.length === 0 || face.bytes.length > maxBytes) {
        failed.push(face.family);
        continue;
      }
      const data = face.odttf
        ? deobfuscateOdttf(face.bytes, face.fontKey ?? '')
        : face.bytes;
      // Copy into a standalone ArrayBuffer: FontFace(source) reads the buffer,
      // and `data` may be a subarray view into a larger WASM/transfer buffer.
      const buf = data.buffer.slice(
        data.byteOffset,
        data.byteOffset + data.byteLength,
      ) as ArrayBuffer;
      const ff = new FontFace(face.family, buf, {
        weight: face.weight,
        style: face.style,
      });
      set.add(ff);
      added.push(ff);
    } catch {
      // Malformed fontKey / unreadable part: skip this face, keep the document.
      failed.push(face.family);
    }
  }

  if (added.length > 0) {
    await withFontCeiling(
      Promise.allSettled(added.map((f) => f.load())).then((results) => {
        results.forEach((res, i) => {
          if (res.status === 'rejected') failed.push(added[i].family);
        });
        return set.ready;
      }),
    );
  }

  if (failed.length > 0) {
    console.warn(
      `[ooxml] failed to register embedded font(s): ${[...new Set(failed)].join(', ')}; ` +
        `falling back to substitute fonts (text may shift or differ).`,
    );
  }
}
