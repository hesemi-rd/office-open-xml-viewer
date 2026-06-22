/**
 * Symbol / Wingdings font-encoding → Unicode normalization.
 *
 * The classic "Symbol" and "Wingdings" families do NOT use a Unicode cmap: they
 * ship a private, font-specific encoding where ordinary code points map to
 * pictographic glyphs. Crucially the two fonts use DIFFERENT encodings, so the
 * same code point is a different glyph in each (e.g. 0xB7 is BULLET "•" in
 * Symbol but a clock face in Wingdings; 0xA7 is BLACK CLUB SUIT "♣" in Symbol
 * but a small filled square "▪" in Wingdings). OOXML stores marker / run text as
 * the *font's own* code points — and Word commonly emits them in the Private Use
 * Area (U+F020–U+F0FF), so a Symbol bullet is serialized as U+F0B7 and a
 * Wingdings square as U+F0A7.
 *
 * Because the two encodings disagree on the same code points, a single shared
 * table cannot be correct for both fonts. We therefore keep TWO font-specific
 * tables ({@link SYMBOL_MAP}, {@link WINGDINGS_MAP}) and pick the right one from
 * the requested font family.
 *
 * ECMA-376 / ISO-29500 does NOT define these fonts' glyph repertoires:
 *   - §17.9.x  (numbering level text, `w:lvlText`) and §17.3.2.26 (`w:rPr`
 *     `w:rFonts`) for docx tell us WHICH font and WHICH code point, but the
 *     code-point→glyph mapping belongs to the font implementation, not the spec.
 *   - §21.1.2.3.10 (`a:sym`) for pptx is likewise just a font + char reference.
 *
 * When such a marker is drawn in a generic fallback face (the Symbol/Wingdings
 * font is frequently unavailable, especially on non-Windows hosts), the PUA code
 * point has no glyph and renders as tofu (□). We normalize the well-known
 * Symbol/Wingdings code points — both their bare form (0xA7, 0xB7, …) and their
 * PUA-shifted form (0xF0A7, 0xF0B7, …) — to the equivalent Unicode characters
 * so the intended glyph (•, ▪, →, …) renders in any fallback font.
 *
 * This is a documented font-encoding → Unicode conversion, NOT a per-sample
 * heuristic: it is keyed solely on the requested font family ("symbol" /
 * "wingdings") and each font's published code-point repertoire.
 *
 * SOURCES (authoritative, cited per entry below):
 *   - Symbol:   Adobe "Symbol Encoding to Unicode", unicode.org/Public/MAPPINGS/
 *               VENDORS/ADOBE/symbol.txt (field 1 = Unicode, field 2 = Symbol
 *               code point). This is the Adobe Symbol PostScript encoding.
 *   - Wingdings: the Microsoft Wingdings cmap as standardized in the Unicode
 *               proposals L2/11-052 and L2/11-344 (the glyph repertoire that
 *               added the Wingdings pictographs to Unicode 7.0). The byte→Unicode
 *               targets below were cross-checked against the PostScript glyph
 *               names in the shipping Wingdings.ttf (e.g. 0x21 "pencil",
 *               0xA7 "square4", 0x24 "readingglasses").
 *
 * POLICY (carried over from PR #519): a code point is mapped ONLY when its target
 * is confirmed from the source above. Glyphs whose only faithful Unicode target
 * is an astral pictograph that renders as tofu in ordinary fallback fonts (e.g.
 * the Wingdings clock faces U+1F550…, the office/hand dingbats U+1F5xx) are left
 * as passthrough: emitting the bare PUA code point is no worse, and we never
 * substitute a *wrong* but available glyph. Better tofu than a wrong glyph.
 */

/**
 * Adobe Symbol encoding → Unicode (subset).
 *
 * Values taken directly from Adobe's `symbol.txt` (Symbol code point → Unicode).
 * Only the entries that (a) appear as markers/inline symbols in practice and
 * (b) have a faithful BMP Unicode target are listed; pure math/Greek letters are
 * omitted (they already round-trip as their own Unicode characters when the
 * Symbol font is unavailable is irrelevant — Word stores those as the real Greek
 * letters, not Symbol code points).
 */
export const SYMBOL_MAP: Record<number, string> = {
  // symbol.txt: 0xA7 → U+2663 BLACK CLUB SUIT
  0xa7: '♣',
  // symbol.txt: 0xA8 → U+2666, 0xA9 → U+2665, 0xAA → U+2660 (card suits)
  0xa8: '♦', 0xa9: '♥', 0xaa: '♠',
  // symbol.txt arrows: 0xAB → U+2194, 0xAC → U+2190, 0xAD → U+2191,
  // 0xAE → U+2192, 0xAF → U+2193
  0xab: '↔', 0xac: '←', 0xad: '↑', 0xae: '→', 0xaf: '↓',
  // symbol.txt: 0xB7 → U+2022 BULLET
  0xb7: '•',
  // symbol.txt: 0xB8 → U+00F7 DIVISION SIGN (restored from PR #519 removal)
  0xb8: '÷',
  // symbol.txt: 0xB9 → U+2260 NOT EQUAL TO (restored from PR #519 removal)
  0xb9: '≠',
  // symbol.txt: 0xD7 → U+22C5 DOT OPERATOR; 0xB4 → U+00D7 MULTIPLICATION SIGN
  0xb4: '×',
  // symbol.txt: 0xB1 → U+00B1 PLUS-MINUS SIGN; 0xB0 → U+00B0 DEGREE SIGN
  0xb0: '°', 0xb1: '±',
  // symbol.txt: 0xA3 → U+2264 LESS-THAN OR EQUAL TO; 0xB3 → U+2265 GREATER-THAN OR EQUAL TO
  0xa3: '≤', 0xb3: '≥',
};

/**
 * Microsoft Wingdings cmap → Unicode (subset).
 *
 * Each byte's glyph is identified from the shipping Wingdings.ttf PostScript
 * glyph names (e.g. 0xA7 "square4", 0x21 "pencil"); the Unicode 7.0 proposals
 * L2/11-052 / L2/11-344 catalogue those glyphs but at DEDICATED ASTRAL code
 * points (barb arrows U+1F868…, bold check/X U+1F5F6…) that tofu in ordinary
 * fonts. So the targets below are deliberate faithful BMP substitutions (→, ✓,
 * ▪, …), NOT those canonical astral code points. Entries whose only faithful
 * target is an astral pictograph (the office/hand/clock dingbats U+1F5xx /
 * U+1F55x) are OMITTED — passthrough, since the bare PUA point is no worse.
 */
export const WINGDINGS_MAP: Record<number, string> = {
  // 0x24 "readingglasses" → U+1F453 EYEGLASSES. Astral with no BMP equivalent;
  // unlike the omitted clock/hand dingbats we DO map it because U+1F453 is the
  // semantically exact glyph — it renders in an emoji-capable fallback and is
  // never a *wrong* glyph (passthrough would always tofu the bare PUA point).
  0x24: '\u{1F453}',
  // 0x4A "smileface" → U+263A; 0x4B "neutralface" → U+1F610 (astral, exact —
  // same rationale as 0x24); 0x4C "frownface" → U+2639.
  0x4a: '☺', 0x4b: '\u{1F610}', 0x4c: '☹',
  // 0x76 "xrhombus" → U+2756 BLACK DIAMOND MINUS WHITE X (restored from PR #519)
  0x76: '❖',
  // 0xA7 "square4" → U+25AA BLACK SMALL SQUARE (the Wingdings list bullet)
  0xa7: '▪',
  // 0x6C "circle6" → U+25CF BLACK CIRCLE; 0x6E "square6" → U+25A0 BLACK SQUARE.
  // 0x74 "lozenge6" and 0x77 "rhombus4" are both SOLID/black diamond-family
  // glyphs (a tall lozenge and a wider rhombus); BMP doesn't distinguish the
  // proportion, so both approximate to U+25C6 BLACK DIAMOND. (0x77 was wrongly
  // U+25C7 WHITE DIAMOND — a fill error — before this fix.)
  0x6c: '●', 0x6e: '■', 0x74: '◆', 0x77: '◆',
  // 0xFB "xmarkbld" → U+2717 BALLOT X; 0xFC "checkbld" → U+2713 CHECK MARK;
  // 0xFD "boxxmarkbld" → U+2612 BALLOT BOX WITH X;
  // 0xFE "boxcheckbld" → U+2611 BALLOT BOX WITH CHECK (restored from PR #519)
  0xfb: '✗', 0xfc: '✓', 0xfd: '☒', 0xfe: '☑',
  // Wingdings barb2 arrow block — 0xDF "barb2left" … 0xE6 "barb2se". Mapped to
  // Unicode arrows so they render in any font when Wingdings is unavailable.
  0xdf: '←', 0xe0: '→', 0xe1: '↑', 0xe2: '↓',
  0xe3: '↖', 0xe4: '↗', 0xe5: '↙', 0xe6: '↘',
};

/**
 * Backwards-compatible alias. Historically a single `SYMBOL_FONT_MAP` was
 * exported; it leaned toward the Wingdings repertoire. New code should select
 * {@link SYMBOL_MAP} / {@link WINGDINGS_MAP} explicitly via
 * {@link symbolFontToUnicode}.
 *
 * @deprecated Use {@link symbolFontToUnicode} (it picks the right font table).
 */
export const SYMBOL_FONT_MAP: Record<number, string> = WINGDINGS_MAP;

/** Add the PUA-shifted (0xF000 + code) variant for every entry in `base`. */
function withPua(base: Record<number, string>): Record<number, string> {
  const out: Record<number, string> = {};
  for (const key of Object.keys(base)) {
    const code = Number(key);
    out[code] = base[code];
    out[0xf000 + code] = base[code];
  }
  return out;
}

const SYMBOL_LOOKUP = withPua(SYMBOL_MAP);
const WINGDINGS_LOOKUP = withPua(WINGDINGS_MAP);

/**
 * Normalize a single Symbol/Wingdings code point to its Unicode equivalent.
 *
 * The lookup table is chosen by `fontFamily`: an exact "symbol" selects the
 * Adobe Symbol encoding ({@link SYMBOL_MAP}); any family containing "wingdings"
 * selects the Wingdings cmap ({@link WINGDINGS_MAP}). Both the bare code point
 * (0xA7, 0xB7, …) and its PUA-shifted form (0xF0A7, 0xF0B7, …) resolve.
 *
 * Returns `char` unchanged when `fontFamily` is null/undefined, is not a symbol
 * font, or has no mapping in the selected table (passthrough — better the bare
 * code point than a wrong glyph).
 *
 * @param char       the marker/run character as stored in the OOXML (often a
 *                   PUA code point such as U+F0B7)
 * @param fontFamily the requested font family (ascii / hAnsi axis)
 */
export function symbolFontToUnicode(
  char: string,
  fontFamily: string | null | undefined,
): string {
  if (!fontFamily) return char;
  const lower = fontFamily.toLowerCase();
  const table = lower.includes('wingdings')
    ? WINGDINGS_LOOKUP
    : lower === 'symbol'
      ? SYMBOL_LOOKUP
      : null;
  if (!table) return char;
  const code = char.charCodeAt(0);
  return table[code] ?? char;
}

/**
 * True when `fontFamily` is one of the font-specific (non-Unicode-cmap)
 * encodings that {@link symbolFontToUnicode} knows how to normalize — currently
 * exactly "Symbol" and any "Wingdings" family (§17.3.2.26 / §21.1.2.3.10).
 *
 * Intended as the shared "is this a symbol font?" gate so callers don't
 * hand-roll divergent regexes. docx uses it; pptx still has its own broader
 * inline gate (`/wingding|webding|symbol/i`) and is a follow-up to migrate. It
 * deliberately does NOT match "Webdings" (core has no Webdings table yet — a
 * Webdings PUA point would pass the gate only to return unchanged) nor a Latin
 * face like "SymbolMT" (the "symbol" check is exact, not a substring), so a
 * caller wanting either must add it explicitly and accept passthrough.
 */
export function isSymbolFontFamily(
  fontFamily: string | null | undefined,
): boolean {
  if (!fontFamily) return false;
  const lower = fontFamily.toLowerCase();
  return lower === 'symbol' || lower.includes('wingdings');
}

/** A maximal run of characters sharing a single rendering disposition. */
export interface SymbolTextSegment {
  /** The (possibly normalized) text for this run. */
  text: string;
  /**
   * True when at least one character in `text` was mapped from the symbol
   * font's private encoding to a Unicode equivalent. Such a segment must be
   * drawn in a generic fallback face (the Symbol/Wingdings font, if present,
   * would re-interpret the Unicode code point as the WRONG glyph); an unmapped
   * segment keeps the requested symbol font so an installed Symbol/Wingdings
   * still draws its native glyph.
   */
  mapped: boolean;
}

/**
 * String-level companion to {@link symbolFontToUnicode}: normalize every
 * character of `text` through the font's private encoding and split the result
 * into maximal same-disposition runs so the caller can switch the draw font at
 * each mapped/unmapped boundary.
 *
 * When `fontFamily` is not a symbol font (per {@link isSymbolFontFamily}), the
 * whole string is returned as a single `{ text, mapped: false }` segment
 * unchanged — the caller then needs no special handling. Iterates by code point
 * so astral targets (e.g. Wingdings 0x24 → U+1F453) survive intact.
 */
export function symbolTextToUnicodeSegments(
  text: string,
  fontFamily: string | null | undefined,
): SymbolTextSegment[] {
  if (!isSymbolFontFamily(fontFamily) || text.length === 0) {
    return [{ text, mapped: false }];
  }
  const out: SymbolTextSegment[] = [];
  let buf = '';
  let bufMapped: boolean | null = null;
  for (const ch of text) {
    const normalized = symbolFontToUnicode(ch, fontFamily);
    const mapped = normalized !== ch;
    if (bufMapped === null || mapped === bufMapped) {
      bufMapped = mapped;
      buf += normalized;
    } else {
      out.push({ text: buf, mapped: bufMapped });
      bufMapped = mapped;
      buf = normalized;
    }
  }
  if (buf.length > 0) out.push({ text: buf, mapped: bufMapped ?? false });
  return out;
}
