/**
 * Symbol / Wingdings font-encoding → Unicode normalization.
 *
 * The classic "Symbol" (§) and "Wingdings" families do NOT use a Unicode cmap:
 * they ship a private, font-specific encoding where ordinary code points map to
 * pictographic glyphs (e.g. in Symbol, code point 0xB7 is BULLET "•", not MIDDLE
 * DOT; in Wingdings, 0xA7 is a small filled square). OOXML stores marker / run
 * text as the *font's own* code points — and Word commonly emits them in the
 * Private Use Area (U+F020–U+F0FF), so a Symbol bullet is serialized as U+F0B7
 * and a Wingdings square as U+F0A7.
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
 * "wingdings") and the font's published code-point repertoire.
 *
 * NOTE: This table duplicates `WINGDINGS_MAP` / `applySymbolFont` in
 * `packages/pptx/src/renderer.ts`. The pptx renderer should later be refactored
 * to import from here; that change is left for the pptx session because the
 * branch that introduced this module may not edit `packages/pptx/**`.
 */
// Only entries verified against authoritative sources (unicode.org
// ADOBE/symbol.txt for Symbol; the Wingdings cmap for Wingdings) are listed.
// The load-bearing markers are the bullets (Symbol 0xB7 "•", Wingdings 0xA7
// "▪") and the Wingdings barb arrow block; the pencil/scissors/face dingbats
// below are correct under Wingdings.
//
// LIMITATION: this is a single table shared by both fonts, so a code point that
// means different glyphs in Symbol vs Wingdings (e.g. 0xB8: Symbol "÷",
// Wingdings a clock face) cannot be disambiguated here. Such ambiguous /
// font-specific code points are deliberately OMITTED rather than guessed — they
// fall through unchanged (better tofu than a wrong glyph). Splitting into
// font-specific SYMBOL_MAP / WINGDINGS_MAP is tracked as a follow-up.
export const SYMBOL_FONT_MAP: Record<number, string> = {
  0x21: '✏', 0x22: '✂', 0x23: '✁',
  0x4A: '☺', 0x4C: '☹',
  0xFC: '✓', 0xFB: '✗',
  0xA7: '▪', 0xB7: '•',
  0xF0A7: '▪', 0xF0B7: '•',
  // Wingdings barb2 arrow block — glyph names verified against the Wingdings
  // cmap (0xDF=barb2left … 0xE6=barb2se). Mapped to Unicode arrows so they
  // render in any font even when Wingdings is unavailable.
  0xDF: '←', 0xE0: '→', 0xE1: '↑', 0xE2: '↓',
  0xE3: '↖', 0xE4: '↗', 0xE5: '↙', 0xE6: '↘',
  0xF0DF: '←', 0xF0E0: '→', 0xF0E1: '↑', 0xF0E2: '↓',
  0xF0E3: '↖', 0xF0E4: '↗', 0xF0E5: '↙', 0xF0E6: '↘',
};

/**
 * Normalize a single Symbol/Wingdings code point to its Unicode equivalent.
 *
 * Returns `char` unchanged when `fontFamily` is null/undefined, is not a symbol
 * font ("symbol" exactly, or any family containing "wingdings"), or has no
 * mapping in {@link SYMBOL_FONT_MAP}.
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
  if (lower.includes('wingdings') || lower === 'symbol') {
    const code = char.charCodeAt(0);
    return SYMBOL_FONT_MAP[code] ?? char;
  }
  return char;
}
