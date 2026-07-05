// UAX #50 — Unicode Vertical Text Layout (https://www.unicode.org/reports/tr50/).
// The single source of truth for how a code point orients when set in a vertical
// line (tbRl / eaVert), consumed by every renderer's vertical-text draw path.
//
// The `Vertical_Orientation` (vo) property has four values:
//   • U  — Upright: same orientation as the code charts. CJK ideographs, kana,
//          Hangul, fullwidth forms, and the already-vertical presentation forms
//          (U+FE10–U+FE19). Drawn standing up.
//   • R  — Rotated 90° clockwise. Latin letters, Western digits, and most Latin
//          punctuation. This is the property's file-wide default (see below).
//   • Tu — Transformed typographically, fallback Upright. A glyph the font may
//          substitute with a dedicated vertical form; if it does not, the glyph
//          is drawn UPRIGHT. Small kana, the ideographic comma/full stop 、。,
//          fullwidth ！？：；，．, etc.
//   • Tr — Transformed typographically, fallback Rotated. Like Tu, but the
//          fallback (no vertical glyph available) is to ROTATE. Corner brackets
//          「」, parentheses （）, angle brackets 〈〉, the katakana-hiragana
//          prolonged sound mark ー (U+30FC), quotation marks, etc.
//
// The generated table is built straight from the UCD `VerticalOrientation.txt`
// data section plus its `@missing: 0000..10FFFF; R` default, so code points not
// listed in the file resolve to R exactly as UAX #50 specifies. See
// packages/core/scripts/gen-vertical-orientation.mjs for provenance.

import {
  VO_NAMES,
  VO_RANGE_STARTS,
  VO_RANGE_VALUE,
} from './vertical-orientation.generated.js';

export { UNICODE_VERSION as VO_UNICODE_VERSION } from './vertical-orientation.generated.js';

/** UAX #50 Vertical_Orientation property value. */
export type VerticalOrientation = 'U' | 'R' | 'Tu' | 'Tr';

/**
 * Vertical_Orientation index (into {@link VO_NAMES}) for a code point.
 * Binary search for the greatest range start ≤ cp; ranges are gap-free and cover
 * [0, 0x110000), so a match always exists for a valid Unicode scalar value.
 */
function verticalOrientationIndex(cp: number): number {
  let lo = 0;
  let hi = VO_RANGE_STARTS.length - 1;
  while (lo < hi) {
    const mid = (lo + hi + 1) >> 1;
    if (VO_RANGE_STARTS[mid] <= cp) lo = mid;
    else hi = mid - 1;
  }
  return VO_RANGE_VALUE[lo];
}

/**
 * The UAX #50 `Vertical_Orientation` of a code point — one of `'U'`, `'R'`,
 * `'Tu'`, `'Tr'`. Backed by the generated UCD table; no heuristics.
 *
 * @param cp A Unicode scalar value (e.g. from `String.prototype.codePointAt`).
 */
export function verticalOrientation(cp: number): VerticalOrientation {
  return VO_NAMES[verticalOrientationIndex(cp)] as VerticalOrientation;
}

/**
 * DRAW-TIME vertical-form substitution for vo=Tu punctuation: the U+FE10–U+FE19
 * "Vertical Forms" presentation glyph to paint instead of `cp`, or `null` when
 * there is none to apply.
 *
 * UAX #50 §5 ("Glyph Changes for Vertical Orientation") describes the Tu/Tr
 * transform as substituting a vertical glyph variant. A Canvas cannot invoke the
 * font's `vert`/`vrt2` OpenType feature, so for the Tu punctuation that has a
 * dedicated Unicode presentation form we substitute the code point itself and
 * let the font supply the pre-positioned glyph (upper-right cell corner).
 *
 * NOTE: this map is a rendering mapping (fullwidth/CJK source → vertical form),
 * NOT the strict inverse of UnicodeData's `<vertical>` compatibility
 * decompositions. Per UnicodeData.txt, FE10/FE13–FE16 decompose to the ASCII
 * comma / colon / semicolon / exclamation / question mark
 * (U+002C/003A/003B/0021/003F) — but vertical Japanese text carries the
 * FULLWIDTH forms, so the map keys on those. Only FE11 (← 3001 、) and
 * FE12 (← 3002 。) are exact inverses.
 *
 * Every key is vo=Tu, i.e. actually reaches the renderers' upright/substitute
 * draw branch. Deliberately NOT in the map:
 *   • ：FF1A / ；FF1B and 〖3016 / 〗3017 are vo=Tr — the Tr draw branch rotates
 *     them and never consults this map. Making Tr substitute-first (FE13/FE14/
 *     FE17/FE18 here, plus the U+FE35+ forms for fullwidth parens/brackets —
 *     UAX#50's Tr means "substitute a vertical glyph, rotate only as fallback")
 *     is tracked as follow-up in issues #790 / #771.
 *   • … U+2026 is vo=R — the sideways branch rotates it 90° with the page, so
 *     its three dots already stack vertically, visually equivalent to Word's
 *     vertical ellipsis; no substitution is needed.
 *   • Small kana (vo=Tu) have no U+FExx presentation form — a true vertical
 *     variant needs the font's `vert` feature — so they draw upright unchanged.
 *
 * Renderers apply this ONLY at glyph-draw time (glyph selection): the text model,
 * advance/width (kept at 1 em), selection, and find/highlight all continue to use
 * the ORIGINAL code point, so searching for 。 still matches a substituted 。.
 *
 * @param cp A Unicode scalar value.
 * @returns The vertical presentation-form code point, or null.
 */
export function verticalFormSubstitute(cp: number): number | null {
  return VERTICAL_FORM_MAP.get(cp) ?? null;
}

// vo=Tu punctuation → U+FE10–U+FE19 vertical presentation form. See the
// verticalFormSubstitute doc above for why the map keys on the fullwidth forms
// and why Tr/R punctuation (：；〖〗…) is excluded. Names per the Unicode
// "Vertical Forms" chart (U+FE10–U+FE1F).
const VERTICAL_FORM_MAP: ReadonlyMap<number, number> = new Map<number, number>([
  [0xff0c, 0xfe10], // ， fullwidth comma       → PRESENTATION FORM FOR VERTICAL COMMA
  [0x3001, 0xfe11], // 、 ideographic comma     → … FOR VERTICAL IDEOGRAPHIC COMMA
  [0x3002, 0xfe12], // 。 ideographic full stop → … FOR VERTICAL IDEOGRAPHIC FULL STOP
  [0xff01, 0xfe15], // ！ fullwidth exclamation → … FOR VERTICAL EXCLAMATION MARK
  [0xff1f, 0xfe16], // ？ fullwidth question    → … FOR VERTICAL QUESTION MARK
]);

/**
 * DRAW-TIME vertical-form substitution for vo=Tr BRACKETS: the U+FE35–U+FE44
 * "Presentation Forms For Vertical" glyph for a fullwidth paren / corner
 * bracket / angle bracket / brace / lenticular bracket, or `null` when the code
 * point is not one of them (or is a Tr code point with no vertical form).
 *
 * WHY this is separate from {@link verticalFormSubstitute} (the Tu map): the two
 * fallbacks differ. A Tu code point with no vertical form draws UPRIGHT; a Tr
 * one ROTATES. Keeping the maps distinct lets each draw branch pick the right
 * fallback (see the docx renderer's `drawVerticalRun`) without conflating the
 * two UAX #50 transform classes.
 *
 * WHY substitute a Tr bracket at all: UAX #50 §5 defines the Tr transform as
 * "substitute a vertical glyph variant; ROTATE only as the fallback when none is
 * available." A Canvas cannot reach the font's `vert`/`vrt2` OpenType feature via
 * `fillText`, but the Unicode "Presentation Forms For Vertical" block supplies a
 * dedicated, already-vertical code point for every fullwidth bracket, and those
 * code points ARE reachable. Drawing the vertical form UPRIGHT (rather than
 * rotating the horizontal bracket) is both closer to Word (which uses the same
 * `vert` glyphs) and — critically — measurable: an upright glyph's along-column
 * position is governed by its VERTICAL ink extent (`actualBoundingBoxAscent/
 * Descent`, which a Canvas exposes), whereas a ROTATED bracket's along-column
 * position is governed by its HORIZONTAL ink offset inside the advance box, which
 * `measureText` does NOT expose (it reports the advance box, not the tight ink
 * box). So substitution is what makes metric-driven cell-centring possible.
 *
 * Deliberately NOT here (they are vo=Tr but have no U+FE3x vertical form, so the
 * renderer keeps the rotate fallback):
 *   • ー U+30FC prolonged sound mark
 *   • “ ” U+201C/201D double quotation marks
 *
 * Mapping source: the UnicodeData.txt `<vertical>` compatibility decompositions
 * of the U+FE30 block (each vertical form decomposes to its horizontal bracket).
 *
 * Renderers apply this ONLY at glyph-draw time (glyph selection): the advance/
 * width, text model, selection, and find/highlight all keep the ORIGINAL code
 * point, so searching for （ still matches a substituted （.
 *
 * @param cp A Unicode scalar value.
 * @returns The U+FE3x vertical presentation-form code point, or null.
 */
export function verticalBracketFormSubstitute(cp: number): number | null {
  return VERTICAL_BRACKET_FORM_MAP.get(cp) ?? null;
}

// vo=Tr fullwidth brackets → U+FE35–U+FE44 "Presentation Forms For Vertical".
// Keyed on the horizontal bracket; values from the UnicodeData `<vertical>`
// decompositions. ー (30FC) and the double quotes (201C/201D) are vo=Tr with no
// vertical form and are absent, so the renderer rotates them.
const VERTICAL_BRACKET_FORM_MAP: ReadonlyMap<number, number> = new Map<number, number>([
  [0xff08, 0xfe35], // （ fullwidth left parenthesis  → ︵
  [0xff09, 0xfe36], // ） fullwidth right parenthesis → ︶
  [0xff5b, 0xfe37], // ｛ fullwidth left curly brace  → ︷
  [0xff5d, 0xfe38], // ｝ fullwidth right curly brace → ︸
  [0x3014, 0xfe39], // 〔 left tortoise-shell bracket → ︹
  [0x3015, 0xfe3a], // 〕 right tortoise-shell bracket→ ︺
  [0x3010, 0xfe3b], // 【 left black lenticular       → ︻
  [0x3011, 0xfe3c], // 】 right black lenticular      → ︼
  [0x300a, 0xfe3d], // 《 left double angle bracket    → ︽
  [0x300b, 0xfe3e], // 》 right double angle bracket   → ︾
  [0x3008, 0xfe3f], // 〈 left angle bracket           → ︿
  [0x3009, 0xfe40], // 〉 right angle bracket          → ﹀
  [0x300c, 0xfe41], // 「 left corner bracket          → ﹁
  [0x300d, 0xfe42], // 」 right corner bracket         → ﹂
  [0x300e, 0xfe43], // 『 left white corner bracket    → ﹃
  [0x300f, 0xfe44], // 』 right white corner bracket   → ﹄
]);
