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
 * font's `vert`/`vrt2` OpenType feature, so for the CORNER-HANGING Tu punctuation
 * that has a dedicated Unicode presentation form we substitute the code point
 * itself and let the font supply the pre-positioned glyph (upper-right cell
 * corner). The caller then draws it em-box-centred (NO ink re-centring), so the
 * font's designed corner offset is preserved — that is what places 、。 in the
 * cell's upper-right the way Word does (PDF-verified on sample-26: 、 ink at
 * −0.32em along-column, +0.33em cross-axis).
 *
 * NOTE: this map is a rendering mapping (fullwidth/CJK source → vertical form),
 * NOT the strict inverse of UnicodeData's `<vertical>` compatibility
 * decompositions. Per UnicodeData.txt, FE10 decomposes to the ASCII comma
 * (U+002C) — but vertical Japanese text carries the FULLWIDTH comma, so the map
 * keys on that. Only FE11 (← 3001 、) and FE12 (← 3002 。) are exact inverses.
 *
 * Every key is vo=Tu, i.e. actually reaches the renderers' upright/substitute
 * draw branch. Deliberately NOT in the map:
 *   • ！ FF01 (→ FE15) and ？ FF1F (→ FE16) — although vo=Tu, the exclamation /
 *     question marks stand UPRIGHT and HORIZONTALLY CENTRED in a vertical line,
 *     unlike the corner-hanging comma / full stop. The original fullwidth ！／？
 *     drawn upright is already the correct vertical shape (the marks are
 *     vertically symmetric) and lands centred on the column in every font. The
 *     U+FE15/FE16 "vertical form" glyphs are corner-designed in many fonts
 *     (Hiragino ink at +0.31em cross-axis), so substituting them pushed ！／？ to
 *     the right of the column — the sample-26 "！ shifted right" defect (#771).
 *     PDF ground truth: Word centres ！ (+0.03em), matching the upright original.
 *   • ：FF1A / ；FF1B and 〖3016 / 〗3017 are vo=Tr, not Tu, so this Tu-only map
 *     stays null for them. The white lenticular brackets 〖〗 (→ FE17/FE18) are
 *     substituted via {@link verticalBracketFormSubstitute}; the fullwidth colon /
 *     semicolon take a GEOMETRIC fallback in the renderers (their FE13/FE14 forms
 *     are absent from most render fonts — see {@link verticalTrUprightFallback}),
 *     issue #969, following #790 / #771.
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
//
// ONLY the CORNER-HANGING punctuation (comma / full stop) is substituted. In
// vertical Japanese typography these hang in the UPPER-RIGHT of the cell (JIS X
// 4051 §4.3 kutōten placement), which is exactly how a font designs its FE10–FE12
// glyph — ink pushed to the corner of the em box. Substituting the code point and
// drawing it em-box-centred reproduces that corner placement directly.
//
// Deliberately NOT substituted: ！ FF01 (→ FE15) and ？ FF1F (→ FE16) — see the
// verticalFormSubstitute doc above. They stand upright and centred, so the FE15/
// FE16 corner-designed forms would shift them off-column; the original ！／？ is
// drawn upright instead.
const VERTICAL_FORM_MAP: ReadonlyMap<number, number> = new Map<number, number>([
  [0xff0c, 0xfe10], // ， fullwidth comma       → PRESENTATION FORM FOR VERTICAL COMMA
  [0x3001, 0xfe11], // 、 ideographic comma     → … FOR VERTICAL IDEOGRAPHIC COMMA
  [0x3002, 0xfe12], // 。 ideographic full stop → … FOR VERTICAL IDEOGRAPHIC FULL STOP
]);

/**
 * DRAW-TIME vertical-form substitution for a vo=Tr code point: the U+FE1x/FE3x
 * vertical presentation glyph to paint instead of `cp`, or `null` when the code
 * point is not one that has a form (or is a Tr code point with no vertical form).
 *
 * Covers, keyed by their horizontal form:
 *   • the fullwidth parens / corner / angle / brace / tortoise-shell / black-
 *     lenticular brackets → U+FE35–U+FE44 ("Presentation Forms For Vertical");
 *   • the WHITE lenticular brackets 〖〗 → U+FE17/FE18 (the U+FE1x "Vertical Forms"
 *     block, issue #969), which the common substitute fonts DO contain.
 *
 * NOTE: the fullwidth colon ： / semicolon ； (→ FE13/FE14) were REMOVED from this
 * map (issue #969 follow-up). Those forms are absent from most render fonts and a
 * Canvas cannot invoke the font's `vert` feature, so substituting reached a
 * mispositioned system-fallback glyph. They now take a GEOMETRIC fallback in the
 * renderers (colon rotate, semicolon upright per {@link verticalTrUprightFallback}).
 *
 * WHY this is separate from {@link verticalFormSubstitute} (the Tu map): the two
 * fallbacks differ. A Tu code point with no vertical form draws UPRIGHT; a Tr
 * one ROTATES. Keeping the maps distinct lets each draw branch pick the right
 * fallback (see the docx renderer's `drawVerticalRun`) without conflating the
 * two UAX #50 transform classes. (The name says "Bracket" for its original scope;
 * it is really "the vo=Tr vertical-form map" — brackets plus the white lenticular
 * 〖〗. The fullwidth colon/semicolon were removed — see the NOTE above.)
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
 * Deliberately NOT here (they are vo=Tr but have no Unicode vertical presentation
 * form, so the renderer keeps the rotate fallback):
 *   • ー U+30FC prolonged sound mark
 *   • “ ” U+201C/201D double quotation marks
 *
 * Mapping source: the UnicodeData.txt `<vertical>` compatibility decompositions
 * of the U+FE10 and U+FE30 blocks (each vertical form decomposes to its horizontal
 * source). FE17/FE18 decompose exactly to 〖/〗.
 *
 * Renderers apply this ONLY at glyph-draw time (glyph selection): the advance/
 * width, text model, selection, and find/highlight all keep the ORIGINAL code
 * point, so searching for （ or ： still matches a substituted （ / ：.
 *
 * @param cp A Unicode scalar value.
 * @returns The U+FE1x/FE3x vertical presentation-form code point, or null.
 */
export function verticalBracketFormSubstitute(cp: number): number | null {
  return VERTICAL_BRACKET_FORM_MAP.get(cp) ?? null;
}

// vo=Tr code points → their U+FE1x/FE3x vertical presentation form (drawn upright,
// rotate only as the fallback for the entries NOT listed here). Values are the
// UnicodeData `<vertical>` decompositions read backwards.
//
// Two groups share this map because they share the vo=Tr substitute-first / rotate-
// fallback BEHAVIOUR — which is what the renderer keys on — even though the second
// group is punctuation, not brackets:
//   1. Fullwidth brackets / parens / braces → U+FE35–FE44 ("Presentation Forms For
//      Vertical" block). Keyed on the horizontal bracket.
//   2. The white lenticular brackets 〖〗 → the U+FE17/FE18 forms in the U+FE1x
//      "Vertical Forms" block (issue #969), which the common substitute fonts DO
//      contain, so they stay substituted upright (PDF-adjudicated). The fullwidth
//      colon / semicolon ：；(→ FE13/FE14) were REMOVED — those forms are absent
//      from most render fonts, so they take a geometric fallback in the renderers
//      instead (colon rotate → FE13's side-by-side dots, semicolon upright → FE14's
//      dot-over-comma; see verticalTrUprightFallback). FE17/FE18 decompose to
//      the white lenticular brackets 3016/3017 exactly.
//
// ー (30FC) and the double quotes (201C/201D) are vo=Tr with NO vertical form and are
// absent, so the renderer rotates them. The fullwidth colon ：(FF1A) rotates too
// (→ FE13's side-by-side dots); the semicolon ；(FF1B) is upright (→ FE14) — see
// verticalTrUprightFallback.
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
  // vo=Tr white lenticular brackets in the U+FE1x block (issue #969). The
  // fullwidth colon ： / semicolon ； (→ FE13/FE14) were REMOVED here (issue #969
  // follow-up): their U+FE13/FE14 forms are absent from most render fonts
  // (e.g. Hiragino Mincho ProN, macOS's Yu Mincho substitute) and a Canvas
  // cannot reach the font's `vert` OpenType feature, so substituting the code
  // point reaches the system fallback cascade and paints a DIFFERENT font's glyph
  // positioned wrong (measured: FE13/FE14 ink lands ~0.25em RIGHT of the column
  // centre in both skia AND Chrome). They now take a GEOMETRIC fallback that
  // reproduces each vertical form's design directly — colon rotates (→ FE13's
  // side-by-side dots), semicolon stays upright (→ FE14's dot-over-comma). See
  // {@link verticalTrUprightFallback}. The white lenticular brackets FE17/FE18
  // ARE present in the substitute font, so they stay substituted.
  [0x3016, 0xfe17], // 〖 left white lenticular        → ︗ vertical left white lenticular
  [0x3017, 0xfe18], // 〗 right white lenticular       → ︘ vertical right white lenticular
]);

// The vo=Tr code points whose no-vertical-form rotate fallback must be OVERRIDDEN
// to UPRIGHT. This is an APPLICATION-LEVEL layout override, not the normative UAX
// #50 Tr fallback (which is rotation): Vertical_Orientation is an informative
// property that a layout application may override, and Word does here. The fullwidth
// semicolon ；(FF1B) is the sole member: its Unicode vertical form FE14 (︔) is a
// dot-over-comma drawn UPRIGHT, NOT a rotation of the base — Word/JIS X 4051 verified
// (sample-47 PDF: ；renders as a vertical dot+comma centred on the column, whereas a
// 90° rotation would put the comma and dot side by side). The colon ：(FF1A) is NOT
// here: FE13 (︓) IS a 90° rotation of the base (two dots that go from vertically
// stacked to side by side), so the generic Tr rotate fallback reproduces it. ASCII
// `;` (003B) is vo=R (Latin), so only the fullwidth form participates.
const VERTICAL_TR_UPRIGHT_FALLBACK: ReadonlySet<number> = new Set<number>([
  0xff1b, // ； fullwidth semicolon → FE14 design is upright dot-over-comma
]);

/**
 * True when a vo=Tr code point with NO substituted vertical form should take an
 * UPRIGHT fallback rather than the generic ROTATE fallback. Only the fullwidth
 * semicolon ；(FF1B) qualifies (its FE14 vertical form is upright dot-over-comma,
 * not a rotation). All other rotate-fallback Tr code points — the colon ：, ー,
 * quotes, un-substituted brackets — return false and rotate. This is an
 * application-level override of the informative UAX #50 property (Word behaviour),
 * not the normative Tr fallback (rotation).
 *
 * @param cp A Unicode scalar value.
 */
export function verticalTrUprightFallback(cp: number): boolean {
  return VERTICAL_TR_UPRIGHT_FALLBACK.has(cp);
}

// The three Tr long marks initially wired to the DOM `vert` capability probe by
// #1023. This is a routing class only: it makes no claim about the inaccessible
// vertical design and must never select a geometric mirror/shear approximation.
const VERTICAL_TR_LONG_MARKS: ReadonlySet<number> = new Set<number>([
  0x30fc, // ー katakana-hiragana prolonged sound mark
  0x301c, // 〜 wave dash
  0xff5e, // ～ fullwidth tilde
]);

/**
 * Whether `cp` is one of the three Tr long marks whose reachable `vert` glyph was
 * enabled by #1023. If `vert` is unreachable, UAX #50's plain rotated fallback is
 * used without fabricating the glyph's unknown design.
 */
export function verticalTrLongMark(cp: number): boolean {
  return VERTICAL_TR_LONG_MARKS.has(cp);
}
