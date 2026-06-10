// Font line-height metrics for fonts whose real vertical metrics differ
// substantially from whatever the browser substitutes when the font is not
// installed.
//
// Why this exists
// ---------------
// For `lineRule="auto"` / single spacing (ECMA-376 §17.3.1.33), Word's line
// height is `multiplier × singleLineHeight`, where the single-line height is
// the font's design line height. On Windows (the platform Word's PDF export
// targets) that single line height is `(usWinAscent + usWinDescent) /
// unitsPerEm` per the OS/2 table.
//
// Our renderer normally derives the single-line height from the Canvas
// `fontBoundingBoxAscent + fontBoundingBoxDescent` of the font the browser
// actually used. That is correct when the document's font is installed, but
// when it is substituted the fallback's metrics can differ from the intended
// font's — in EITHER direction — by enough to break line spacing, vertical
// centering, and pagination:
//
//   * Fallback TALLER than the document font: the substitute's design line box
//     overstates Word's line height. The lines spread out, and in a fixed-size
//     box (e.g. a `w:trHeight hRule="exact"` table row with `w:vAlign="center"`,
//     §17.4.81 / §17.4.84) the over-measured content height pushes the centered
//     text above the true center. `Sakkal Majalla` substituted by `Noto Naskh
//     Arabic` is the canonical example (see below).
//   * Fallback SHORTER than the document font: the substitute understates Word's
//     line height, so lines overlap and vertical drift accumulates down the
//     page. `Meiryo` / `Meiryo UI` is the canonical example.
//
// Both are the same defect: the line box must follow the DOCUMENT font's design
// line height, not the substitute's. On Windows (Word's PDF-export target) that
// design line height is `(usWinAscent + usWinDescent) / unitsPerEm` from the
// OS/2 table whenever the font does not set the USE_TYPO_METRICS fsSelection
// bit (bit 7); both fonts below clear that bit, so the win metric is what Word
// uses.
//
// Examples:
//   * `Meiryo` / `Meiryo UI` — unitsPerEm 2048, usWinAscent 2210 +
//     usWinDescent 1059 ≈ 3269 → 1.596 em. The macOS/Chromium fallback
//     (`Hiragino Sans`) reports only ≈1.0 em from `fontBoundingBox`. A 48 pt
//     Meiryo title with `line="168"` (0.7×) renders at 0.7 × 1.60 × 48 ≈
//     53.8 pt in Word but collapses to ≈34 pt with the fallback metric.
//   * `Sakkal Majalla` — unitsPerEm 2048, usWinAscent 1810 + usWinDescent
//     1050 = 2860 → 1.3965 em (extracted directly from the installed
//     `Sakkal Majalla Regular.ttf`, version 5.01; USE_TYPO_METRICS bit clear,
//     so Word uses the win metric). The Google-Fonts substitute `Noto Naskh
//     Arabic` reports a far larger ≈2.2 em line box (winAscent ≈1.535 em +
//     winDescent ≈0.665 em) from `fontBoundingBox`. In sample-7's page-2 table
//     header (12 pt, `w:trHeight` 350 twips hRule="exact" → 17.5 pt row,
//     `w:vAlign="center"`) the substitute's 2.2 × 12 = 26.4 pt content box
//     centers to (17.5 − 26.4)/2 = −4.4 pt, pushing the white header text above
//     the band top; Word fits the 1.3965 × 12 ≈ 16.8 pt box inside the row with
//     visible top/bottom gaps.
//
// This table provides the intended font's win ascent and descent ratios
// (`usWinAscent / unitsPerEm` and `usWinDescent / unitsPerEm`) so the line box
// can be sized — and the baseline placed within it — as Word would, independent
// of which fallback ends up drawing the glyphs. Carrying ascent and descent
// separately (rather than only their sum) matters when the substitute's
// ascent:descent split differs from the document font's: using the substitute's
// split would mis-place the baseline inside an otherwise correctly sized box,
// shifting vertically-centered text off center. Only fonts whose metrics are
// verified from real OS/2 data belong here — never a value tuned to make one
// sample look right. Latin fonts are intentionally absent: their win ratio
// (~1.15–1.22 em) is close to what the browser fallback already reports, so the
// correction is negligible.

interface WinMetric {
  /** `usWinAscent / unitsPerEm`. */ asc: number;
  /** `usWinDescent / unitsPerEm`. */ desc: number;
}

/** A known font's win metrics. Keyed by a normalized (lowercased) name test. */
const WIN_METRICS: ReadonlyArray<readonly [test: (n: string) => boolean, m: WinMetric]> = [
  // Meiryo / Meiryo UI — unitsPerEm 2048, usWinAscent 2210, usWinDescent 1059
  // (OS/2 table; USE_TYPO_METRICS fsSelection bit 7 is clear, so Word uses the
  // win metric for `lineRule="auto"` single spacing, §17.3.1.33). The single-
  // line ratio is therefore the raw win sum 3269/2048 = 1.5962 em, carried as
  // the true ascent/descent ratios so the baseline split stays Meiryo's own.
  //
  // Provenance / cross-check: sample-3's single-spaced body uses Meiryo UI.
  // Measured against the Word-export reference PNG, the 11 pt intro paragraph
  // renders at 17.5 px ≈ 1.591 em/pt — i.e. exactly the 1.5962 win ratio. (An
  // earlier revision pinned this to a round 1.60 from a 48 pt-title eyeball;
  // that title is Latin/Arial Nova, not Meiryo, so the pin had no Meiryo basis.
  // Per the package CLAUDE.md "spec over empirical constant" rule it is replaced
  // here by the documented OS/2 win sum.) The 9 pt body in the same document
  // measures ~15.0 px = 1.667 em/pt in Word, ABOVE 1.5962: that residual is
  // Word's per-line device-pixel rounding at small CJK sizes (cumulative line
  // tops snapped to whole pixels), not a different single-line ratio — it is not
  // reproducible from any OS/2 value and is intentionally NOT encoded as a
  // constant here.
  [
    (n) => n.includes('meiryo') || n.includes('メイリオ'),
    { asc: 2210 / 2048, desc: 1059 / 2048 },
  ],
  // Sakkal Majalla — unitsPerEm 2048, usWinAscent 1810, usWinDescent 1050
  // → asc 0.8838, desc 0.5127, sum = 1.3965. Extracted directly from the
  // installed `Sakkal Majalla Regular.ttf` (version 5.01, © 2008 Microsoft) via
  // fontTools OS/2: USE_TYPO_METRICS bit (fsSelection bit 7) is clear, so Word
  // uses the win metric. The Google-Fonts substitute (Noto Naskh Arabic)
  // reports ≈2.2 em from fontBoundingBox with a more ascent-heavy split, so
  // without this both the line box and the baseline position are wrong.
  [(n) => n.includes('sakkal majalla') || n.includes('majalla'), { asc: 1810 / 2048, desc: 1050 / 2048 }],
];

function lookupWinMetric(family: string | null | undefined): WinMetric | null {
  if (!family) return null;
  const n = family.toLowerCase();
  for (const [test, m] of WIN_METRICS) {
    if (test(n)) return m;
  }
  return null;
}

/**
 * Win line-height ratio (`(usWinAscent+usWinDescent)/unitsPerEm`) for a
 * requested font family, or `null` when the font is not in the table (the
 * caller should then fall back to the substituted font's Canvas metrics).
 */
export function fontWinLineHeightRatio(family: string | null | undefined): number | null {
  const m = lookupWinMetric(family);
  return m === null ? null : m.asc + m.desc;
}

/**
 * Intended single-line height in px for a run of `family` at `emPx` (the font
 * size already multiplied by the render scale), or `0` when the font is not in
 * the table. `0` is a no-op sentinel for the line-box math, which takes the
 * max of this and the substituted font's natural ascent+descent.
 */
export function intendedSingleLinePx(family: string | null | undefined, emPx: number): number {
  const ratio = fontWinLineHeightRatio(family);
  return ratio === null ? 0 : ratio * emPx;
}

/**
 * Correct a substitute font's measured Canvas `fontBoundingBox` ascent/descent
 * to the DOCUMENT font's design line box, so both the rendered line height AND
 * the baseline position within it match Word regardless of which fallback drew
 * the glyphs.
 *
 * When `family` is in the win-metric table the returned ascent/descent are the
 * document font's `usWinAscent / unitsPerEm × emPx` and
 * `usWinDescent / unitsPerEm × emPx` — i.e. the intended line box with the
 * intended baseline split, not the substitute's. This both RAISES boxes for
 * substitutes that understate the document font (Meiryo via Hiragino) and
 * SHRINKS boxes for substitutes that overstate it (Sakkal Majalla via Noto
 * Naskh Arabic, §17.4.81 / §17.4.84 cell centering). Using the document font's
 * own ascent:descent split keeps vertically-centered text on center even when
 * the substitute's split differs. When `family` is not in the table, the
 * measured metrics are returned unchanged. The `ascentPx`/`descentPx` arguments
 * are unused for tabled fonts but kept so callers can pass the substitute
 * metrics uniformly.
 */
export function correctLineMetrics(
  family: string | null | undefined,
  emPx: number,
  ascentPx: number,
  descentPx: number,
): { ascent: number; descent: number } {
  const m = lookupWinMetric(family);
  if (m === null) return { ascent: ascentPx, descent: descentPx };
  return { ascent: m.asc * emPx, descent: m.desc * emPx };
}
