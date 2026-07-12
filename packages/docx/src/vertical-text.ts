// ECMA-376 §17.6.20 vertical writing (`<w:textDirection w:val="tbRl">`) — the
// glyph-level primitives for rendering a page that has been laid out in the
// SWAPPED logical coordinate space and rotated +90° into physical space by the
// renderer's page transform (see `renderDocumentToCanvas`).
//
// After the page transform, a normal `ctx.fillText(text, x, baseline)` paints
// the run flowing DOWNWARD in physical space (logical +x → physical +y), which
// is exactly the character progression a vertical line wants — but every glyph
// is lying on its right side (rotated +90° CW with the page). The per-glyph
// orientation is decided by the Unicode UAX#50 Vertical_Orientation (vo)
// property (core `verticalOrientation`), NOT by an ad-hoc CJK-vs-Latin guess:
//   • vo=U  (upright): CJK ideographs, kana, Hangul, fullwidth forms. Drawn with
//     a local −90° counter-rotation about the glyph's own centre, cancelling the
//     page rotation so it stands UPRIGHT while still advancing down the line.
//   • vo=Tu (transform, fallback upright): 、。，．！？ and small kana. UAX#50
//     substitutes a vertical presentation glyph. For the CORNER-HANGING comma/full
//     stop (、。， → U+FE10–FE12, core `verticalFormSubstitute`) we draw THAT code
//     point upright and em-box-centred so the font's designed upper-right ink
//     placement lands the punctuation in the cell corner as Word does. ！？ are
//     NOT substituted — they stand upright and CENTRED, so the original fullwidth
//     mark drawn upright is the correct vertical form (the FE15/FE16 forms are
//     corner-designed in many fonts and would shift ！？ off-column). Where no
//     form exists (．, small kana) we draw the original upright unchanged (．
//     keeps a corner-nudge fallback).
//   • vo=Tr (transform, fallback rotate): the fullwidth brackets （「」〈〉【】… and
//     the white lenticular brackets 〖〗 have a U+FE1x/FE3x vertical presentation
//     form (core `verticalBracketFormSubstitute`) present in the substitute fonts;
//     UAX#50 §5 makes Tr "substitute a vertical glyph, ROTATE only as fallback", so
//     we SUBSTITUTE and draw them upright (Word/PowerPoint-verified, #969). Tr code
//     points with NO substituted form take a geometric fallback, one of three:
//       – ROTATE (plain): the quotes “” and the fullwidth colon ：— drawn CENTRED on
//         the column via a plain `fillText` in the +90° page frame. The rotation IS
//         the font's designed vertical form for these (font-verified: the quotes'
//         comma-hooks match, and the colon's FE13 side-by-side dots fall out of the
//         base rotation since its FE13 form is absent from most render fonts).
//       – ROTATE + REFLECT: the long-stroke marks ー (prolonged sound mark) and 〜 ～
//         (wave dash / tilde) — core `verticalTrMirrorFallback`. Their font-DESIGNED
//         vertical glyph is the HORIZONTAL REFLECTION of the +90° rotation, not the
//         rotation (the 起筆/curvature flips left↔right between orientations — a
//         documented Japanese typographic convention; Word PDF sample-47 + font `vert`
//         glyph verified). A Canvas cannot reach the `vert` glyph, so we rotate AND
//         reflect via `scale(1, -1)` about the cell centre (the on-screen horizontal
//         mirror in the +90° page frame).
//       – UPRIGHT: the fullwidth semicolon ；, whose FE14 form is an upright dot-over-
//         comma, not a rotation (issue #969 follow-up; core `verticalTrUprightFallback`).
//   • vo=R  (rotated): Latin letters, Western digits, Latin punctuation. Stay
//     SIDEWAYS (rotated with the page) — the conventional "縦中横 not applied"
//     appearance — drawn as an ordinary contextual `fillText` at the alphabetic
//     baseline, preserving the browser's shaping/advance for the run.
//
// This module owns ONLY the pure geometry + classification; the renderer wires
// it into the whole-run glyph draw sites, the anchor/inline/float image draws,
// and the text-selection overlay behind the `verticalCJK` flag, so the
// horizontal path stays byte-identical.
//
// SCOPE (issue #771). Implemented: +90° page rotation; vo-driven upright(U) /
// substituted-upright(Tu) / rotated(Tr) / sideways(R) glyph draw; anchor images
// resolved against the physical page then projected into the logical flow
// (PDF-verified centroid); inline/anchored/float image uprighting; and the
// vertical text-layer transform. Still approximated / deferred (flagged inline):
// the `0.12em` upright-centring nudge and the Tu upper-right corner nudge are
// font-dependent stage-1 heuristics; paragraph-relative vertical anchors are a
// follow-up. `btLr` shares the +90° page FRAME but bypasses this module's
// upright/substitute glyph handling entirely (issue #988 re-adjudication: every
// glyph rides the page rotation — see RenderState.verticalAllRotated).

import {
  verticalOrientation,
  verticalFormSubstitute,
  verticalBracketFormSubstitute,
  verticalTrUprightFallback,
  verticalTrMirrorFallback,
  verticalVertFeatureSupported,
  withVertFeature,
} from '@silurus/ooxml-core';

/** How a code point is painted inside the +90°-rotated vertical page:
 *   - `upright`  — counter-rotated −90° to stand up (vo=U, and vo=Tu).
 *   - `rotate`   — left rotated with the page but CENTRED on the column axis,
 *                  the UAX#50 Tr fallback for full-width brackets / 長音符.
 *   - `sideways` — left rotated with the page at the alphabetic baseline (vo=R,
 *                  Latin/digits). */
export type VerticalDrawMode = 'upright' | 'rotate' | 'sideways';

/**
 * The draw mode for a code point in vertical text, from its UAX#50
 * Vertical_Orientation (vo). Single source of truth: core `verticalOrientation`.
 *
 *   U  → upright   (stand the glyph up)
 *   Tu → upright   (draw upright; the caller substitutes a vertical form glyph
 *                   via {@link verticalFormSubstitute} when one exists so the
 *                   comma/full stop land in the upper-right of the cell)
 *   Tr → rotate    (rotate 90° CW — the fallback when no vertical glyph is
 *                   reachable on a Canvas — centred on the column)
 *   R  → sideways  (leave rotated with the page: Latin/digits)
 *
 * @param cp A Unicode scalar value (e.g. from `String.prototype.codePointAt`).
 */
export function verticalDrawMode(cp: number): VerticalDrawMode {
  const vo = verticalOrientation(cp);
  if (vo === 'U' || vo === 'Tu') return 'upright';
  if (vo === 'Tr') return 'rotate';
  return 'sideways'; // vo === 'R'
}

/**
 * True when `cp` stands UPRIGHT in vertical text (UAX#50 vo ∈ {U, Tu}). Kept for
 * callers that only need the upright/not-upright split; new code should prefer
 * {@link verticalDrawMode} which also distinguishes the Tr (rotate) case.
 *
 * @param cp A Unicode scalar value (e.g. from `String.prototype.codePointAt`).
 */
export function isUprightVerticalGlyph(cp: number): boolean {
  return verticalDrawMode(cp) === 'upright';
}

/** The Tu punctuation whose upper-right cell position is approximated by a
 *  draw-time nudge WHEN the font has no U+FExx vertical form to substitute (see
 *  {@link verticalGlyphOffset}). The comma/full stop that DO have a vertical form
 *  ({@link verticalFormSubstitute}: 、。，) are substituted instead and the font
 *  positions them, so they are NOT nudged. The fullwidth full stop ． (FF0E) has
 *  no vertical form in Unicode, so it stays on the nudge fallback. */
const VERTICAL_PUNCT_UPPER_RIGHT = new Set<number>([
  0xff0e, // ． fullwidth full stop (no U+FExx vertical form → nudge fallback)
]);

/**
 * Per-glyph draw offset (in em fractions of the font size) applied in the
 * glyph's own UPRIGHT local frame — i.e. after the −90° counter-rotation, in
 * physical (dx = rightward, dy = downward) terms. Returns `{ dx, dy }` em
 * fractions; the caller multiplies by the font px size.
 *
 * This is the FALLBACK for a Tu code point whose upper-right cell position would
 * otherwise be supplied by a substituted vertical presentation form
 * ({@link verticalFormSubstitute}) but for which Unicode has none (only ． FF0E
 * today). The nudge moves the glyph toward the upper-right corner of the cell.
 * Everything with a vertical form is substituted and returns `{0,0}` here.
 */
export function verticalGlyphOffset(cp: number): { dx: number; dy: number } {
  if (VERTICAL_PUNCT_UPPER_RIGHT.has(cp)) {
    // HEURISTIC (approximation, font-dependent): move ． toward the upper-right
    // corner of the cell by ~0.4em each way. NOT a spec constant — JIS X 4051
    // §4.x gives the punctuation cell geometry (the glyph occupies a quarter-em
    // corner box), not a 0.4em nudge. The correct fix would be a Unicode vertical
    // form for ． (none exists), so this narrow fallback remains. Tracked in issue
    // #771 (vertical-text).
    return { dx: 0.4, dy: -0.4 };
  }
  return { dx: 0, dy: 0 };
}

/**
 * Split a run's text into maximal runs of same-draw-mode code points (UAX#50 vo,
 * via {@link verticalDrawMode}), so the vertical draw path can counter-rotate
 * the UPRIGHT segments per glyph, rotate the Tr segments, and draw the SIDEWAYS
 * (Latin/digit) segments as a single contextual `fillText`. Preserves surrogate
 * pairs (iterates by code point) and returns the pieces in logical order.
 *
 * @param text The run's text.
 * @returns Ordered pieces, each `{ text, mode }`.
 */
export function splitVerticalOrientationRuns(
  text: string,
): Array<{ text: string; mode: VerticalDrawMode }> {
  const pieces: Array<{ text: string; mode: VerticalDrawMode }> = [];
  let cur = '';
  let curMode: VerticalDrawMode | null = null;
  for (const ch of text) {
    const cp = ch.codePointAt(0) ?? 0;
    const mode = verticalDrawMode(cp);
    if (curMode === null) {
      curMode = mode;
      cur = ch;
    } else if (mode === curMode) {
      cur += ch;
    } else {
      pieces.push({ text: cur, mode: curMode });
      cur = ch;
      curMode = mode;
    }
  }
  if (cur !== '' && curMode !== null) {
    pieces.push({ text: cur, mode: curMode });
  }
  return pieces;
}

type Ctx2D = CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D;

/**
 * Cross-axis (column-thickness) offset, in px, from the alphabetic baseline to
 * the font's EM-BOX CENTRE — i.e. `(fontBoundingBoxAscent − fontBoundingBoxDescent)/2`.
 *
 * This is a FONT metric (glyph-independent): `fontBoundingBox*` describe the
 * font's design box, not one glyph's ink. In vertical text the column's cross
 * axis is where every cell centres, and the UPRIGHT cells (drawn with a `middle`
 * textBaseline) already sit their em box on the caller's `baseline`. A SIDEWAYS
 * (Latin/digit) glyph, however, is drawn on its ALPHABETIC baseline, so its ink
 * (which sits ~this many px above the baseline) would land off the column centre
 * by exactly this amount. Shifting the sideways draw down the cross axis by this
 * offset re-centres its em box on the same line the upright cells use — so mixed
 * columns like "電話 03-1234-5678" share one centreline (ECMA-376 §17.6.20).
 *
 * Falls back to `0.38 × fontPx` (the near-universal CJK/Latin em-box centre ratio)
 * only if the Canvas does not report `fontBoundingBox*` (older engines); on those
 * engines the previous baseline-anchored placement is no worse than today.
 */
function emBoxCenterAboveBaselinePx(ctx: Ctx2D, sample: string, fontPx: number): number {
  const prevBaseline = ctx.textBaseline;
  ctx.textBaseline = 'alphabetic';
  const m = ctx.measureText(sample);
  ctx.textBaseline = prevBaseline;
  const asc = m.fontBoundingBoxAscent;
  const desc = m.fontBoundingBoxDescent;
  if (typeof asc === 'number' && typeof desc === 'number' && (asc !== 0 || desc !== 0)) {
    return (asc - desc) / 2;
  }
  return 0.38 * fontPx;
}

/**
 * Along-column offset, in px, from a glyph's own cell centre to its INK centre
 * when the glyph is drawn UPRIGHT — i.e. `(actualBoundingBoxAscent −
 * actualBoundingBoxDescent)/2` measured with a `middle` textBaseline.
 *
 * For an upright-drawn glyph the page transform maps the glyph's VERTICAL extent
 * onto the along-column axis, so a glyph whose ink is not vertically centred in
 * its em box (most visibly a substituted vertical bracket form ︵ ︶ ﹁ ﹂,
 * whose ink hugs one end of the cell) lands off the cell centre by this amount.
 * The renderer shifts the draw by `+this` so the ink re-centres — a per-GLYPH
 * measured metric (`actualBoundingBox*` is the tight ink box), NOT a constant.
 * For an ordinary ideograph/kana this is ≈0, so upright CJK cells are unaffected.
 *
 * Returns 0 when the Canvas does not report `actualBoundingBox*` (older engines):
 * the glyph then draws at the cell centre exactly as before this metric existed.
 */
function inkCenterAboveMiddlePx(ctx: Ctx2D, drawStr: string): number {
  const prevAlign = ctx.textAlign;
  const prevBaseline = ctx.textBaseline;
  ctx.textAlign = 'center';
  ctx.textBaseline = 'middle';
  const m = ctx.measureText(drawStr);
  ctx.textAlign = prevAlign;
  ctx.textBaseline = prevBaseline;
  const asc = m.actualBoundingBoxAscent;
  const desc = m.actualBoundingBoxDescent;
  if (typeof asc === 'number' && typeof desc === 'number') {
    return (asc - desc) / 2;
  }
  return 0;
}

/**
 * True for a vo=Tr code point that takes the GEOMETRIC ROTATE fallback in
 * {@link drawVerticalRun} — i.e. `mode==='rotate'` with NO substituted vertical
 * bracket form and NOT the upright-fallback semicolon. These are the marks drawn
 * by a plain (optionally reflected) `fillText` in the +90° page frame: ー 〜 ～,
 * the quotes “”, and the colon ：. This is the SINGLE predicate shared by the
 * paint path and the {@link verticalRunInkExtraPx} measure path (issue #1014), so
 * the two agree on which glyphs get the ink-sized cell.
 *
 * @param cp A Unicode scalar value.
 */
function isVerticalRotateFallback(cp: number): boolean {
  return (
    verticalDrawMode(cp) === 'rotate' &&
    verticalBracketFormSubstitute(cp) === null &&
    !verticalTrUprightFallback(cp)
  );
}

/**
 * Along-column ink geometry of a vo=Tr rotate-fallback glyph (issue #1014). The
 * glyph is painted by a plain `fillText` in the +90°-rotated page frame, so its
 * HORIZONTAL ink extent maps onto the ALONG-COLUMN axis (the advance axis). Read
 * the tight horizontal ink box with a `center`/`middle` alignment:
 *   - `extentPx` = actualBoundingBoxLeft + actualBoundingBoxRight — the ink width
 *     along the column (used to size the cell so the ink cannot spill past it).
 *   - `shiftPx`  = (actualBoundingBoxLeft − actualBoundingBoxRight)/2 — the local
 *     along-column shift that re-centres the ink on the (grown) cell, since a
 *     `center` draw centres the glyph's ADVANCE and an under-reported advance is
 *     off-centre from the ink.
 * Returns `null` when the Canvas does not report `actualBoundingBox*` (older
 * engines / node mocks) so callers degrade to the advance-sized, advance-centred
 * draw exactly as before this metric existed.
 */
function verticalRotateInkGeometry(
  ctx: Ctx2D,
  ch: string,
): { extentPx: number; shiftPx: number } | null {
  const prevAlign = ctx.textAlign;
  const prevBaseline = ctx.textBaseline;
  ctx.textAlign = 'center';
  ctx.textBaseline = 'middle';
  const m = ctx.measureText(ch);
  ctx.textAlign = prevAlign;
  ctx.textBaseline = prevBaseline;
  const l = m.actualBoundingBoxLeft;
  const r = m.actualBoundingBoxRight;
  if (
    typeof l !== 'number' ||
    typeof r !== 'number' ||
    !Number.isFinite(l) ||
    !Number.isFinite(r)
  ) {
    return null;
  }
  return { extentPx: l + r, shiftPx: (l - r) / 2 };
}

/**
 * ECMA-376 §17.6.20 (tbRl) + issue #1014 — the EXTRA along-column advance (px, at
 * the run's current font, BEFORE the §17.3.2.43 `w:w` scale and §17.3.2.35 pitch)
 * that a vertical run needs so its vo=Tr rotate-fallback glyphs' INK fits inside
 * their cells. For each rotate-fallback glyph whose along-column ink extent
 * exceeds its `measureText` advance (a substitute font UNDER-REPORTING the
 * advance — Chrome-only; skia and normal fonts report ink ≤ advance), this adds
 * the deficit `max(0, inkExtent − advance)`. The layout folds this into the
 * segment's natural advance (`segAdvanceWidth`'s `naturalWidthPx`) so the grown
 * cell {@link drawVerticalRun} paints is matched by the measured box — measure ==
 * paint (wrapping, the next run's position, and the selection overlay all track
 * the drawn cell). Returns 0 for a run with no such glyph, for every font that
 * does not under-report (the common path — byte-identical), and when ink metrics
 * are unavailable.
 *
 * The caller must set `ctx.font` (and any kerning state) for the run before
 * calling, exactly as it does for the `measureText` that produces `naturalWidthPx`.
 *
 * @param ctx  2D context with the run's font selected.
 * @param text The run's text.
 */
export function verticalRunInkExtraPxWithCapability(
  ctx: Ctx2D,
  text: string,
  vertCapable: boolean,
): number {
  let extra = 0;
  for (const ch of text) {
    const cp = ch.codePointAt(0) ?? 0;
    if (!isVerticalRotateFallback(cp)) continue;
    // A real vert long-stroke glyph is upright in its ordinary one-em cell. All
    // other rotate-fallback glyphs (notably quotes and the colon) still use the
    // geometric path and therefore retain #1014/#1019 ink growth.
    if (vertCapable && verticalTrMirrorFallback(cp)) continue;
    const geom = verticalRotateInkGeometry(ctx, ch);
    if (geom === null) continue;
    const advance = ctx.measureText(ch).width;
    if (geom.extentPx > advance) extra += geom.extentPx - advance;
  }
  return extra;
}

export function verticalRunInkExtraPx(ctx: Ctx2D, text: string): number {
  return verticalRunInkExtraPxWithCapability(ctx, text, verticalVertFeatureSupported(ctx));
}

/**
 * Draw one run's glyphs in vertical mode. The context is assumed to already be
 * in the page's SWAPPED logical frame (the +90° page rotation is installed by
 * `renderDocumentToCanvas`), so an ordinary `fillText` advances DOWN the line.
 *
 * The run flows along logical +x (physical +y). Each glyph occupies a cell of
 * width = its horizontal advance (`ctx.measureText`) plus `letterSpacingPx`
 * (the docGrid / justification pitch the layout measured the box with), so the
 * total advance equals the run's measured width — measure == draw. Upright
 * (CJK) glyphs are counter-rotated −90° about their cell centre so they stand
 * upright; sideways (Latin/digit) pieces are painted as a single contextual
 * `fillText`, preserving the browser's shaping.
 *
 * @param ctx              2D context, already in the rotated logical page frame.
 *                         `ctx.font`/`ctx.fillStyle` are set by the caller.
 * @param text             The run's text.
 * @param x                Logical left edge of the run (px).
 * @param baseline         Logical baseline y of the line (px).
 * @param fontPx           Effective font size in px (for cell centring).
 * @param letterSpacingPx  Per-glyph extra advance: the combined docGrid cell
 *                         delta + §17.3.2.35 `w:spacing` pitch (the layout's
 *                         `segLetterSpacingPx`); 0 for the common path.
 * @param charScale        ECMA-376 §17.3.2.43 `w:w` fraction; 1 by default.
 * @param growTrRotateInk  issue #1014 — when true, a vo=Tr GEOMETRIC rotate-fallback
 *                         glyph (ー 〜 ～ “” ：) whose substitute font under-reports
 *                         its advance is sized to its along-column INK extent (and
 *                         ink-centred) so its ink cannot spill past the cell into the
 *                         next run. MUST be set ONLY where the layout advance was
 *                         grown by the SAME deficit ({@link verticalRunInkExtraPx},
 *                         gated on `LayoutTextSeg.verticalRun`) so paint == measure;
 *                         the caller passes `s.verticalRun === true`. Default false
 *                         keeps the advance-sized, advance-centred draw byte-identical
 *                         (markers and unwired vertical text boxes).
 */
export function drawVerticalRunWithCapability(
  ctx: Ctx2D,
  text: string,
  x: number,
  baseline: number,
  fontPx: number,
  letterSpacingPx: number,
  charScale = 1,
  growTrRotateInk = false,
  vertCapable = false,
): void {
  const prevAlign = ctx.textAlign;
  const prevBaseline = ctx.textBaseline;
  // Cross-axis (column-thickness) distance from the alphabetic baseline to the
  // font's em-box centre — the line the UPRIGHT cells centre on. Measured once
  // per run (font-level, glyph-independent). Used to re-centre SIDEWAYS glyphs,
  // which are otherwise drawn on their baseline and so land off the centreline.
  const emBoxCenterPx = emBoxCenterAboveBaselinePx(ctx, text, fontPx);
  // In this rotate-layout architecture the direction-independent layout kernel
  // applies ECMA-376 §17.3.2.43 `w:w` to the line axis (`segAdvanceWidth`) even
  // for tbRl; wrapping, selection, and run boxes already depend on that advance,
  // so paint must follow measure. Sideways glyphs are rotated horizontal text,
  // making their `w:w` width axis the vertical advance axis directly. Upright
  // glyphs therefore scale the equivalent local y axis. Tate-chu-yoko is kept
  // separate: ECMA-376 §17.3.2.10 fixes its advance to one em and uses `w:w` on
  // the cross axis.
  const scaled = charScale !== 1;
  let ax = 0; // cumulative advance from run left (logical +x)
  for (const ch of text) {
    const cp = ch.codePointAt(0) ?? 0;
    const mode = verticalDrawMode(cp);
    // A vo=Tr code point with a substituted Unicode vertical presentation form — the
    // brackets （）「」〈〉… and the white lenticular 〖〗 (#969) — is SUBSTITUTED and
    // drawn upright, exactly like the upright cells — UAX#50 §5 Tr means "substitute a
    // vertical glyph; rotate only as fallback". Tr code points with NO substituted form
    // (ー, quotes “”, and the colon ：/ semicolon ；whose FE13/FE14 forms are absent
    // from most render fonts) take a geometric fallback below (rotate, or — for ；—
    // upright).
    const bracketCp = mode === 'rotate' ? verticalBracketFormSubstitute(cp) : null;
    // A vo=Tr code point with NO substituted vertical form whose fallback is
    // UPRIGHT rather than the generic UAX#50 §5 ROTATE — the fullwidth semicolon
    // ；(FF1B), whose FE14 vertical form is an upright dot-over-comma, not a
    // rotation (Word-verified, issue #969 follow-up). It draws upright exactly
    // like the vo=U / vo=Tu cells; the colon ：is NOT here (its FE13 form IS a 90°
    // rotation, so it takes the rotate branch below → side-by-side dots).
    const uprightFallback = mode === 'rotate' && bracketCp === null && verticalTrUprightFallback(cp);
    // Advance/width uses the ORIGINAL code point (measure == draw, and the text
    // model / selection / find keep the original character — see the module doc).
    // #1014: a vo=Tr GEOMETRIC rotate-fallback glyph (ー 〜 ～ “” ：) is painted by a
    // plain `fillText` in the +90° page frame, so its HORIZONTAL ink maps onto the
    // along-column (advance) axis. When a substitute font UNDER-REPORTS the advance
    // (Chrome), that ink spills PAST the advance-sized cell into the next run. Size
    // the cell to the along-column INK extent instead so the ink is contained; the
    // SAME per-glyph deficit is folded into the layout advance by
    // `verticalRunInkExtraPx` (measure == draw). NO-OP unless the ink exceeds the
    // advance (every real font here reports ink ≤ advance ⇒ byte-identical), and only
    // for the geometric rotate branch (substituted/upright Tr glyphs keep their path).
    let cellNaturalPx = ctx.measureText(ch).width;
    let rotateInkShiftPx = 0;
    if (
      !(vertCapable && verticalTrMirrorFallback(cp)) &&
      growTrRotateInk &&
      mode === 'rotate' &&
      bracketCp === null &&
      !uprightFallback
    ) {
      const geom = verticalRotateInkGeometry(ctx, ch);
      if (geom !== null && geom.extentPx > cellNaturalPx) {
        cellNaturalPx = geom.extentPx;
        rotateInkShiftPx = geom.shiftPx;
      }
    }
    const adv = cellNaturalPx * charScale + letterSpacingPx;
    if (vertCapable && verticalTrMirrorFallback(cp)) {
      // The font's `vert` table supplies the designed upright form for the three
      // long-stroke marks whose geometric fallback otherwise needs reflection.
      // Keep every other glyph on its Word-adjudicated manual path below.
      const cx = x + ax + adv / 2;
      ctx.save();
      ctx.translate(cx, baseline);
      ctx.rotate(-Math.PI / 2);
      if (scaled) ctx.scale(1, charScale);
      ctx.textAlign = 'center';
      ctx.textBaseline = 'middle';
      withVertFeature(ctx, () => ctx.fillText(ch, 0, 0));
      ctx.restore();
    } else if (mode === 'upright' || bracketCp !== null || uprightFallback) {
      // vo=U / Tu, or a substituted Tr bracket. Counter-rotate −90° about the
      // cell centre so the glyph (which the page rotation would otherwise lay on
      // its side) stands upright. For corner-hanging Tu punctuation with a Unicode
      // vertical form (、。， → U+FE10–FE12) and Tr brackets (（）「」… → U+FE35–FE44)
      // draw THAT glyph so the font supplies the vertical shape; the original
      // advance is kept. Substitution is a GLYPH-only change: the width above and
      // everything the renderer reports (selection, find) use the original `ch`.
      // ！？ are NOT substituted (see verticalFormSubstitute) — they draw upright
      // as the original fullwidth mark, which is already centred on the column.
      // A substituted Tu punctuation form (、。， → FE10–FE12) vs. everything else.
      // The Tr bracket substitute is tracked separately by `bracketCp`.
      const puncCp = bracketCp !== null ? null : verticalFormSubstitute(cp);
      const drawCp = bracketCp !== null ? bracketCp : puncCp;
      const drawStr = drawCp !== null ? String.fromCodePoint(drawCp) : ch;
      const cx = x + ax + adv / 2;
      // Corner nudge fallback only for a Tu punct with NO vertical form (． FF0E);
      // every substituted glyph is positioned by its own vertical metric below.
      const off = drawCp !== null ? { dx: 0, dy: 0 } : verticalGlyphOffset(cp);
      // ALONG-COLUMN centring: an upright glyph's VERTICAL ink extent maps to the
      // column axis, so shift by its measured ink centre so the ink lands on the
      // cell centre. Per-GLYPH metric (the drawn glyph's tight ink box): for an
      // ideograph/kana it is ≈0 (cells unchanged); for a substituted vertical
      // bracket (ink hugging one cell end) it is the needed correction. Replaces
      // the old `+0.12em` font-tuned heuristic. Skipped when the ． corner nudge is
      // active (`off.dy`), which is a self-contained upper-right cell placement.
      //
      // NOT applied to a substituted Tu punctuation form (comma/full stop 、。，
      // → FE10–FE12): those glyphs are DESIGNED with their ink in the cell's
      // upper-right corner (JIS X 4051 §4.3 kutōten placement — Word keeps them
      // there, PDF-verified on sample-26: 、 ink at −0.32em along-column). Ink-
      // centring would force that intentional offset back to the geometric cell
      // centre, dropping the comma/full stop LOW — the reported "、。 sit too low"
      // defect (#771). Drawing them em-box-centred preserves the font's corner
      // design. The Tr brackets DO get the correction: their two halves must sit a
      // full cell apart and the font centres the em box, not the ink (#792).
      const isPunctSubstitute = puncCp !== null;
      const alongEm =
        off.dy === 0 && !isPunctSubstitute
          ? inkCenterAboveMiddlePx(ctx, drawStr) / fontPx
          : 0;
      ctx.save();
      ctx.translate(cx, baseline);
      ctx.rotate(-Math.PI / 2);
      if (scaled) ctx.scale(1, charScale);
      // In the upright local frame: `center`/`middle` puts the em box on the cell
      // centre; local +x = cross axis, local +y = along-column. `off.dx` nudges
      // ． toward the cell's upper-right corner (cross axis); `alongEm + off.dy`
      // centres the ink along the column.
      ctx.textAlign = 'center';
      ctx.textBaseline = 'middle';
      ctx.fillText(drawStr, off.dx * fontPx, (alongEm + off.dy) * fontPx);
      ctx.restore();
    } else if (mode === 'rotate') {
      // vo=Tr with NO substituted vertical form and NOT the upright-fallback
      // semicolon: ー (U+30FC), the wave dash / tilde 〜 ～, the double quotes “”,
      // and the fullwidth colon ：(FF1A). UAX#50's Tr fallback (no vertical glyph
      // reachable on a Canvas) is to ROTATE the glyph 90° CW; a plain `fillText` in
      // the +90° page frame IS that rotation, centred on the column with
      // `center`/`middle` at the cell centre. For the colon this reproduces FE13's
      // design directly (the two vertically-stacked dots become side by side),
      // Word-verified (issue #969 follow-up); for the quotes the rotation matches the
      // font's designed vertical form exactly (font-verified).
      //
      // The long-stroke marks ー and 〜 ～ (verticalTrMirrorFallback) are the
      // EXCEPTION: their font-DESIGNED vertical form is the HORIZONTAL REFLECTION of
      // that +90° rotation, not the rotation — the 起筆/curvature flips left↔right
      // between orientations (Word PDF sample-47 + font `vert` glyph verified: a plain
      // rotation of ー bulges LEFT, Word/the designed glyph bulge RIGHT). Since a
      // Canvas cannot invoke the font's `vert` OpenType glyph, we reproduce it by
      // reflecting: in the +90° page frame the on-screen horizontal mirror is
      // `scale(1, -1)` about the cell centre. Advance/measure and the column centring
      // are unchanged (the em box is symmetric about the cell centre), so only the
      // glyph's chirality flips. (Substituted bracket forms never reach here — they
      // were drawn upright above; a rotated bracket's ink offset is not measurable.)
      const cx = x + ax + adv / 2;
      const mirror = verticalTrMirrorFallback(cp);
      ctx.textAlign = 'center';
      ctx.textBaseline = 'middle';
      // #1014: `rotateInkShiftPx` (glyph-space, non-zero ONLY when the cell was grown
      // to the ink extent above) re-centres the ink on the grown cell — a `center`
      // draw centres the glyph's ADVANCE, and an under-reported advance is off-centre
      // from the ink. It rides the local frame so it composes with the §17.3.2.43
      // `w:w` scale and the reflection; 0 (the common path) leaves today's advance-
      // centred draw byte-identical.
      if (mirror || scaled || rotateInkShiftPx !== 0) {
        ctx.save();
        ctx.translate(cx, baseline);
        // `scale(1, -1)` is the on-screen horizontal mirror in the +90° page frame
        // (screen −x ↔ page-frame +y); combine with the §17.3.2.43 `w:w` width
        // compression on the line axis. Non-mirror glyphs keep sy=+1.
        ctx.scale(charScale, mirror ? -1 : 1);
        ctx.fillText(ch, rotateInkShiftPx, 0);
        ctx.restore();
      } else {
        ctx.fillText(ch, cx, baseline);
      }
    } else {
      // vo=R (Latin/digit): drawn SIDEWAYS (rotated with the page). Keep the
      // caller's alphabetic baseline and position the glyph's left at the current
      // advance, but shift the cross axis by the em-box centre so the glyph's ink
      // centres on the SAME column centreline the upright cells use — otherwise a
      // baseline-anchored sideways glyph sits ~0.38em off (its ink is above the
      // baseline), the "電話 / 03-…" left-right drift. A group of consecutive
      // sideways glyphs would ideally be one fillText for shaping, but per-glyph
      // keeps the advance model uniform and Latin advances are context-free enough
      // at these sizes.
      // `alphabetic` baseline pins the em-box-centre shift (emBoxCenterPx is
      // measured relative to the alphabetic baseline).
      ctx.textAlign = prevAlign;
      ctx.textBaseline = 'alphabetic';
      if (scaled) {
        ctx.save();
        ctx.translate(x + ax, 0);
        ctx.scale(charScale, 1);
        ctx.fillText(ch, 0, baseline + emBoxCenterPx);
        ctx.restore();
      } else {
        ctx.fillText(ch, x + ax, baseline + emBoxCenterPx);
      }
    }
    ax += adv;
  }
  ctx.textAlign = prevAlign;
  ctx.textBaseline = prevBaseline;
}

export function drawVerticalRun(
  ctx: Ctx2D,
  text: string,
  x: number,
  baseline: number,
  fontPx: number,
  letterSpacingPx: number,
  charScale = 1,
  growTrRotateInk = false,
): void {
  const vertCapable = verticalVertFeatureSupported(ctx);
  drawVerticalRunWithCapability(
    ctx,
    text,
    x,
    baseline,
    fontPx,
    letterSpacingPx,
    charScale,
    growTrRotateInk,
    vertCapable,
  );
}

/**
 * Draw one 縦中横 (tate-chū-yoko / horizontal-in-vertical) run — ECMA-376
 * §17.3.2.10 `<w:eastAsianLayout w:vert="1">`. In a vertical (tbRl) page the run
 * "keeps the text on the same line as all other text" while its characters are
 * rendered HORIZONTALLY: the whole run string is drawn UPRIGHT (counter-rotated
 * −90° to cancel the +90° page rotation, exactly like the upright CJK cells) so
 * the glyphs read left-to-right ACROSS the column, packed into ONE cell of the
 * vertical line.
 *
 * Geometry (all in the rotated logical page frame; `x` advances DOWN the column
 * = logical +x, `baseline` is the column centre-line = logical +y):
 *   - The cell spans `[x, x + cellAdvance]` along the column; the run centres on
 *     the cell centre `x + cellAdvance/2`. `cellAdvance` is one em (one cell) —
 *     the same value the layout measured (`segAdvanceWidth`), so measure==paint.
 *   - After the −90° counter-rotation the run is upright; local +x is the
 *     cross-column (the glyphs' own left→right width) and local +y is the
 *     along-column (the text's height). Drawn `center`/`middle`, so the run's
 *     em box centres on the cell centre AND on the column centre-line.
 *   - `charScale` (§17.3.2.43 `w:w`) compresses the glyphs' WIDTH via
 *     `ctx.scale(charScale, 1)` in the upright local frame — i.e. across the
 *     column — matching Word (PDF-verified on sample-26: "２９" at w:w=67 spans
 *     ≈15.6 pt wide inside a 12 pt cell). It does NOT change the along-column
 *     cell height.
 *   - `vertCompress` (§17.3.2.10) compresses the run's HEIGHT to one cell so the
 *     rotated text never grows the line: if the run's natural upright height
 *     (`fontBoundingBox*`) exceeds one em, scale the along-column axis down to
 *     fit. For a single-line run (height ≈ 1 em) this is a no-op, so the common
 *     2-digit date case is unaffected; it only bites a run whose glyphs are
 *     taller than the em box.
 *
 * The whole run is one contextually-shaped `fillText`, so kerning/shaping across
 * the digits is preserved and the text model / selection keep the original
 * characters.
 *
 * @param ctx          2D context, already in the rotated logical page frame.
 *                     `ctx.font`/`ctx.fillStyle` are set by the caller.
 * @param text         The run's text (e.g. "２９").
 * @param x            Logical left edge of the cell along the column (px).
 * @param baseline     Logical column centre-line y (px).
 * @param fontPx       Effective font size in px (one em = one cell).
 * @param cellAdvance  The cell's along-column advance in px (one em; the value
 *                     the layout measured for this segment).
 * @param charScale    §17.3.2.43 `w:w` fraction (1 = 100%); compresses the
 *                     glyphs' cross-column width.
 * @param compress     §17.3.2.10 `w:vertCompress`; fit the run's height to one em.
 */
export function drawTateChuYokoRun(
  ctx: Ctx2D,
  text: string,
  x: number,
  baseline: number,
  fontPx: number,
  cellAdvance: number,
  charScale: number,
  compress: boolean,
): void {
  const prevAlign = ctx.textAlign;
  const prevBaseline = ctx.textBaseline;
  // Along-column compression factor (§17.3.2.10 vertCompress). The run is drawn
  // upright, so its along-column extent is the font's design HEIGHT
  // (fontBoundingBoxAscent + descent). When that exceeds one em and vertCompress
  // is set, scale the along-column (local y) axis so the height fits one cell.
  // Measured with a `middle` baseline (the box used below). For ordinary
  // single-line text the height is ≈1 em, so `compY` stays 1 (no-op).
  let compY = 1;
  if (compress) {
    const m = ctx.measureText(text);
    const asc = m.fontBoundingBoxAscent;
    const desc = m.fontBoundingBoxDescent;
    if (typeof asc === 'number' && typeof desc === 'number') {
      const heightPx = asc + desc;
      if (heightPx > fontPx && heightPx > 0) compY = fontPx / heightPx;
    }
  }
  const cx = x + cellAdvance / 2;
  ctx.save();
  ctx.translate(cx, baseline);
  ctx.rotate(-Math.PI / 2);
  // Upright local frame: local +x = cross-column (glyph width), local +y =
  // along-column (glyph height). `w:w` compresses width (local x); vertCompress
  // fits height (local y). center/middle centres the run's em box on the cell.
  ctx.scale(charScale, compY);
  ctx.textAlign = 'center';
  ctx.textBaseline = 'middle';
  ctx.fillText(text, 0, 0);
  ctx.restore();
  ctx.textAlign = prevAlign;
  ctx.textBaseline = prevBaseline;
}

/**
 * Run `draw` with the context counter-rotated so a graphic that would otherwise
 * be painted lying on its side (rotated with the +90° page transform) appears
 * UPRIGHT. Used for inline / anchored images and shapes in a vertical (tbRl)
 * page: an image is not text, so it keeps its natural upright orientation even
 * though the surrounding characters advance downward.
 *
 * The box is specified by its logical top-left `(x, y)` and logical size
 * `(w, h)` — the same coordinates the horizontal draw path uses. We rotate −90°
 * about the box centre (cancelling the page rotation) and invoke `draw(dx, dy,
 * dw, dh)` with the box re-expressed in the upright local frame: the logical
 * width becomes the local HEIGHT and vice-versa, so the caller draws the image
 * at `(-h/2, -w/2, h, w)` centred on the pivot. The net effect places an upright
 * image inside the rotated page footprint.
 *
 * @param ctx  2D context already in the rotated logical page frame.
 * @param x    Logical left of the box (px).
 * @param y    Logical top of the box (px).
 * @param w    Logical width of the box (px).
 * @param h    Logical height of the box (px).
 * @param draw Callback painting the graphic at `(dx, dy, dw, dh)` in the upright
 *             local frame.
 */
export function drawUprightBox(
  ctx: Ctx2D,
  x: number,
  y: number,
  w: number,
  h: number,
  draw: (dx: number, dy: number, dw: number, dh: number) => void,
): void {
  const cx = x + w / 2;
  const cy = y + h / 2;
  ctx.save();
  ctx.translate(cx, cy);
  ctx.rotate(-Math.PI / 2);
  // In the upright local frame the logical width spans the local y-axis and the
  // logical height spans the local x-axis, so the box is (−h/2, −w/2, h, w).
  draw(-h / 2, -w / 2, h, w);
  ctx.restore();
}

/** A rectangle `{ x, y, w, h }` in some coordinate frame (px). */
export interface Box {
  x: number;
  y: number;
  w: number;
  h: number;
}

/**
 * Map a DrawingML anchor's box from PHYSICAL page space into the SWAPPED LOGICAL
 * layout frame the vertical (tbRl) renderer flows text in (ECMA-376 §17.6.20 +
 * §20.4.3.x).
 *
 * A `<wp:positionH>` / `<wp:positionV>` anchor is resolved against the PHYSICAL
 * page — Word places the drawing layer before/independently of the text-flow
 * rotation, so the image stays upright at physical `(px, py, w, h)` exactly as in
 * a horizontal document. The body text, however, is laid out in the logical frame
 * that the page paint transform `physical = (cssWidth − logical.y, logical.x)`
 * maps to physical. Inverting that transform (`logical.x = physical.y`,
 * `logical.y = cssWidth − physical.x`) projects the physical image rectangle onto
 * the logical frame:
 *   - logical x-range = `[py, py + h]`         (physical y ↦ logical x, downward)
 *   - logical y-range = `[cssWidth − (px + w), cssWidth − px]`
 *                                               (physical x ↦ logical y, reversed)
 * so the logical box has `w ↔ h` swapped: logical width = physical height and
 * logical height = physical width.
 *
 * The returned box drives BOTH the float-exclusion rectangle (text wraps around
 * this logical projection, in the same frame as the flow) AND {@link drawUprightBox}
 * (which un-swaps it back to the upright physical image). Because the two derive
 * from one box, the wrap band and the painted image stay locked together
 * (packages/docx/CLAUDE.md — no duplicated geometry).
 *
 * @param px         Physical left of the image box (px).
 * @param py         Physical top of the image box (px).
 * @param w          Physical image width (px).
 * @param h          Physical image height (px).
 * @param cssWidthPx The canvas CSS width in px (= physical page width) — the
 *                   `translate(cssWidth, 0)` term of the page transform.
 */
export function physicalToLogicalAnchorBox(
  px: number,
  py: number,
  w: number,
  h: number,
  cssWidthPx: number,
): Box {
  return {
    x: py,
    y: cssWidthPx - (px + w),
    w: h,
    h: w,
  };
}

/**
 * CSS placement (top-left + transform) for one vertical-page text-selection
 * overlay span (ECMA-376 §17.6.20 tbRl). The renderer emits each run's geometry
 * in the SWAPPED LOGICAL frame (`onTextRun` reports logical `x`/`y`/`w`/`h`); the
 * canvas is the PHYSICAL landscape page rotated +90° at paint. A DOM overlay span
 * is horizontal text, so to land it on the drawn (rotated) glyphs we place it at
 * the physical point the logical top-left maps to and rotate it +90° about that
 * corner — matching the page transform.
 *
 * Page transform: `physical = (cssWidth − logical.y, logical.x)`. The logical
 * run top-left `(x, y)` therefore lands at physical `(cssWidth − y, x)`. Applying
 * `transform: rotate(90deg)` with `transform-origin: top left` sends the span's
 * own advance axis (+x local) to physical +y (down the column) and its line-box
 * thickness (+y local) to physical −x — exactly the drawn run's footprint, so the
 * transparent span overlays the glyphs and native selection/search hit-tests land
 * correctly. (CJK glyphs are drawn upright while the span text is rotated sideways;
 * the span still covers the same cell rectangle, which is what selection needs.
 * Per-glyph upright overlay spans are a follow-up.)
 *
 * @returns `{ left, top }` in physical CSS px and the `transform` string, or
 *          `null` when `!vertical` (horizontal pages place the span at `(x, y)`
 *          untransformed — the byte-identical legacy path).
 */
export function verticalTextLayerPlacement(
  x: number,
  y: number,
  cssWidthPx: number,
  vertical: boolean,
): { left: number; top: number; transform: string } | null {
  if (!vertical) return null;
  return { left: cssWidthPx - y, top: x, transform: 'rotate(90deg)' };
}
