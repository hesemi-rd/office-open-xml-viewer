// ECMA-376 §17.6.20 vertical writing (`<w:textDirection w:val="tbRl">`) — the
// glyph-level primitives for rendering a page that has been laid out in the
// SWAPPED logical coordinate space and rotated +90° into physical space by the
// renderer's page transform (see `renderDocumentToCanvas`).
//
// After the page transform, a normal `ctx.fillText(text, x, baseline)` paints
// the run flowing DOWNWARD in physical space (logical +x → physical +y), which
// is exactly the character progression a vertical line wants — but every glyph
// is lying on its right side (rotated +90° CW with the page). For vertical
// Japanese:
//   • CJK ideographs, kana and CJK punctuation must stand UPRIGHT. We draw each
//     such glyph with a local −90° counter-rotation about its own centre, which
//     cancels the page rotation so it appears upright while still advancing down
//     the line.
//   • Latin letters and Western digits stay SIDEWAYS (rotated with the page):
//     that is the conventional "縦中横 not applied" appearance for runs of Latin
//     text set in vertical Japanese, so they need no counter-rotation and are
//     drawn as an ordinary contextual `fillText` (which also preserves the
//     browser's shaping/advance for the run).
//
// This module owns ONLY the pure geometry + classification; the renderer wires
// it into the whole-run glyph draw sites, the anchor/inline/float image draws,
// and the text-selection overlay behind the `verticalCJK` flag, so the
// horizontal path stays byte-identical.
//
// STAGE-1 SCOPE (issue #771 tracks stage-2). Implemented: +90° page rotation,
// upright-CJK / sideways-Latin glyph draw, anchor images resolved against the
// physical page then projected into the logical flow (PDF-verified centroid),
// inline/anchored/float image uprighting, and the vertical text-layer transform.
// NOT yet implemented (approximated or deferred): vertical-form punctuation
// substitution (U+FE10–U+FE19) and the `0.12em`/`0.4em` glyph nudges are
// font-dependent stage-1 heuristics (flagged inline); 縦中横 (tate-chū-yoko),
// `btLr` flow, header/footer + tables in tbRl, and paragraph-relative vertical
// anchors are follow-ups.

import { isCjkBreakChar } from '@silurus/ooxml-core';

/**
 * True when `cp` renders UPRIGHT in vertical Japanese (UTR#50 Vertical_
 * Orientation ≈ "U"/"Tu"): CJK ideographs, kana, Hangul, and CJK/fullwidth
 * punctuation. Latin letters, Western digits and other rotate-in-vertical
 * ("R") code points return false so the renderer leaves them sideways.
 *
 * Stage-1 scope (§ JIS X 4051): we reuse core's CJK block predicate
 * ({@link isCjkBreakChar}), which already covers the upright set — CJK
 * Symbols & Punctuation, Hiragana, Katakana, CJK Unified Ideographs (incl.
 * Ext-A), Hangul Syllables, CJK Compatibility Ideographs, and the Halfwidth/
 * Fullwidth Forms block — while excluding ASCII Latin/digits, which rotate.
 *
 * @param cp A Unicode scalar value (e.g. from `String.prototype.codePointAt`).
 */
export function isUprightVerticalGlyph(cp: number): boolean {
  return isCjkBreakChar(cp);
}

/** A punctuation glyph that is DRAWN at a shifted position in vertical text
 *  (§ JIS X 4051): the small ideographic comma/period sit in the upper-right
 *  cell corner rather than the lower-left, and the corner brackets / parens are
 *  mirrored to their vertical forms. Stage-1 handles the most visible ones by a
 *  draw-time offset (see {@link verticalGlyphOffset}); full vertical-form glyph
 *  substitution (U+FE10–U+FE19) is a follow-up. */
const VERTICAL_PUNCT_UPPER_RIGHT = new Set<number>([
  0x3001, // 、 ideographic comma
  0x3002, // 。 ideographic full stop
  0xff0c, // ， fullwidth comma
  0xff0e, // ． fullwidth full stop
]);

/**
 * Per-glyph draw offset (in em fractions of the font size) applied in the
 * glyph's own UPRIGHT local frame — i.e. after the −90° counter-rotation, in
 * physical (dx = rightward, dy = downward) terms. Returns `{ dx, dy }` em
 * fractions; the caller multiplies by the font px size.
 *
 * The small ideographic comma / full stop occupy the upper-right of their cell
 * in vertical text (unlike horizontal, where they sit at the lower-left). We
 * approximate that by nudging them up and to the right within the cell. Every
 * other glyph returns `{0,0}` (centred in its cell).
 */
export function verticalGlyphOffset(cp: number): { dx: number; dy: number } {
  if (VERTICAL_PUNCT_UPPER_RIGHT.has(cp)) {
    // HEURISTIC (stage-1 approximation, font-dependent): move the small comma/
    // full stop toward the upper-right corner of the cell by ~0.4em each way.
    // This is NOT a spec constant — JIS X 4051 §4.x gives the punctuation cell
    // geometry (the glyph occupies a quarter-em corner box), not a 0.4em nudge.
    // The correct fix is full vertical-form glyph substitution (U+FE10–U+FE19,
    // Unicode CJK Compatibility Forms) so the font supplies the pre-positioned
    // vertical comma/period; then this offset is deleted. Tracked in issue #771
    // (vertical-text stage-2).
    return { dx: 0.4, dy: -0.4 };
  }
  return { dx: 0, dy: 0 };
}

/**
 * Split a run's text into maximal runs of same-orientation code points, so the
 * vertical draw path can counter-rotate the UPRIGHT (CJK) segments per glyph
 * while drawing SIDEWAYS (Latin/digit) segments as a single contextual
 * `fillText`. Preserves surrogate pairs (iterates by code point) and returns
 * the pieces in logical order with each piece's starting code-point index.
 *
 * @param text The run's text.
 * @returns Ordered pieces, each `{ text, upright }`.
 */
export function splitVerticalOrientationRuns(
  text: string,
): Array<{ text: string; upright: boolean }> {
  const pieces: Array<{ text: string; upright: boolean }> = [];
  let cur = '';
  let curUpright: boolean | null = null;
  for (const ch of text) {
    const cp = ch.codePointAt(0) ?? 0;
    const upright = isUprightVerticalGlyph(cp);
    if (curUpright === null) {
      curUpright = upright;
      cur = ch;
    } else if (upright === curUpright) {
      cur += ch;
    } else {
      pieces.push({ text: cur, upright: curUpright });
      cur = ch;
      curUpright = upright;
    }
  }
  if (cur !== '' && curUpright !== null) {
    pieces.push({ text: cur, upright: curUpright });
  }
  return pieces;
}

type Ctx2D = CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D;

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
 * @param letterSpacingPx  Per-glyph extra advance (docGrid cell delta or
 *                         justification pitch); 0 for the common path.
 */
export function drawVerticalRun(
  ctx: Ctx2D,
  text: string,
  x: number,
  baseline: number,
  fontPx: number,
  letterSpacingPx: number,
): void {
  const prevAlign = ctx.textAlign;
  const prevBaseline = ctx.textBaseline;
  let ax = 0; // cumulative advance from run left (logical +x)
  for (const ch of text) {
    const cp = ch.codePointAt(0) ?? 0;
    const upright = isUprightVerticalGlyph(cp);
    const adv = ctx.measureText(ch).width + letterSpacingPx;
    if (upright) {
      // Cell centre in logical space. Counter-rotate −90° about it so the glyph
      // (which the page rotation would otherwise lay on its side) stands upright.
      const cx = x + ax + adv / 2;
      const off = verticalGlyphOffset(cp);
      ctx.save();
      ctx.translate(cx, baseline);
      ctx.rotate(-Math.PI / 2);
      // In the upright local frame: draw centred horizontally (the cell centre)
      // with a `middle` baseline that centres the glyph box on the cell.
      // HEURISTIC (stage-1 approximation, font-dependent): the extra +0.12em
      // vertical shift nudges typical CJK metrics so the visual centre lands on
      // the cell centre — it is NOT a spec constant (JIS X 4051 positions the
      // glyph from the font's ideographic em-box / vertical baseline, which the
      // browser does not expose here). The correct fix is to place each glyph
      // from the font's vertical metrics (ideographic baseline / `ideographic`
      // textBaseline once broadly supported); replace 0.12em then. Tracked in
      // issue #771 (vertical-text stage-2).
      ctx.textAlign = 'center';
      ctx.textBaseline = 'middle';
      ctx.fillText(
        ch,
        off.dx * fontPx,
        (0.12 + off.dy) * fontPx,
      );
      ctx.restore();
    } else {
      // Sideways (Latin/digit): draw as-is (rotated with the page). Keep the
      // caller's alphabetic baseline; position the glyph's left at the current
      // advance. Group of consecutive sideways glyphs would ideally be one
      // fillText for shaping, but per-glyph keeps the advance model uniform and
      // Latin advances are context-free enough at these sizes.
      ctx.textAlign = prevAlign;
      ctx.textBaseline = prevBaseline;
      ctx.fillText(ch, x + ax, baseline);
    }
    ax += adv;
  }
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
