import { describe, it, expect } from 'vitest';
import {
  layoutLines,
  rescaleLayoutLines,
  charSpacingDeltaPx,
  segAdvanceWidth,
  type LayoutSeg,
  type LayoutLine,
  type LayoutTextSeg,
} from './line-layout.js';

// ─────────────────────────────────────────────────────────────────────────────
// ECMA-376 §17.3.2.35 `<w:spacing>` — run character-spacing pitch must be part of
// a segment's laid-out ADVANCE at the PAINT scale, not only at the scale-1 stamp.
//
// `layoutLines` (the scale-1 paginator stamp) derives each segment's measuredWidth
// from `segAdvanceWidth`, which adds `cpCount × (w:spacing pt × scale)` on top of
// the natural `measureText` width. The paint pass, however, does NOT re-run
// `layoutLines` at the device scale — it reuses the scale-1 stamp through
// `rescaleLayoutLines`, whose docstring promises the result is "identical to a
// fresh paint-scale layout of the SAME partition" so that measure == paint.
//
// This pins the observed sample-34 defect ("氏名又は名称", where the run bearing
// `<w:spacing w:val="96"/>` = 4.8 pt is followed by another run): at the paint
// scale the following run's pen x collapses to the FIRST run's NATURAL width
// (charSpacing dropped), so the next run is painted on top of the expanded glyphs
// of the previous run. Because the draw path DOES apply `ctx.letterSpacing`
// (= charSpacing) inside the run, the run visually widens but the neighbour never
// gets pushed right → overlap.
//
// The width model itself (`segAdvanceWidth`, `charSpacingDeltaPx`) is already
// correct and unit-tested (run-char-metrics.test.ts). This file instead exercises
// the PAINT-SCALE reuse path (`rescaleLayoutLines`) end-to-end, which is where the
// advance loses the spacing.
// ─────────────────────────────────────────────────────────────────────────────

/** A recording 2D-context stub whose glyph advance is EXACTLY `0.5 · px · n`
 *  (linear in the font px size). Font-metric ascent/descent are the fixed 0.8/0.2
 *  em ratios the renderer's fallback uses. Being scale-linear, it isolates the
 *  advance MODEL (charSpacing) from font-hinting noise — any scale non-linearity
 *  in the assertions therefore comes from the layout code, not the font. */
function makeLinearCtx(perPx = 0.5): CanvasRenderingContext2D {
  let font = '10px serif';
  const pxOf = (): number => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
  const ctx = {
    get font() { return font; },
    set font(v: string) { font = v; },
    letterSpacing: '0px',
    fontKerning: 'auto',
    measureText: (s: string) => {
      const p = pxOf();
      const per = p * perPx;
      return {
        width: [...s].length * per,
        fontBoundingBoxAscent: p * 0.8, fontBoundingBoxDescent: p * 0.2,
        actualBoundingBoxAscent: p * 0.8, actualBoundingBoxDescent: p * 0.2,
      } as TextMetrics;
    },
  };
  return ctx as unknown as CanvasRenderingContext2D;
}

function textSeg(text: string, extra: Partial<LayoutTextSeg> = {}): LayoutSeg {
  return {
    text, bold: false, italic: false, underline: false, strikethrough: false,
    fontSize: 12, color: null, fontFamily: 'Times New Roman', vertAlign: null,
    measuredWidth: 0, ...extra,
  } as LayoutSeg;
}

function cloneSegs(segs: LayoutSeg[]): LayoutSeg[] {
  return segs.map((s) => ({ ...s }));
}

/** sample-34 "氏名又は名称" partition: run1 (4 CJK cps) + run2 (1 cp) both carry
 *  `<w:spacing>` = 4.8 pt; run3 carries none. All fit on one line here so the
 *  segment partition (3 text segments) is stable and we can read each segment's
 *  advance directly. No docGrid (`gridDeltaPx = 0`) so the ONLY per-glyph delta in
 *  play is §17.3.2.35 character spacing. */
const SPACING_PT = 4.8;
function makeSegs(): LayoutSeg[] {
  return [
    textSeg('氏名又は', { charSpacing: SPACING_PT }),
    textSeg('名', { charSpacing: SPACING_PT }),
    textSeg('称'),
  ];
}
const WIDE = 1000; // pt — wide enough that the paragraph never wraps

function isText(s: LayoutSeg): s is LayoutTextSeg {
  return 'text' in s && !('isTab' in s) && !('imagePath' in s) && !('mathNodes' in s);
}

describe('§17.3.2.35 run charSpacing survives into the paint-scale advance', () => {
  it('the scale-1 stamp already folds charSpacing into the advance (baseline)', () => {
    const ctx = makeLinearCtx();
    const stamp = layoutLines(ctx, makeSegs(), WIDE, 0, 1);
    expect(stamp).toHaveLength(1);
    const segs = stamp[0].segments.filter(isText);
    expect(segs).toHaveLength(3);
    // run1: natural (4 cps × 6px) + 4 × (4.8 pt × scale 1). measure==paint holds
    // at the paginator scale, so this is not the regression — it anchors the
    // expected value the paint scale must reproduce.
    const natural1 = 4 * (12 * 0.5); // 24 px
    expect(segs[0].measuredWidth).toBeCloseTo(natural1 + 4 * charSpacingDeltaPx(segs[0], 1), 6);
  });

  it('rescaleLayoutLines keeps charSpacing in the advance at the paint scale', () => {
    const scale = 2;
    const ctx = makeLinearCtx();
    const stamp = layoutLines(ctx, makeSegs(), WIDE, 0, 1);
    const painted = rescaleLayoutLines(stamp, scale, ctx, {}, 0);

    expect(painted).toHaveLength(1);
    const segs = painted[0].segments.filter(isText);
    expect(segs).toHaveLength(3);

    // run1 "氏名又は": the paint-scale advance MUST include the §17.3.2.35 pitch.
    //   natural(scale 2) = 4 cps × (12·2 × 0.5) = 48 px
    //   + 4 × (4.8 pt × scale 2 = 9.6 px)      = 38.4 px
    //   = 86.4 px
    // Before the fix the reuse path re-derived the advance from the natural
    // width + docGrid delta only (the deleted `gridWidth` helper) → 48 px,
    // dropping the 38.4 px of spacing, so this assertion failed.
    const natural2 = 4 * (12 * scale * 0.5); // 48 px
    const expected1 = natural2 + 4 * charSpacingDeltaPx(segs[0], scale); // 86.4 px
    expect(segs[0].measuredWidth).toBeCloseTo(expected1, 6);

    // run2 "名" (1 cp): natural (12 px) + 1 × 9.6 = 21.6 px.
    const natural2b = 1 * (12 * scale * 0.5); // 12 px
    expect(segs[1].measuredWidth).toBeCloseTo(
      natural2b + 1 * charSpacingDeltaPx(segs[1], scale), 6);

    // run3 "称" (no spacing): natural only, unchanged by the fix.
    expect(segs[2].measuredWidth).toBeCloseTo(1 * (12 * scale * 0.5), 6);
  });

  it('the following run does not overlap the first run at the paint scale', () => {
    // The paint pen advances by each segment's measuredWidth; run2 therefore starts
    // at run1.measuredWidth. run1 is drawn with ctx.letterSpacing = charSpacing, so
    // its painted right edge is `segAdvanceWidth(run1)`. The two must coincide (the
    // pen must land where the glyphs end) or run2 is painted over run1's tail.
    const scale = 2;
    const ctx = makeLinearCtx();
    const stamp = layoutLines(ctx, makeSegs(), WIDE, 0, 1);
    const painted = rescaleLayoutLines(stamp, scale, ctx, {}, 0);
    const segs = painted[0].segments.filter(isText);

    const run2StartX = segs[0].measuredWidth; // run1 is first on the line
    const run1PaintedRightEdge = segAdvanceWidth(segs[0], 4 * (12 * scale * 0.5), 0, scale);
    // No overlap: run2 begins at or after run1's painted right edge.
    expect(run2StartX).toBeGreaterThanOrEqual(run1PaintedRightEdge - 1e-6);
  });

  it('§17.3.2.43 w:w (charScale) also survives into the paint-scale advance', () => {
    // The same reuse path dropped `w:w` too (it multiplies the natural width, so
    // a 200%-stretched run collapsed back to its natural advance at the paint
    // scale). Pin it alongside charSpacing: the consolidation must route BOTH
    // run character metrics through the one advance authority.
    const scale = 2;
    const ctx = makeLinearCtx();
    const stamp = layoutLines(ctx, [
      textSeg('wide', { charScale: 2 }),
      textSeg('after'),
    ], WIDE, 0, 1);
    const painted = rescaleLayoutLines(stamp, scale, ctx, {}, 0);
    const segs = painted[0].segments.filter(isText);
    expect(segs).toHaveLength(2);
    // "wide" = 4 cps: natural(scale 2) = 4 × (12·2 × 0.5) = 48 px, × w:w 2 = 96 px.
    expect(segs[0].measuredWidth).toBeCloseTo(96, 6);
    // "after" (no w:w, 5 cps): natural only = 5 × 12 = 60 px.
    expect(segs[1].measuredWidth).toBeCloseTo(60, 6);
  });

  it('the reuse path matches a fresh paint-scale layout of the same partition', () => {
    // rescaleLayoutLines' contract: "identical to a fresh paint-scale layout of the
    // SAME partition". A direct layoutLines call at the paint scale uses
    // segAdvanceWidth (charSpacing included), so every segment advance must agree.
    const scale = 2;
    const stampCtx = makeLinearCtx();
    const stamp = layoutLines(stampCtx, makeSegs(), WIDE, 0, 1);
    const reused = rescaleLayoutLines(stamp, scale, makeLinearCtx(), {}, 0);

    const freshCtx = makeLinearCtx();
    const fresh = layoutLines(freshCtx, cloneSegs(makeSegs()), WIDE * scale, 0, scale);

    const reusedSegs = reused[0].segments.filter(isText);
    const freshSegs = (fresh[0] as LayoutLine).segments.filter(isText);
    expect(reusedSegs).toHaveLength(freshSegs.length);
    for (let i = 0; i < reusedSegs.length; i++) {
      expect(reusedSegs[i].measuredWidth).toBeCloseTo(freshSegs[i].measuredWidth, 6);
    }
  });
});
