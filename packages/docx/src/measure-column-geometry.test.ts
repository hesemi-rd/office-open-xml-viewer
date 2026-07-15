import { describe, it, expect } from 'vitest';
import { computeFrameBox, registerFrameFloat } from './frame-geometry.js';
import {
  computeFloatTableBox,
  registerTableFloat,
  floatTableWrapSide,
} from './float-table-geometry.js';
import type { FrameBox } from './frame-geometry.js';
import type { FloatTableBox } from './float-table-geometry.js';
import type { FramePr, TblpPr } from './types.js';
import type { FloatRect } from './float-layout.js';

// #513 regression: the change-page MEASURE pass (computePages) must resolve an
// out-of-flow frame (§17.3.1.11 framePr) / floating table (§17.4.57 tblpPr) with
// the SAME per-column content band the PAINT pass (renderBodyElements) uses, so
// their FloatRect x / wrap side agree inside a multi-column (§17.6.4) section.
//
// The paint pass sets state.contentX/contentW = col.xPt × scale / col.wPt × scale
// per element BEFORE resolving the float (renderer.ts ~2410). The measure pass
// (measureState.scale = 1) now re-points measureState.contentX/contentW to
// colX()/colW() via the withColumnBand helper. Both feed frameXContainer
// (hAnchor/horzAnchor="text") and floatTableWrapSide, which read contentX/contentW
// directly. These tests pin that:
//   (1) measure (column band, scale 1, pt) == paint (column band, scale 1, px)
//       for box.x and the registered FloatRect x-range and wrap side; and
//   (2) the OLD measure model (full band) would have diverged — guarding the fix.
//
// A 2-column page modeled at scale=1 so px == pt:
//   pageWidth=600, margins L/R/T/B = 60/60/72/72 ⇒ content band [60,540] (w=480).
//   Two equal columns with a 24-wide gutter:
//     col 0: x=60,  w=228   col 1: x=312, w=228.
// The frame / table is anchored in COLUMN 1 (the right column), so the measured
// box must land at col1.x — NOT at the full band's left (60) the old code used.

const PAGE = {
  scale: 1,
  marginLeft: 60,
  marginRight: 60,
  marginTop: 72,
  marginBottom: 72,
  pageWidth: 600,
  pageH: 800,
};

// Full content band (the OLD, buggy measure band) and the two columns.
const FULL_BAND = { contentX: 60, contentW: 480 };
const COL0 = { xPt: 60, wPt: 228 };
const COL1 = { xPt: 312, wPt: 228 };

interface MinState {
  scale: number;
  contentX: number;
  contentW: number;
  marginLeft: number;
  marginRight: number;
  marginTop: number;
  marginBottom: number;
  pageWidth: number;
  pageH: number;
  floats: FloatRect[];
  floatParaSeq: number;
}

function state(band: { contentX: number; contentW: number }): MinState {
  return {
    ...PAGE,
    contentX: band.contentX,
    contentW: band.contentW,
    floats: [],
    floatParaSeq: 0,
  };
}

// The paint pass sets contentX/contentW = col.xPt × scale / col.wPt × scale.
function paintBand(col: { xPt: number; wPt: number }, scale: number) {
  return { contentX: col.xPt * scale, contentW: col.wPt * scale };
}
// The measure pass (withColumnBand, scale=1) sets contentX/contentW = colX()/colW().
function measureBand(col: { xPt: number; wPt: number }) {
  return { contentX: col.xPt, contentW: col.wPt };
}

function frame(over: Partial<FramePr> = {}): FramePr {
  return {
    dropCap: 'none',
    lines: 1,
    wrap: 'around',
    hAnchor: 'text',
    vAnchor: 'text',
    hRule: 'auto',
    hSpace: 0,
    vSpace: 0,
    ...over,
  };
}

function tblp(over: Partial<TblpPr> = {}): TblpPr {
  return {
    leftFromText: 0,
    rightFromText: 0,
    topFromText: 0,
    bottomFromText: 0,
    horzAnchor: 'text',
    horzSpecified: true,
    vertAnchor: 'text',
    tblpX: 0,
    tblpY: 0,
    ...over,
  };
}

const fbox = (fp: FramePr, st: MinState, paraTop: number, cW: number, cH: number, aH: number): FrameBox =>
  computeFrameBox(fp, st as never, paraTop, cW, cH, aH);
const regFrame = (b: FrameBox, fp: FramePr, st: MinState): void =>
  registerFrameFloat(b, fp, st as never);
const tbox = (tp: TblpPr, st: MinState, paraTop: number, w: number, h: number): FloatTableBox =>
  computeFloatTableBox(tp, st as never, paraTop, w, h);
const regTable = (b: FloatTableBox, tp: TblpPr, st: MinState, side: string, ov: boolean): void =>
  registerTableFloat(b, tp, st as never, side, ov);
const wrapSide = (b: FloatTableBox, st: MinState): string =>
  floatTableWrapSide(b, st as never);

describe('measure/paint column-band parity for a horzAnchor="text" FRAME (§17.3.1.11, #513)', () => {
  // A generic frame at tblpX-like x offset 8 inside column 1.
  const fp = frame({ hAnchor: 'text', x: 8, hRule: 'exact', h: 40, w: 120 });
  const paraTop = 300;

  it('measure (scale=1 pt) box.x == paint (scale=1 px) box.x == column-1 left + offset', () => {
    const paint = fbox(fp, state(paintBand(COL1, 1)), paraTop, 120, 40, 14);
    const measure = fbox(fp, state(measureBand(COL1)), paraTop, 120, 40, 14);
    // Both anchor against column 1's left edge (312) + the 8 offset.
    expect(measure.x).toBe(312 + 8);
    expect(measure.x).toBe(paint.x);
  });

  it('the registered FloatRect x-range matches between measure and paint', () => {
    const stPaint = state(paintBand(COL1, 1));
    const bPaint = fbox(fp, stPaint, paraTop, 120, 40, 14);
    regFrame(bPaint, fp, stPaint);

    const stMeasure = state(measureBand(COL1));
    const bMeasure = fbox(fp, stMeasure, paraTop, 120, 40, 14);
    regFrame(bMeasure, fp, stMeasure);

    expect(stMeasure.floats).toHaveLength(1);
    expect(stPaint.floats).toHaveLength(1);
    expect(stMeasure.floats[0].xLeft).toBe(stPaint.floats[0].xLeft);
    expect(stMeasure.floats[0].xRight).toBe(stPaint.floats[0].xRight);
    // And it really is column-1-relative, not full-band-relative.
    expect(stMeasure.floats[0].xLeft).toBe(312 + 8);
  });

  it('the OLD full-band measure would have diverged (guards the fix)', () => {
    const oldMeasure = fbox(fp, state(FULL_BAND), paraTop, 120, 40, 14);
    // The pre-fix path anchored against the full band left (60) + offset.
    expect(oldMeasure.x).toBe(60 + 8);
    // …which is NOT where the paint pass draws it (column 1).
    const paint = fbox(fp, state(paintBand(COL1, 1)), paraTop, 120, 40, 14);
    expect(oldMeasure.x).not.toBe(paint.x);
  });
});

describe('measure/paint column-band parity for a tblpPr FLOATING TABLE (§17.4.57, #513)', () => {
  // A 120-wide table at horzAnchor="text", tblpX=10, anchored in column 1.
  const tp = tblp({ horzAnchor: 'text', vertAnchor: 'text', tblpX: 10, tblpY: 1 });
  const tableW = 120;
  const tableH = 50;
  const paraTop = 300;

  it('measure box.x == paint box.x == column-1 left + tblpX', () => {
    const paint = tbox(tp, state(paintBand(COL1, 1)), paraTop, tableW, tableH);
    const measure = tbox(tp, state(measureBand(COL1)), paraTop, tableW, tableH);
    expect(measure.x).toBe(312 + 10);
    expect(measure.x).toBe(paint.x);
  });

  it('wrap side is computed against the SAME column band ⇒ measure == paint', () => {
    // Column 1 band [312,540], centre 426. Table at x=322, right edge 442 straddles
    // the centre once measured against the column — both passes must agree.
    const measure = tbox(tp, state(measureBand(COL1)), paraTop, tableW, tableH);
    const paint = tbox(tp, state(paintBand(COL1, 1)), paraTop, tableW, tableH);
    const sideMeasure = wrapSide(measure, state(measureBand(COL1)));
    const sidePaint = wrapSide(paint, state(paintBand(COL1, 1)));
    expect(sideMeasure).toBe(sidePaint);

    // Side selection is no longer a midpoint heuristic. Both bands pass the
    // actual exclusion to the same widest-free-gap solver.
    const sideOld = wrapSide(measure, state(FULL_BAND));
    expect(sideOld).toBe(sideMeasure);
    expect(sideMeasure).toBe('bothSides');
  });

  it('the registered FloatRect x-range matches between measure and paint', () => {
    const stPaint = state(paintBand(COL1, 1));
    const bPaint = tbox(tp, stPaint, paraTop, tableW, tableH);
    const sidePaint = wrapSide(bPaint, stPaint);
    regTable(bPaint, tp, stPaint, sidePaint, true);

    const stMeasure = state(measureBand(COL1));
    const bMeasure = tbox(tp, stMeasure, paraTop, tableW, tableH);
    const sideMeasure = wrapSide(bMeasure, stMeasure);
    regTable(bMeasure, tp, stMeasure, sideMeasure, true);

    expect(stMeasure.floats).toHaveLength(1);
    expect(stPaint.floats).toHaveLength(1);
    expect(stMeasure.floats[0].xLeft).toBe(stPaint.floats[0].xLeft);
    expect(stMeasure.floats[0].xRight).toBe(stPaint.floats[0].xRight);
    expect(stMeasure.floats[0].side).toBe(stPaint.floats[0].side);
    expect(stMeasure.floats[0].xLeft).toBe(312 + 10);
  });
});

describe('single-column / page|margin anchor are UNCHANGED by the column band (#513 regression)', () => {
  // For a single full-width column the measure band == full band == paint band, and
  // page/margin anchors ignore contentX entirely — so nothing moves.
  it('single full-width column: measure band == full band (frame x unchanged)', () => {
    const fp = frame({ hAnchor: 'text', x: 8, hRule: 'exact', h: 40, w: 120 });
    const single = { xPt: 60, wPt: 480 }; // one full-width column
    const measure = fbox(fp, state(measureBand(single)), 300, 120, 40, 14);
    const full = fbox(fp, state(FULL_BAND), 300, 120, 40, 14);
    expect(measure.x).toBe(full.x);
    expect(measure.x).toBe(60 + 8);
  });

  it('horzAnchor="page" ignores the column band (same x in col 0 and col 1)', () => {
    const tp = tblp({ horzAnchor: 'page', vertAnchor: 'text', tblpX: 5 });
    const inCol0 = tbox(tp, state(measureBand(COL0)), 300, 120, 50);
    const inCol1 = tbox(tp, state(measureBand(COL1)), 300, 120, 50);
    expect(inCol0.x).toBe(5); // page left + tblpX, column-independent
    expect(inCol1.x).toBe(5);
  });

  it('horzAnchor="margin" ignores the column band (anchors at marginLeft)', () => {
    const fp = frame({ hAnchor: 'margin', x: 8, hRule: 'exact', h: 40, w: 120 });
    const inCol0 = fbox(fp, state(measureBand(COL0)), 300, 120, 40, 14);
    const inCol1 = fbox(fp, state(measureBand(COL1)), 300, 120, 40, 14);
    expect(inCol0.x).toBe(60 + 8); // marginLeft + x, column-independent
    expect(inCol1.x).toBe(60 + 8);
  });
});
