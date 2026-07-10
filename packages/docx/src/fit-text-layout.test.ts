import { describe, expect, it } from 'vitest';
import {
  buildSegments,
  layoutLines,
  rescaleLayoutLines,
  type LayoutSeg,
  type LayoutTextSeg,
} from './line-layout.js';
import type { DocRun, DocxTextRun } from './types.js';

function textRun(text: string, extra: Partial<DocxTextRun> = {}): DocRun {
  return {
    type: 'text',
    text,
    bold: false,
    italic: false,
    underline: false,
    strikethrough: false,
    fontSize: 12,
    color: null,
    fontFamily: 'serif',
    isLink: false,
    background: null,
    vertAlign: null,
    allCaps: false,
    smallCaps: false,
    doubleStrikethrough: false,
    ...extra,
  } as unknown as DocRun;
}

function makeLinearCtx(): CanvasRenderingContext2D {
  let font = '12px serif';
  const ctx = {
    get font() { return font; },
    set font(value: string) { font = value; },
    letterSpacing: '0px',
    fontKerning: 'auto',
    measureText(text: string) {
      const px = Number(/([\d.]+)px/.exec(font)?.[1] ?? 12);
      return {
        width: [...text].length * px,
        fontBoundingBoxAscent: px * 0.8,
        fontBoundingBoxDescent: px * 0.2,
        actualBoundingBoxAscent: px * 0.8,
        actualBoundingBoxDescent: px * 0.2,
      } as TextMetrics;
    },
  };
  return ctx as unknown as CanvasRenderingContext2D;
}

function textSegments(segs: LayoutSeg[]): LayoutTextSeg[] {
  return segs.filter((seg): seg is LayoutTextSeg => 'text' in seg);
}

const ID = -1431456512;
const ENV = { pageIndex: 0, totalPages: 1 };

function fitRuns(prefix = false): DocRun[] {
  const fit = { fitTextVal: 2400, fitTextId: ID, charSpacing: 4.8 };
  return [
    ...(prefix ? [textRun('X')] : []),
    textRun('氏名又は', fit),
    textRun('名', fit),
    textRun('称', fit),
  ];
}

describe('ECMA-376 §17.3.2.14 fitText layout integration', () => {
  it('assigns one stable run-level region across script-split segments', () => {
    const segs = textSegments(buildSegments([
      textRun('AB氏', { fitTextVal: 2400, fitTextId: ID }),
      textRun('名', { fitTextVal: 2400, fitTextId: ID }),
    ], ENV));

    expect(segs).toHaveLength(3); // latin + eastAsia from run 0, eastAsia from run 1
    expect(new Set(segs.map((seg) => seg.fitTextRegionIndex)).size).toBe(1);
    expect(segs[0].fitTextRunIndex).toBe(segs[1].fitTextRunIndex);
    expect(segs[2].fitTextRunIndex).not.toBe(segs[0].fitTextRunIndex);
    expect(segs.map((seg) => seg.joinPrev)).toEqual([undefined, true, true]);
  });

  it('folds one cross-run per-gap into the canonical segment advances', () => {
    const lines = layoutLines(makeLinearCtx(), buildSegments(fitRuns(), ENV), 1000, 0, 1);
    const segs = textSegments(lines[0].segments);

    expect(lines).toHaveLength(1);
    expect(segs).toHaveLength(3);
    expect(segs.map((seg) => seg.fitTextPerGapPx)).toEqual([9.6, 9.6, 9.6]);
    expect(segs[0].measuredWidth).toBeCloseTo(86.4, 9);
    expect(segs[1].measuredWidth).toBeCloseTo(21.6, 9);
    expect(segs[2].measuredWidth).toBeCloseTo(12, 9);
    expect(segs.reduce((sum, seg) => sum + seg.measuredWidth, 0)).toBeCloseTo(120, 9);
  });

  it('moves the fixed-width region as one atomic unit', () => {
    const lines = layoutLines(makeLinearCtx(), buildSegments(fitRuns(true), ENV), 125, 0, 1);

    expect(lines).toHaveLength(2);
    expect(textSegments(lines[0].segments).map((seg) => seg.text)).toEqual(['X']);
    expect(textSegments(lines[1].segments).map((seg) => seg.text)).toEqual(['氏名又は', '名', '称']);
    expect(textSegments(lines[1].segments).reduce((sum, seg) => sum + seg.measuredWidth, 0))
      .toBeCloseTo(120, 9);
  });

  it('recomputes the scale-relative per-gap for stamp reuse', () => {
    const stamp = layoutLines(makeLinearCtx(), buildSegments(fitRuns(), ENV), 1000, 0, 1);
    const painted = rescaleLayoutLines(stamp, 2, makeLinearCtx(), {}, 0);
    const segs = textSegments(painted[0].segments);

    expect(segs[0].fitTextPerGapPx).toBeCloseTo(19.2, 9);
    expect(segs.reduce((sum, seg) => sum + seg.measuredWidth, 0)).toBeCloseTo(240, 9);
  });

  it('recomputes a single-glyph region cell at paint scale', () => {
    const stamp = layoutLines(
      makeLinearCtx(),
      buildSegments([textRun('氏', { fitTextVal: 2400 })], ENV),
      1000,
      0,
      1,
    );
    const painted = rescaleLayoutLines(stamp, 2, makeLinearCtx(), {}, 0);

    expect(textSegments(stamp[0].segments)[0].measuredWidth).toBeCloseTo(120, 9);
    expect(textSegments(painted[0].segments)[0].measuredWidth).toBeCloseTo(240, 9);
  });
});
