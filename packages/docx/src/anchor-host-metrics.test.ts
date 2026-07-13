import { DEFAULT_KINSOKU_RULES } from '@silurus/ooxml-core';
import { describe, expect, it } from 'vitest';
import {
  buildSegments,
  layoutLines,
  lineBoxHeight,
  type LayoutSeg,
  type LayoutTextSeg,
} from './line-layout.js';
import type { DocRun } from './types.js';

function linearCtx(): CanvasRenderingContext2D {
  let font = '10px serif';
  const px = (): number => Number.parseFloat(/([\d.]+)px/.exec(font)?.[1] ?? '10');
  return {
    get font() { return font; },
    set font(value: string) { font = value; },
    letterSpacing: '0px',
    measureText: (text: string) => ({
      width: [...text].length * px() * 0.5,
      fontBoundingBoxAscent: px() * 0.8,
      fontBoundingBoxDescent: px() * 0.2,
      actualBoundingBoxAscent: px() * 0.8,
      actualBoundingBoxDescent: px() * 0.2,
    } as TextMetrics),
  } as unknown as CanvasRenderingContext2D;
}

function anchoredShape(): DocRun {
  return {
    type: 'shape',
    widthPt: 100,
    heightPt: 1,
    anchorXPt: 0,
    anchorYPt: 0,
    anchorXFromMargin: false,
    anchorYFromPara: true,
    zOrder: 0,
    subpaths: [],
    presetGeometry: 'line',
    fill: null,
    stroke: '000000',
    anchorHostMetrics: {
      fontSize: 20,
      fontFamily: 'Arial',
      fontFamilyEastAsia: 'Yu Mincho',
      bold: false,
      italic: false,
    },
  } as DocRun;
}

describe('floating drawing anchor-host metrics', () => {
  it('emits a zero-width metric segment using the anchor character formatting', () => {
    const segments = buildSegments([anchoredShape()], { pageIndex: 0, totalPages: 1 });

    expect(segments).toHaveLength(1);
    expect(segments[0]).toMatchObject({
      text: '',
      measuredWidth: 0,
      fontSize: 20,
      fontFamily: 'Yu Mincho',
      metricOnly: true,
      metricEastAsian: true,
    });
  });

  it('uses the zero-width metric segment to allocate East Asian document-grid cells', () => {
    const segments = buildSegments([anchoredShape()], { pageIndex: 0, totalPages: 1 });
    const [line] = layoutLines(
      linearCtx(),
      segments.map((segment) => ({ ...segment })) as LayoutSeg[],
      400,
      0,
      1,
      undefined,
      undefined,
      {},
      0,
      DEFAULT_KINSOKU_RULES,
      0,
      36,
      400,
      false,
      false,
      false,
    );

    expect((line.segments[0] as LayoutTextSeg).measuredWidth).toBe(0);
    expect(line.eastAsian).toBe(true);
    expect(lineBoxHeight(
      null,
      line.ascent,
      line.descent,
      1,
      { type: 'lines', linePitchPt: 18 },
      false,
      line.intendedSingle,
      line.eastAsian,
      line.gridCountSingle,
    )).toBe(36);
  });
});
