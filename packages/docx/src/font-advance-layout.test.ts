import { describe, expect, it } from 'vitest';
import {
  buildSegments,
  layoutLines,
  segGlyphScaleFactor,
  type LayoutSeg,
  type LayoutTextSeg,
} from './line-layout.js';
import type { DocRun, DocxTextRun } from './types.js';

const ENV = { pageIndex: 0, totalPages: 1 };

type AdvanceStampedSeg = LayoutTextSeg & {
  fontAdvanceScaleCandidate?: number;
  fontAdvanceScale?: number;
};

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
    fontFamily: 'Meiryo UI',
    fontFamilyEastAsia: 'Meiryo UI',
    isLink: false,
    background: null,
    vertAlign: null,
    allCaps: false,
    smallCaps: false,
    doubleStrikethrough: false,
    ...extra,
  } as unknown as DocRun;
}

function measuringCtx(kanaEm: number): CanvasRenderingContext2D {
  let font = '12px serif';
  const ctx = {
    get font() { return font; },
    set font(value: string) { font = value; },
    letterSpacing: '0px',
    fontKerning: 'auto',
    measureText(text: string) {
      const px = Number(/([\d.]+)px/.exec(font)?.[1] ?? 12);
      return {
        width: [...text].length * px * kanaEm,
        fontBoundingBoxAscent: px * 0.8,
        fontBoundingBoxDescent: px * 0.2,
        actualBoundingBoxAscent: px * 0.8,
        actualBoundingBoxDescent: px * 0.2,
      } as TextMetrics;
    },
  };
  return ctx as unknown as CanvasRenderingContext2D;
}

function textSegments(segs: LayoutSeg[]): AdvanceStampedSeg[] {
  return segs.filter((seg): seg is AdvanceStampedSeg => 'text' in seg);
}

describe('requested-font horizontal advance substitution probe', () => {
  it('keeps the profile only for a full-width substitute, not an installed condensed face', () => {
    const condensed = textSegments(buildSegments([textRun('ひら')], ENV));
    expect(condensed).toHaveLength(1);
    expect(condensed[0].fontAdvanceScaleCandidate).toBeCloseTo(0.7775, 10);
    const condensedLines = layoutLines(measuringCtx(0.78), condensed, 1000, 0, 1);
    const condensedSeg = textSegments(condensedLines[0].segments)[0];
    expect(condensedSeg.fontAdvanceScale).toBe(1);
    expect(condensedSeg.measuredWidth).toBeCloseTo(2 * 12 * 0.78, 9);

    // Use the localized alias so the module-level resolution cache has a distinct
    // built-font key while exercising the same harvested profile.
    const substitute = textSegments(buildSegments([
      textRun('ひら', { fontFamily: 'メイリオ UI', fontFamilyEastAsia: 'メイリオ UI' }),
    ], ENV));
    const substituteLines = layoutLines(measuringCtx(1), substitute, 1000, 0, 1);
    const substituteSeg = textSegments(substituteLines[0].segments)[0];
    expect(substituteSeg.fontAdvanceScale).toBeCloseTo(0.7775, 10);
    expect(substituteSeg.measuredWidth).toBeCloseTo(2 * 12 * 0.7775, 9);
  });

  it('leaves vertical CJK kana at full-em advances', () => {
    const segs = textSegments(buildSegments([textRun('ひら')], { ...ENV, verticalCJK: true }));
    expect(segs).toHaveLength(1);
    expect(segs[0].fontAdvanceScaleCandidate).toBeUndefined();

    const lines = layoutLines(measuringCtx(1), segs, 1000, 0, 1);
    const laidOut = textSegments(lines[0].segments)[0];
    expect(segGlyphScaleFactor(laidOut)).toBe(1);
    expect(laidOut.measuredWidth).toBe(24);
  });

  it('keeps a mixed-script ruby base in one unscaled segment', () => {
    const segs = textSegments(buildSegments([
      textRun('漢かな', { ruby: { text: 'かん', fontSizePt: 6 } }),
    ], ENV));

    expect(segs).toHaveLength(1);
    expect(segs[0].text).toBe('漢かな');
    expect(segs[0].ruby?.text).toBe('かん');
    expect(segs[0].fontAdvanceScaleCandidate).toBeUndefined();
    expect(segGlyphScaleFactor(segs[0])).toBe(1);
  });

  it('uses the resolved glyph scale when fitting a Meiryo UI region to w:fitText', () => {
    const segs = buildSegments([
      textRun('ひら', {
        fontFamily: 'メイリオ UI',
        fontFamilyEastAsia: 'メイリオ UI',
        fitTextVal: 2400,
      }),
    ], ENV);
    const lines = layoutLines(measuringCtx(1), segs, 1000, 0, 1);
    const fitted = textSegments(lines[0].segments);

    expect(fitted.reduce((sum, seg) => sum + seg.measuredWidth, 0)).toBeCloseTo(120, 9);
  });
});
