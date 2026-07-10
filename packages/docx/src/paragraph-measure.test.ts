import { describe, expect, it } from 'vitest';
import { DEFAULT_KINSOKU_RULES } from '@silurus/ooxml-core';
import type { FloatRect } from './float-layout.js';
import {
  createFloatWrapOracle,
  measureParagraph,
  type ParagraphMeasurementEnvironment,
  type ParagraphPlacement,
  type TextMeasurer,
  type WrapOracle,
} from './paragraph-measure.js';
import type { ParagraphLayoutContext } from './layout-context.js';
import type { DocParagraph, DocxTextRun, FieldRun, ImageRun } from './types.js';

function makeContext(ascentRatio = 0.8, descentRatio = 0.2): CanvasRenderingContext2D {
  let font = '10px serif';
  const fontSize = (): number => Number.parseFloat(/([\d.]+)px/.exec(font)?.[1] ?? '10');
  return {
    get font() { return font; },
    set font(value: string) { font = value; },
    letterSpacing: '0px',
    measureText: (text: string) => {
      const size = fontSize();
      return {
        width: [...text].length * size * 0.5,
        fontBoundingBoxAscent: size * ascentRatio,
        fontBoundingBoxDescent: size * descentRatio,
        actualBoundingBoxAscent: size * ascentRatio,
        actualBoundingBoxDescent: size * descentRatio,
      } as TextMetrics;
    },
  } as unknown as CanvasRenderingContext2D;
}

const measurer: TextMeasurer = {
  context: makeContext(),
  fontFamilyClasses: {},
};

const environment = (
  overrides: Partial<ParagraphMeasurementEnvironment> = {},
): ParagraphMeasurementEnvironment => ({
  pageIndex: 0,
  totalPages: 1,
  documentHasEastAsianText: false,
  ...overrides,
});

const paragraph = (overrides: Partial<DocParagraph> = {}): DocParagraph => ({
  alignment: 'left',
  indentLeft: 0,
  indentRight: 0,
  indentFirst: 0,
  spaceBefore: 3,
  spaceAfter: 4,
  lineSpacing: null,
  numbering: null,
  tabStops: [],
  runs: [],
  ...overrides,
});

const textRun = (text: string, overrides: Partial<DocxTextRun> = {}): DocxTextRun => ({
  text,
  bold: false,
  italic: false,
  underline: false,
  strikethrough: false,
  fontSize: 10,
  color: null,
  fontFamily: null,
  isLink: false,
  background: null,
  vertAlign: null,
  hyperlink: null,
  ...overrides,
});

const layoutContext = (
  overrides: Partial<ParagraphLayoutContext> = {},
): ParagraphLayoutContext => ({
  lineGrid: { active: false, pitchPt: null },
  characterGrid: { active: false, deltaPt: 0 },
  physicalIndentLeftPt: 0,
  physicalIndentRightPt: 0,
  firstIndentPt: 0,
  lineSpacing: null,
  spaceBeforePt: 3,
  spaceAfterPt: 4,
  baseRtl: false,
  tabStops: [],
  hasRuby: false,
  hasEastAsianText: false,
  kinsoku: DEFAULT_KINSOKU_RULES,
  defaultTabPt: 36,
  ...overrides,
});

const placement = (overrides: Partial<ParagraphPlacement> = {}): ParagraphPlacement => ({
  startYPt: 10,
  paragraphXPt: 0,
  availableWidthPt: 200,
  maximumYPt: 300,
  suppressSpaceBefore: false,
  ...overrides,
});

describe('measureParagraph', () => {
  it('measures a no-float paragraph and excludes trailing spacing from contentEndYPt', () => {
    const result = measureParagraph(
      paragraph({ runs: [{ type: 'text', ...textRun('hello') }] }),
      layoutContext(),
      placement(),
      measurer,
      environment(),
    );

    expect(result.markOnly).toBe(false);
    expect(result.lines).toHaveLength(1);
    expect(result.lines[0].topYPt).toBe(13);
    expect(result.lines[0].advancePt).toBe(10);
    expect(result.contentStartYPt).toBe(13);
    expect(result.contentEndYPt).toBe(23);
    expect(result.requestedSpaceBeforePt).toBe(3);
    expect(result.requestedSpaceAfterPt).toBe(4);
    expect(result.placement).toEqual(placement());
  });

  it('uses a float-backed wrap window for line placement and width', () => {
    const float: FloatRect = {
      kind: 'shape', mode: 'square', imageKey: 'float',
      imageX: 0, imageY: 10, imageW: 80, imageH: 30,
      xLeft: 0, xRight: 80, yTop: 10, yBottom: 40,
      side: 'bothSides', distLeft: 0, distRight: 0, distTop: 0, distBottom: 0,
      paraId: 1, drawn: false,
    };

    const result = measureParagraph(
      paragraph({ spaceBefore: 0, runs: [{ type: 'text', ...textRun('wrapped') }] }),
      layoutContext({ spaceBeforePt: 0 }),
      placement({ wrap: createFloatWrapOracle([float]) }),
      measurer,
      environment(),
    );

    expect(result.lines[0].topYPt).toBe(10);
    expect(result.lines[0].layout.xOffset).toBe(80);
    expect(result.lines[0].layout.availWidth).toBe(120);
  });

  it('reserves one paragraph-mark line for an empty paragraph', () => {
    const result = measureParagraph(
      paragraph(), layoutContext(), placement(), measurer, environment(),
    );

    expect(result.markOnly).toBe(true);
    expect(result.lines).toEqual([]);
    expect(result.contentStartYPt).toBe(13);
    expect(result.contentEndYPt).toBe(23);
  });

  it('uses resolved context line spacing for an empty paragraph mark', () => {
    const authoredSpacing = { value: 6, rule: 'exact' as const, explicit: true };
    const resolvedSpacing = { value: 18, rule: 'exact' as const, explicit: true };
    const result = measureParagraph(
      paragraph({ spaceBefore: 0, lineSpacing: authoredSpacing }),
      layoutContext({ spaceBeforePt: 0, lineSpacing: resolvedSpacing }),
      placement({ startYPt: 0 }),
      measurer,
      environment(),
    );

    expect(result.markOnly).toBe(true);
    expect(result.contentEndYPt).toBe(18);
  });

  it('uses document-level East Asian metrics for an empty paragraph mark', () => {
    const tallMeasurer: TextMeasurer = {
      context: makeContext(0.9, 0.2),
      fontFamilyClasses: {},
    };
    const eastAsianMetrics = measureParagraph(
      paragraph({ defaultFontSize: 20 }),
      layoutContext({
        lineGrid: { active: true, pitchPt: 20 },
        spaceBeforePt: 0,
      }),
      placement({ startYPt: 0 }),
      tallMeasurer,
      environment({ documentHasEastAsianText: true }),
    );
    const paragraphOnlyMetrics = measureParagraph(
      paragraph({ defaultFontSize: 20 }),
      layoutContext({
        lineGrid: { active: true, pitchPt: 20 },
        spaceBeforePt: 0,
      }),
      placement({ startYPt: 0 }),
      tallMeasurer,
      environment({ documentHasEastAsianText: false }),
    );

    expect(eastAsianMetrics.contentEndYPt).toBe(40);
    expect(paragraphOnlyMetrics.contentEndYPt).toBe(22);
  });

  it('treats an anchor-only paragraph as a paragraph mark', () => {
    const anchor: ImageRun = {
      imagePath: 'word/media/anchor.png', mimeType: 'image/png',
      widthPt: 40, heightPt: 30, anchor: true,
    };
    const result = measureParagraph(
      paragraph({ runs: [{ type: 'image', ...anchor }] }),
      layoutContext(), placement(), measurer, environment(),
    );

    expect(result.markOnly).toBe(true);
    expect(result.lines).toEqual([]);
    expect(result.contentEndYPt).toBe(23);
  });

  it('uses one uniform snapped advance for every line in a ruby paragraph', () => {
    const result = measureParagraph(
      paragraph({
        spaceBefore: 0,
        runs: [{ type: 'text', ...textRun('aa aa', { ruby: { text: 'ruby', fontSizePt: 8 } }) }],
      }),
      layoutContext({
        lineGrid: { active: true, pitchPt: 10 },
        spaceBeforePt: 0,
        hasRuby: true,
      }),
      placement({ availableWidthPt: 12 }),
      measurer,
      environment(),
    );

    expect(result.lines.length).toBeGreaterThan(1);
    expect(new Set(result.lines.map((line) => line.advancePt))).toEqual(new Set([30]));
  });

  it('passes bidi policy through to RTL tab layout', () => {
    const result = measureParagraph(
      paragraph({
        spaceBefore: 0,
        tabStops: [{ pos: 50, alignment: 'left', leader: 'none' }],
        runs: [{ type: 'text', ...textRun('A\tB') }],
        bidi: true,
      }),
      layoutContext({
        spaceBeforePt: 0,
        baseRtl: true,
        tabStops: [{ pos: 50, alignment: 'left', leader: 'none' }],
      }),
      placement({ availableWidthPt: 100 }),
      measurer,
      environment(),
    );

    const tab = result.lines[0].layout.segments.find((segment) => 'isTab' in segment);
    expect(tab?.measuredWidth).toBe(45);
  });

  it('includes an inline image in line height', () => {
    const image: ImageRun = {
      imagePath: 'word/media/inline.png', mimeType: 'image/png',
      widthPt: 20, heightPt: 24, anchor: false,
    };
    const result = measureParagraph(
      paragraph({ spaceBefore: 0, runs: [{ type: 'image', ...image }] }),
      layoutContext({ spaceBeforePt: 0 }), placement(), measurer, environment(),
    );

    expect(result.markOnly).toBe(false);
    expect(result.lines[0].advancePt).toBe(24);
  });

  it('preserves exact line spacing verbatim', () => {
    const exact = { value: 18, rule: 'exact' as const, explicit: true };
    const result = measureParagraph(
      paragraph({ spaceBefore: 0, lineSpacing: exact, runs: [{ type: 'text', ...textRun('exact') }] }),
      layoutContext({ spaceBeforePt: 0, lineSpacing: exact }),
      placement(), measurer, environment(),
    );

    expect(result.lines[0].advancePt).toBe(18);
  });

  it('resolves fields from the explicit line-layout environment', () => {
    const field: FieldRun = {
      fieldType: 'page', instruction: 'PAGE', fallbackText: '1',
      bold: false, italic: false, underline: false, strikethrough: false,
      fontSize: 10, color: null, fontFamily: null, background: null,
      vertAlign: null,
    };
    const result = measureParagraph(
      paragraph({ spaceBefore: 0, runs: [{ type: 'field', ...field }] }),
      layoutContext({ spaceBeforePt: 0 }), placement(),
      measurer,
      environment({ pageIndex: 8, totalPages: 12, displayPageNumber: 42 }),
    );

    expect(result.lines[0].layout.segments[0]).toMatchObject({ text: '42' });
  });

  it('remeasures at a changed start Y and records the exact placement', () => {
    const wrap: WrapOracle = {
      lineWindow: ({ topYPt, maximumWidthPt }) => topYPt < 50
        ? { topYPt, xOffsetPt: 10, maximumWidthPt: 20 }
        : { topYPt, xOffsetPt: 0, maximumWidthPt },
      skipTopAndBottomBands: ({ yPt }) => yPt,
    };
    const doc = paragraph({
      spaceBefore: 0,
      runs: [{ type: 'text', ...textRun('abcdefghijklmnopqrst') }],
    });
    const context = layoutContext({ spaceBeforePt: 0 });
    const firstPlacement = placement({ startYPt: 10, availableWidthPt: 100, wrap });
    const secondPlacement = placement({ startYPt: 60, availableWidthPt: 100, wrap });

    const first = measureParagraph(doc, context, firstPlacement, measurer, environment());
    const second = measureParagraph(doc, context, secondPlacement, measurer, environment());

    expect(first).not.toBe(second);
    expect(first.placement).toEqual(firstPlacement);
    expect(second.placement).toEqual(secondPlacement);
    expect(first.lines[0].layout.availWidth).toBe(20);
    expect(second.lines[0].layout.availWidth).toBe(100);
    expect(first.lines.length).toBeGreaterThan(second.lines.length);
  });
});
