import { beforeAll, describe, expect, it } from 'vitest';
import { DEFAULT_KINSOKU_RULES } from '@silurus/ooxml-core';
import { layoutDocument } from './document-layout.js';
import { createLayoutServices, resolveColumnWidths } from './renderer.js';
import { measureParagraphIntrinsicWidths } from './layout/intrinsic-width.js';
import type { ParagraphLayoutContext } from './layout-context.js';
import type { TextLayoutService } from './layout/text.js';
import type {
  BodyElement,
  CellElement,
  DocParagraph,
  DocTable,
  DocTableCell,
  DocTableRow,
  DocxDocumentModel,
  SectionProps,
} from './types.js';

type ColumnState = Parameters<typeof resolveColumnWidths>[2];

function measuringContext(
  widthOf: (text: string) => number = (text) => [...text].length * 5,
): CanvasRenderingContext2D {
  let font = '10px serif';
  return {
    get font() { return font; },
    set font(value: string) { font = value; },
    letterSpacing: '0px',
    fontKerning: 'auto',
    measureText(text: string) {
      return {
        width: widthOf(text),
        fontBoundingBoxAscent: 8,
        fontBoundingBoxDescent: 2,
        actualBoundingBoxAscent: 8,
        actualBoundingBoxDescent: 2,
      } as TextMetrics;
    },
    save() {}, restore() {}, beginPath() {}, closePath() {}, moveTo() {}, lineTo() {},
    stroke() {}, fill() {}, fillRect() {}, strokeRect() {}, clip() {}, rect() {},
    scale() {}, translate() {}, rotate() {}, setLineDash() {}, clearRect() {}, arc() {},
    quadraticCurveTo() {}, bezierCurveTo() {}, drawImage() {}, fillText() {}, strokeText() {},
    createLinearGradient() { return { addColorStop() {} }; },
    fillStyle: '#000000', strokeStyle: '#000000', lineWidth: 1,
    textAlign: 'left', direction: 'ltr',
  } as unknown as CanvasRenderingContext2D;
}

beforeAll(() => {
  (globalThis as unknown as { OffscreenCanvas: unknown }).OffscreenCanvas = class {
    private readonly context = measuringContext();
    getContext() { return this.context; }
  };
});

const borders = {
  top: null, right: null, bottom: null, left: null, insideH: null, insideV: null,
};

function textRun(text: string): DocParagraph['runs'][number] {
  return {
    type: 'text', text, fontSize: 10, fontFamily: 'serif',
    bold: false, italic: false, underline: false, strikethrough: false,
    color: null, isLink: false, background: null, vertAlign: null, hyperlink: null,
  } as DocParagraph['runs'][number];
}

function paragraph(
  runs: DocParagraph['runs'],
  overrides: Partial<DocParagraph> = {},
): DocParagraph {
  return {
    type: 'paragraph', alignment: 'left', indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null, tabStops: [],
    widowControl: false, runs, defaultFontSize: 10, defaultFontFamily: 'serif',
    ...overrides,
  } as unknown as DocParagraph;
}

function cell(content: CellElement[]): DocTableCell {
  return {
    content, colSpan: 1, vMerge: null, borders, background: null, vAlign: 'top',
    widthPt: null, widthPct: null,
    marginTop: 0, marginRight: 0, marginBottom: 0, marginLeft: 0,
  } as unknown as DocTableCell;
}

function row(cells: DocTableCell[]): DocTableRow {
  return { cells, rowHeight: null, rowHeightRule: 'auto', isHeader: false } as DocTableRow;
}

function table(
  rows: DocTableRow[],
  colWidths: number[],
  layout?: 'fixed' | 'autofit',
): DocTable {
  return {
    type: 'table', rows, colWidths, borders,
    cellMarginTop: 0, cellMarginRight: 0, cellMarginBottom: 0, cellMarginLeft: 0,
    jc: 'left', ...(layout ? { layout } : {}),
  } as unknown as DocTable;
}

function model(body: BodyElement[]): DocxDocumentModel {
  const section = {
    pageWidth: 220, pageHeight: 300,
    marginTop: 10, marginRight: 10, marginBottom: 10, marginLeft: 10,
    headerDistance: 4, footerDistance: 4, titlePage: false, evenAndOddHeaders: false,
    sectionStart: 'nextPage', columns: null,
  } as SectionProps;
  return {
    section, body,
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    fontFamilyClasses: {}, footnotes: [],
  } as unknown as DocxDocumentModel;
}

function columnState(
  ctx: CanvasRenderingContext2D,
  services = createLayoutServices(model([]), { measureContext: ctx }),
): ColumnState {
  return {
    ctx, fontFamilyClasses: {}, layoutServices: services,
    pageWidth: 200, pageH: 300, scale: 1, pageIndex: 0, totalPages: 1,
    defaultTabPt: 36,
  } as unknown as ColumnState;
}

describe('table intrinsic content widths', () => {
  const intrinsicContext = (overrides: Partial<ParagraphLayoutContext> = {}): ParagraphLayoutContext => ({
    lineGrid: { active: false, pitchPt: null },
    characterGrid: { active: false, deltaPt: 0 },
    physicalIndentLeftPt: 0,
    physicalIndentRightPt: 0,
    firstIndentPt: 0,
    lineSpacing: null,
    spaceBeforePt: 0,
    spaceAfterPt: 0,
    baseRtl: false,
    isJustified: false,
    stretchLastLine: false,
    tabStops: [],
    hasRuby: false,
    hasEastAsianText: true,
    kinsoku: DEFAULT_KINSOKU_RULES,
    defaultTabPt: 36,
    ...overrides,
  });

  it('includes the character-grid pitch in minimum content width', () => {
    const source = paragraph([textRun('漢字')]);

    expect(measureParagraphIntrinsicWidths(
      source,
      intrinsicContext({ characterGrid: { active: true, deltaPt: 2 } }),
      200,
      { context: measuringContext(), fontFamilyClasses: {} },
      { pageIndex: 0, totalPages: 1, documentHasEastAsianText: true },
    )).toEqual({ minWidthPt: 7, maxWidthPt: 14 });
  });

  it('does not merge runs with different character-grid participation', () => {
    const source = paragraph([
      { ...textRun('漢'), snapToGrid: true },
      { ...textRun('字'), snapToGrid: false },
    ] as DocParagraph['runs']);

    expect(measureParagraphIntrinsicWidths(
      source,
      intrinsicContext({ characterGrid: { active: true, deltaPt: 2 } }),
      200,
      { context: measuringContext(), fontFamilyClasses: {} },
      { pageIndex: 0, totalPages: 1, documentHasEastAsianText: true },
    )).toEqual({ minWidthPt: 7, maxWidthPt: 12 });
  });

  it('keeps adjacent tate-chu-yoko runs as separate one-em cells', () => {
    const source = paragraph([
      { ...textRun('12'), eastAsianVert: true },
      { ...textRun('34'), eastAsianVert: true },
    ] as DocParagraph['runs']);

    expect(measureParagraphIntrinsicWidths(
      source,
      intrinsicContext(),
      200,
      { context: measuringContext(), fontFamilyClasses: {} },
      { pageIndex: 0, totalPages: 1, documentHasEastAsianText: true, verticalCJK: true },
    )).toEqual({ minWidthPt: 20, maxWidthPt: 20 });
  });

  it('keeps an inline image as a minimum-content atom', () => {
    const source = table([row([cell([paragraph([{
      type: 'image', imagePath: 'word/media/image.png', mimeType: 'image/png',
      widthPt: 80, heightPt: 10, anchor: false,
    }]) as CellElement])])], [0]);

    expect(resolveColumnWidths(source, 200, columnState(measuringContext()))).toEqual([80]);
  });

  it('uses structural math metadata for minimum and maximum content width', () => {
    const resourceKey = 'math:body:0.0.0:inline';
    const run = {
      type: 'math', display: false, fontSize: 10, resourceKey,
      nodes: [{ kind: 'run', text: 'x', style: 'italic' }],
    } as unknown as DocParagraph['runs'][number];
    const source = table([row([cell([paragraph([run]) as CellElement])])], [0]);
    const ctx = measuringContext();
    const base = createLayoutServices(model([]), { measureContext: ctx });
    const services = Object.freeze({
      ...base,
      math: Object.freeze({
        fingerprint: 'table-intrinsic-math',
        resolve: (key: string) => ({
          resourceKey: key, widthEm: 5, ascentEm: 0.8, descentEm: 0.2, diagnostics: [],
        }),
      }),
    });

    expect(resolveColumnWidths(source, 200, columnState(ctx, services))).toEqual([50]);
  });

  it('retains a left-tab leader in minimum content width', () => {
    const source = table([row([cell([paragraph([textRun('\tX')], {
      tabStops: [{ pos: 60, alignment: 'left', leader: 'dot' }],
    }) as CellElement])])], [0]);

    expect(resolveColumnWidths(source, 200, columnState(measuringContext()))).toEqual([60]);
  });

  it('uses following content when resolving a right-tab maximum', () => {
    const tabbed = cell([paragraph([textRun('\tX')], {
      tabStops: [{ pos: 60, alignment: 'right', leader: 'none' }],
    }) as CellElement]);
    const source = table([row([tabbed, cell([])])], [0, 100]);

    expect(resolveColumnWidths(source, 100, columnState(measuringContext()))).toEqual([60, 40]);
  });

  it('includes retained numbering-marker geometry and paragraph indents', () => {
    const numbered = paragraph([textRun('X')], {
      indentLeft: 18,
      indentFirst: -9,
      numbering: {
        numId: 1, level: 0, format: 'decimal', text: '12345.',
        indentLeft: 18, tab: 18, suff: 'tab', jc: 'left',
      } as NonNullable<DocParagraph['numbering']>,
    });
    const source = table([row([cell([numbered as CellElement])])], [0]);

    // The six-glyph marker overruns the 9pt hanging area, so the suffix tab
    // advances the body to the next 36pt default stop: 18 + 54 + 5 = 77pt.
    expect(resolveColumnWidths(source, 200, columnState(measuringContext()))).toEqual([77]);
  });

  it('uses a fixed nested table as an outer minimum-content atom', () => {
    const nested = table([row([cell([])])], [80], 'fixed');
    const source = table([row([cell([nested as CellElement])])], [0]);

    expect(resolveColumnWidths(source, 200, columnState(measuringContext()))).toEqual([80]);
  });

  it('shapes identical formatting across a run seam as one proportional atom', () => {
    const ctx = measuringContext((text) => text === 'AV' ? 15 : [...text].length * 10);
    const source = table([row([cell([
      paragraph([textRun('A'), textRun('V')]) as CellElement,
    ])])], [0]);

    expect(resolveColumnWidths(source, 200, columnState(ctx))).toEqual([15]);
  });

  it('does not acquire content widths for fixed-layout columns', () => {
    const measured: string[] = [];
    const ctx = measuringContext((text) => {
      measured.push(text);
      return [...text].length * 5;
    });
    const source = table([row([cell([
      paragraph([textRun('fixed-only')]) as CellElement,
    ])])], [80], 'fixed');

    expect(resolveColumnWidths(source, 200, columnState(ctx))).toEqual([80]);
    expect(measured).toEqual([]);
  });
});

describe('table retained marker acquisition', () => {
  it('acquires numbering marker glyph geometry once', () => {
    const marker = '12345.';
    const numbered = paragraph([textRun('X')], {
      indentLeft: 18,
      indentFirst: -9,
      numbering: {
        numId: 1, level: 0, format: 'decimal', text: marker,
        indentLeft: 18, tab: 18, suff: 'tab', jc: 'left',
      } as NonNullable<DocParagraph['numbering']>,
    });
    const source = table([row([cell([numbered as CellElement])])], [100], 'fixed');
    const document = model([source as BodyElement]);
    const base = createLayoutServices(document, { measureContext: measuringContext() });
    let markerShapes = 0;
    const text: TextLayoutService = Object.freeze({
      fingerprint: base.text.fingerprint,
      localMetrics: base.text.localMetrics,
      resolve: (request: Parameters<TextLayoutService['resolve']>[0]) => base.text.resolve(request),
      shape: (request: Parameters<TextLayoutService['shape']>[0]) => {
        if (request.text === marker && request.clusterGeometry !== false) markerShapes += 1;
        return base.text.shape(request);
      },
    });

    layoutDocument(document, Object.freeze({ ...base, text }));

    expect(markerShapes).toBe(1);
  });
});
