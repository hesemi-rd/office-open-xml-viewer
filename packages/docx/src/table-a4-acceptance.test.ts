import { beforeAll, describe, expect, it } from 'vitest';
import { layoutDocument } from './document-layout.js';
import type { TableLayout } from './layout/types.js';
import type { TextLayoutService, TextShapeRequest } from './layout/text.js';
import { paintTableLayout } from './paint/canvas-table.js';
import type {
  CanvasPaintContext,
  CanvasPaintResourcePainter,
  PaintCanvas2D,
} from './paint/types.js';
import { createLayoutServices } from './renderer.js';
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

function measuringContext(): CanvasRenderingContext2D {
  let font = '10px serif';
  return {
    get font() { return font; },
    set font(value: string) { font = value; },
    letterSpacing: '0px',
    fontKerning: 'auto',
    measureText(text: string) {
      const fontSize = Number.parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
      return {
        width: [...text].length * fontSize * 0.5,
        fontBoundingBoxAscent: fontSize * 0.8,
        fontBoundingBoxDescent: fontSize * 0.2,
        actualBoundingBoxAscent: fontSize * 0.8,
        actualBoundingBoxDescent: fontSize * 0.2,
      } as TextMetrics;
    },
    save() {}, restore() {}, beginPath() {}, closePath() {}, moveTo() {}, lineTo() {},
    stroke() {}, fill() {}, fillRect() {}, strokeRect() {}, clip() {}, rect() {},
    scale() {}, translate() {}, rotate() {}, setLineDash() {}, clearRect() {}, arc() {},
    quadraticCurveTo() {}, bezierCurveTo() {}, createLinearGradient() {
      return { addColorStop() {} };
    },
    drawImage() {}, fillText() {}, strokeText() {},
    fillStyle: '#000000', strokeStyle: '#000000', lineWidth: 1,
    textAlign: 'left' as CanvasTextAlign, direction: 'ltr' as CanvasDirection,
  } as unknown as CanvasRenderingContext2D;
}

beforeAll(() => {
  (globalThis as unknown as { OffscreenCanvas: unknown }).OffscreenCanvas = class {
    getContext() { return measuringContext(); }
  };
});

const noBorders = {
  top: null, right: null, bottom: null, left: null, insideH: null, insideV: null,
};

function paragraph(text: string): DocParagraph {
  return {
    type: 'paragraph', alignment: 'left', indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null, tabStops: [],
    runs: [{
      type: 'text', text, bold: false, italic: false, underline: false,
      strikethrough: false, fontSize: 10, color: null,
      fontFamily: 'Times New Roman', fontFamilyEastAsia: '', isLink: false,
      background: null, vertAlign: null, hyperlink: null,
    } as DocParagraph['runs'][number]],
    defaultFontSize: 10, defaultFontFamily: 'Times New Roman', widowControl: false,
  } as unknown as DocParagraph;
}

function textCell(text: string, overrides: Partial<DocTableCell> = {}): DocTableCell {
  return {
    content: [{ type: 'paragraph', ...paragraph(text) } as CellElement],
    colSpan: 1, vMerge: null, borders: noBorders, background: null,
    vAlign: 'top', widthPt: null, ...overrides,
  } as unknown as DocTableCell;
}

function row(cells: DocTableCell[], overrides: Partial<DocTableRow> = {}): DocTableRow {
  return {
    cells, rowHeight: null, rowHeightRule: 'auto', isHeader: false, ...overrides,
  } as DocTableRow;
}

function privateCellWire(marginTwips: number): Record<string, unknown> {
  return {
    preferredWidth: null,
    margins: {
      top: { kind: 'dxa', value: `${marginTwips}` },
      right: { kind: 'dxa', value: `${marginTwips}` },
      bottom: { kind: 'dxa', value: `${marginTwips}` },
      left: { kind: 'dxa', value: `${marginTwips}` },
      start: null, end: null,
    },
  };
}

function table(
  layout: 'fixed' | 'autofit',
  texts: Readonly<{ merged: string; side: string; nested: string }>,
  nestedLayout: 'fixed' | 'autofit',
): DocTable {
  const nestedCell = {
    ...textCell(texts.nested),
    __tableCellLayout: privateCellWire(20),
  } as unknown as DocTableCell;
  const nested = {
    type: 'table',
    colWidths: [32],
    rows: [row([nestedCell])],
    borders: noBorders,
    cellMarginTop: 0, cellMarginRight: 0, cellMarginBottom: 0, cellMarginLeft: 0,
    jc: 'center', layout: nestedLayout,
    __tableLayout: {
      effectiveStyleId: 'NestedSyntheticStyle',
      grid: { authored: true, columns: [{ width: '640' }], requiredColumnCount: 1 },
      preferredWidth: null, layout: { kind: nestedLayout },
      cellSpacing: null, cellMargins: null,
    },
  } as unknown as DocTable;

  const restart = {
    ...textCell(texts.merged, {
      vMerge: true,
      borders: { ...noBorders, right: { width: 2, color: 'ff0000', style: 'thick' } },
    }),
    __tableCellLayout: privateCellWire(20),
  } as unknown as DocTableCell;
  const nestedOwner = {
    content: [{ type: 'table', ...nested } as CellElement],
    colSpan: 1, vMerge: null,
    borders: { ...noBorders, left: { width: 1, color: '0000ff', style: 'single' } },
    background: null, vAlign: 'top', widthPt: null,
    __tableCellLayout: privateCellWire(20),
  } as unknown as DocTableCell;
  const continuation = {
    content: [], colSpan: 1, vMerge: false, borders: noBorders,
    background: null, vAlign: 'top', widthPt: null,
    __tableCellLayout: privateCellWire(20),
  } as unknown as DocTableCell;
  const side = {
    ...textCell(texts.side),
    __tableCellLayout: privateCellWire(20),
  } as unknown as DocTableCell;

  const exact = {
    ...row([restart, nestedOwner], { rowHeight: 12, rowHeightRule: 'exact' }),
    __tableRowLayout: {
      height: { value: '240', rule: 'exact', ruleAuthored: true },
      justification: null, beforeWidth: null, afterWidth: null,
      cellSpacing: null,
      exception: {
        preferredWidth: null, layout: null, justification: null, indent: null,
        borders: null, cellMargins: null, cellSpacing: null,
      },
    },
  } as unknown as DocTableRow;
  const atLeast = {
    ...row([continuation, side], { rowHeight: 10, rowHeightRule: 'auto' }),
    __tableRowLayout: {
      // Missing hRule is a Word atLeast compatibility input ([MS-OI29500] 2.1.180).
      height: { value: '200', rule: 'auto', ruleAuthored: false },
      justification: null, beforeWidth: null, afterWidth: null,
      cellSpacing: { kind: 'dxa', value: '20' }, exception: null,
    },
  } as unknown as DocTableRow;

  return {
    type: 'table', colWidths: [50, 50], rows: [exact, atLeast], borders: noBorders,
    cellMarginTop: 0, cellMarginRight: 0, cellMarginBottom: 0, cellMarginLeft: 0,
    jc: 'left', layout,
    __tableLayout: {
      effectiveStyleId: 'OuterSyntheticStyle',
      grid: {
        authored: true,
        columns: [{ width: '1000' }, { width: '1000' }],
        requiredColumnCount: 2,
      },
      preferredWidth: null, layout: { kind: layout },
      cellSpacing: null, cellMargins: null,
    },
  } as unknown as DocTable;
}

function document(body: BodyElement[]): DocxDocumentModel {
  const section = {
    pageWidth: 240, pageHeight: 400,
    marginTop: 10, marginRight: 10, marginBottom: 10, marginLeft: 10,
    headerDistance: 4, footerDistance: 4, titlePage: false, evenAndOddHeaders: false,
    sectionStart: 'nextPage', columns: null,
  } as SectionProps;
  return {
    section, body,
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    fontFamilyClasses: { 'Times New Roman': 'roman' }, footnotes: [],
  } as unknown as DocxDocumentModel;
}

function paintContext(): CanvasPaintContext {
  const ctx = {
    globalAlpha: 1, fillStyle: '', strokeStyle: '', lineWidth: 1, font: '',
    textAlign: 'left' as CanvasTextAlign,
    textBaseline: 'alphabetic' as CanvasTextBaseline,
    direction: 'ltr' as CanvasDirection, letterSpacing: '0px',
    fontKerning: 'auto' as CanvasFontKerning,
    save() {}, restore() {}, beginPath() {}, rect() {}, clip() {}, translate() {},
    rotate() {}, scale() {}, fillRect() {}, strokeRect() {}, setLineDash() {},
    moveTo() {}, lineTo() {}, stroke() {}, fill() {}, drawImage() {}, fillText() {},
  } as unknown as PaintCanvas2D;
  const resources: CanvasPaintResourcePainter = {
    paint(resourceKey, kind): never {
      throw new Error(`Unexpected ${kind} resource: ${resourceKey}`);
    },
  };
  return { ctx, scale: 1, dpr: 1, resources };
}

function retainedTable(model: DocxDocumentModel): Readonly<{
  layout: TableLayout;
  finalAcquisitions: ReadonlyMap<string, number>;
  intrinsicProbes: ReadonlyMap<string, number>;
}> {
  const base = createLayoutServices(model, { measureContext: measuringContext() });
  const finalAcquisitions = new Map<string, number>();
  const intrinsicProbes = new Map<string, number>();
  const text: TextLayoutService = Object.freeze({
    fingerprint: base.text.fingerprint,
    localMetrics: base.text.localMetrics,
    resolve: (request: Parameters<TextLayoutService['resolve']>[0]) => base.text.resolve(request),
    shape(request: TextShapeRequest) {
      const target = request.clusterGeometry === true ? finalAcquisitions : intrinsicProbes;
      target.set(request.text, (target.get(request.text) ?? 0) + 1);
      return base.text.shape(request);
    },
  });
  const result = layoutDocument(model, Object.freeze({ ...base, text }));
  const placed = result.pages.flatMap((page) => page.fragments)
    .find((item) => item.fragment.kind === 'table');
  if (!placed || placed.fragment.kind !== 'table' || !('flowBounds' in placed.fragment)) {
    throw new Error('Expected one retained ordinary table');
  }
  return { layout: placed.fragment, finalAcquisitions, intrinsicProbes };
}

describe('A4 retained table acceptance', () => {
  it.each([
    ['fixed', 'autofit'],
    ['autofit', 'fixed'],
  ] as const)(
    'acquires final paragraph geometry once for %s outer and %s nested tables',
    (outerLayout, nestedLayout) => {
      // Keep each marker below the narrowest fixed cell width so final line
      // shaping retains one whole-string key; wrapped piece acquisition is a
      // separate paragraph concern covered by the A3 tests.
      const texts = outerLayout === 'fixed'
        ? { merged: 'mfa', side: 'sfa', nested: 'nfa' }
        : { merged: 'maf', side: 'saf', nested: 'naf' };
      const model = document([table(outerLayout, texts, nestedLayout) as BodyElement]);
      const before = structuredClone(model);
      const acquired = retainedTable(model);

      expect(Object.fromEntries(
        Object.values(texts).map((text) => [text, acquired.finalAcquisitions.get(text)]),
      )).toEqual({
        [texts.merged]: 1,
        [texts.side]: 1,
        [texts.nested]: 1,
      });
      expect([...acquired.intrinsicProbes.values()].reduce((sum, count) => sum + count, 0))
        .toBeGreaterThan(0);
      // Exact keeps the authored 12pt track plus Word's 1pt bottom-padding
      // addition; omitted hRule remains an atLeast floor and may grow to text.
      expect(acquired.layout.rows[0]?.advancePt).toBe(13);
      expect(acquired.layout.rows[1]?.advancePt).toBeGreaterThanOrEqual(10);
      expect(acquired.layout.rows[0]?.cells[0]?.verticalMerge).toBe('restart');
      expect(acquired.layout.rows[1]?.cells[0]?.verticalMerge).toBe('continue');
      expect(acquired.layout.rows[0]?.cells[1]?.blocks[0]?.layout.kind).toBe('table');
      expect(acquired.layout.borders).toContainEqual(expect.objectContaining({
        edge: 'between', color: '#ff0000', widthPt: 2,
      }));
      expect(model).toEqual(before);

      const countsBeforePaint = {
        final: [...acquired.finalAcquisitions.entries()],
        intrinsic: [...acquired.intrinsicProbes.entries()],
      };
      const retainedBeforePaint = structuredClone(acquired.layout);
      paintTableLayout(acquired.layout, paintContext());

      expect({
        final: [...acquired.finalAcquisitions.entries()],
        intrinsic: [...acquired.intrinsicProbes.entries()],
      }).toEqual(countsBeforePaint);
      expect(acquired.layout).toEqual(retainedBeforePaint);
      expect(model).toEqual(before);
    },
  );
});
