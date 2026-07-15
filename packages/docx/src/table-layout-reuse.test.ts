import { beforeAll, describe, expect, it } from 'vitest';
import {
  bodyFragmentFor,
  createLayoutServices,
  paginateDocument,
  renderDocumentToCanvas,
} from './renderer.js';
import type {
  BodyElement,
  DocParagraph,
  DocTable,
  DocTableCell,
  DocTableRow,
  DocxDocumentModel,
  PaginatedBodyElement,
  SectionProps,
} from './types';
import type { TableFragmentLayout } from './layout/table-pagination.js';

// A5 P4: table geometry is retained by TableLayout/TableFragmentLayout. The old
// paginator stamps and the reuse on/off switch were implementation details, so
// this suite now checks the retained tree that is shared by pagination and paint.

function emptyBorders() {
  return { top: null, bottom: null, left: null, right: null, insideH: null, insideV: null };
}

function makeCtx(onMeasure: () => void = () => {}): CanvasRenderingContext2D {
  let font = '10px serif';
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
  const ctx = {
    get font() { return font; },
    set font(value: string) { font = value; },
    letterSpacing: '0px',
    measureText: (text: string) => {
      onMeasure();
      const size = px();
      return {
        width: [...text].length * size * 0.5,
        fontBoundingBoxAscent: size * 0.8,
        fontBoundingBoxDescent: size * 0.2,
        actualBoundingBoxAscent: size * 0.8,
        actualBoundingBoxDescent: size * 0.2,
      } as TextMetrics;
    },
    save() {}, restore() {}, beginPath() {}, closePath() {},
    moveTo() {}, lineTo() {}, stroke() {}, fill() {}, fillRect() {}, strokeRect() {},
    rect() {}, clip() {}, scale() {}, translate() {}, rotate() {},
    setLineDash() {}, clearRect() {}, arc() {}, quadraticCurveTo() {}, bezierCurveTo() {},
    createLinearGradient() { return { addColorStop() {} }; },
    drawImage() {}, fillText() {}, strokeText() {},
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    textAlign: 'left' as CanvasTextAlign, direction: 'ltr' as CanvasDirection,
    globalAlpha: 1, lineCap: 'butt' as CanvasLineCap, lineJoin: 'miter' as CanvasLineJoin,
  };
  const canvas = { width: 0, height: 0, style: {}, getContext: () => ctx };
  (ctx as unknown as { canvas: unknown }).canvas = canvas;
  return ctx as unknown as CanvasRenderingContext2D;
}

beforeAll(() => {
  (globalThis as unknown as { OffscreenCanvas: unknown }).OffscreenCanvas = class {
    getContext() { return makeCtx(); }
  };
});

function paintCanvas(onMeasure: () => void): HTMLCanvasElement {
  const ctx = makeCtx(onMeasure);
  return (ctx.canvas as unknown) as HTMLCanvasElement;
}

function paragraph(text: string): DocParagraph {
  return {
    type: 'paragraph', alignment: 'left',
    indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null,
    numbering: null, tabStops: [],
    runs: [{
      type: 'text', text,
      bold: false, italic: false, underline: false, strikethrough: false,
      fontSize: 10, color: null, fontFamily: 'Times New Roman',
      fontFamilyEastAsia: '', isLink: false, background: null,
      vertAlign: null, hyperlink: null,
    } as DocParagraph['runs'][number]],
    defaultFontSize: 10, defaultFontFamily: 'Times New Roman', widowControl: false,
  } as unknown as DocParagraph;
}

function cell(text: string, widthPt: number): DocTableCell {
  return {
    content: [{ type: 'paragraph', ...paragraph(text) }],
    colSpan: 1, vMerge: null, borders: emptyBorders(), background: null,
    vAlign: 'top', widthPt,
  } as unknown as DocTableCell;
}

function row(index: number, isHeader = false): DocTableRow {
  return {
    cells: [cell(`left ${index}`, 200), cell(`right ${index}`, 300)],
    rowHeight: 20, rowHeightRule: 'exact', isHeader,
  } as unknown as DocTableRow;
}

function modelWithTable(rowCount: number, repeatedHeader = false): DocxDocumentModel {
  const rows = Array.from({ length: rowCount }, (_, index) => row(index, repeatedHeader && index === 0));
  const table = {
    type: 'table', colWidths: [200, 300], rows, borders: emptyBorders(),
    cellMarginTop: 0, cellMarginBottom: 0, cellMarginLeft: 0, cellMarginRight: 0,
    jc: 'left', layout: 'fixed',
  } as unknown as DocTable;
  return {
    section: {
      pageWidth: 500, pageHeight: 70,
      marginTop: 0, marginRight: 0, marginBottom: 0, marginLeft: 0,
      headerDistance: 0, footerDistance: 0, titlePage: false, evenAndOddHeaders: false,
    } as SectionProps,
    body: [table as unknown as BodyElement],
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    fontFamilyClasses: { 'Times New Roman': 'roman' },
  } as unknown as DocxDocumentModel;
}

function retainedTables(pages: PaginatedBodyElement[][]) {
  return pages.flatMap((page) => page
    .filter((element) => element.type === 'table')
    .map((element) => {
      const placed = bodyFragmentFor(element);
      expect(placed?.fragment.kind).toBe('table');
      if (placed?.fragment.kind !== 'table' || !('flowBounds' in placed.fragment)) {
        throw new Error('expected retained TableLayout/TableFragmentLayout');
      }
      return { element, placed, fragment: placed.fragment as TableFragmentLayout };
    }));
}

describe('retained table layout reuse', () => {
  it('stores shared column geometry and page-local row geometry without legacy stamps', () => {
    const model = modelWithTable(8);
    const pages = paginateDocument(model, createLayoutServices(model));
    const tables = retainedTables(pages);

    expect(tables.length).toBeGreaterThan(1);
    for (const { element, placed, fragment } of tables) {
      expect(fragment.columnWidthsPt).toEqual([200, 300]);
      expect(fragment.rows.length).toBeGreaterThan(0);
      expect(fragment.rows.every((tableRow) => tableRow.heightPt > 0)).toBe(true);
      expect(placed.heightPt).toBeCloseTo(fragment.advancePt, 6);
      expect(element).not.toHaveProperty('tableColWidthsPt');
      expect(element).not.toHaveProperty('tableRowHeightsPt');
      expect(element).not.toHaveProperty('tableLayoutInputs');
    }
    expect(tables.every(({ fragment }) => fragment.columnWidthsPt === tables[0]!.fragment.columnWidthsPt)).toBe(true);
  });

  it('represents repeated headers and source rows in TableFragmentLayout metadata', () => {
    const model = modelWithTable(8, true);
    const pages = paginateDocument(model, createLayoutServices(model));
    const tables = retainedTables(pages);

    expect(tables.length).toBeGreaterThan(1);
    const sourceRows = tables.flatMap(({ fragment }) => fragment.rows
      .filter((tableRow) => tableRow.ownership === 'source')
      .map((tableRow) => tableRow.logicalRowIndex));
    expect(sourceRows).toEqual(Array.from({ length: 8 }, (_, index) => index));
    for (const { fragment } of tables.slice(1)) {
      expect(fragment.rows[0]?.ownership).toBe('repeated-header');
      expect(fragment.rows[0]?.logicalRowIndex).toBe(0);
    }
  });

  it('paints the retained table tree at non-unit scale without measuring or mutating geometry', async () => {
    const model = modelWithTable(8, true);
    const pages = paginateDocument(model, createLayoutServices(model));
    const tables = retainedTables(pages);
    const fragments = tables.map(({ fragment }) => fragment);
    let measures = 0;

    for (let pageIndex = 0; pageIndex < pages.length; pageIndex += 1) {
      await renderDocumentToCanvas(model, paintCanvas(() => { measures += 1; }), pageIndex, {
        dpr: 1,
        width: 750,
        prebuiltPages: pages,
      });
    }

    expect(measures).toBe(0);
    const afterPaint = retainedTables(pages).map(({ fragment }) => fragment);
    expect(afterPaint).toHaveLength(fragments.length);
    for (let index = 0; index < fragments.length; index += 1) {
      expect(afterPaint[index]).toBe(fragments[index]);
    }
  });
});
