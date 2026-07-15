import { describe, it, expect } from 'vitest';
import {
  createLayoutServices,
  paginateDocument,
  renderDocumentToCanvas,
} from './renderer.js';
import type {
  BodyElement, CellElement, DocNote, DocParagraph, DocTable, DocTableCell,
  DocTableRow, DocxDocumentModel, SectionProps,
} from './types';

// ECMA-376 §17.11.10 — a footnote is drawn at the bottom of the page that holds
// its reference, regardless of WHERE in the document story that reference sits.
// A `<w:footnoteReference>` can appear inside a table cell (the cell paragraph
// carries the noteRef run), so the footnote block must be drawn — and the body
// area reserved — even when the only reference on a page lives in a table.
//
// Regression: `drawPageFootnotes` and the pagination reserve pass only scanned
// TOP-LEVEL paragraphs, so a footnote referenced solely from a table cell had
// its marker painted (the cell run draws normally) but its content silently
// dropped (issue #840). Endnotes were unaffected because `drawEndnotes` draws
// every note unconditionally without scanning the body for references.

const TEST_FONT = 'Times New Roman';

interface Call { text: string; y: number; }
function makeRecordingCanvas(): { canvas: HTMLCanvasElement; calls: Call[] } {
  let font = '10px serif';
  const calls: Call[] = [];
  const ctx = {
    get font() { return font; },
    set font(v: string) { font = v; },
    letterSpacing: '0px',
    measureText: (s: string) => {
      const p = parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
      return {
        width: [...s].length * p * 0.5,
        fontBoundingBoxAscent: p * 0.8, fontBoundingBoxDescent: p * 0.2,
        actualBoundingBoxAscent: p * 0.8, actualBoundingBoxDescent: p * 0.2,
      } as TextMetrics;
    },
    save() {}, restore() {}, beginPath() {}, closePath() {},
    moveTo() {}, lineTo() {}, stroke() {}, fill() {}, fillRect() {},
    strokeRect() {}, clip() {}, rect() {}, scale() {}, translate() {}, rotate() {},
    setLineDash() {}, clearRect() {}, arc() {}, quadraticCurveTo() {},
    bezierCurveTo() {}, createLinearGradient() { return { addColorStop() {} }; },
    drawImage() {},
    fillText(s: string, _x: number, y: number) { calls.push({ text: s, y }); },
    strokeText(s: string, _x: number, y: number) { calls.push({ text: s, y }); },
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    textAlign: 'left' as CanvasTextAlign, direction: 'ltr' as CanvasDirection,
    globalAlpha: 1, lineCap: 'butt' as CanvasLineCap, lineJoin: 'miter' as CanvasLineJoin,
  };
  const canvas = { width: 0, height: 0, style: {} as Record<string, string>, getContext: () => ctx };
  return { canvas: canvas as unknown as HTMLCanvasElement, calls };
}

function textRun(text: string, extra: Record<string, unknown> = {}) {
  return {
    type: 'text', text, bold: false, italic: false, underline: false,
    strikethrough: false, fontSize: 10, color: null, fontFamily: TEST_FONT,
    fontFamilyEastAsia: '', isLink: false, background: null, vertAlign: null, hyperlink: null,
    ...extra,
  };
}

function para(runs: Array<Record<string, unknown>>): DocParagraph {
  return {
    type: 'paragraph', alignment: 'left',
    indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null,
    numbering: null, tabStops: [],
    runs: runs as DocParagraph['runs'],
    defaultFontSize: 10, defaultFontFamily: TEST_FONT, widowControl: false,
  } as unknown as DocParagraph;
}

function cell(content: CellElement[]): DocTableCell {
  return {
    content,
    colSpan: 1, vMerge: null,
    borders: { top: null, bottom: null, left: null, right: null, insideH: null, insideV: null },
    background: null, vAlign: 'top', widthPt: 380,
  } as DocTableCell;
}

function tableOf(rows: DocTableRow[]): DocTable {
  return {
    colWidths: [380], rows,
    borders: { top: null, bottom: null, left: null, right: null, insideH: null, insideV: null },
    cellMarginTop: 0, cellMarginBottom: 0, cellMarginLeft: 0, cellMarginRight: 0,
    jc: 'left',
  } as DocTable;
}

function row(cells: DocTableCell[]): DocTableRow {
  return { cells, rowHeight: null, rowHeightRule: 'auto', isHeader: false } as DocTableRow;
}

function docWith(body: BodyElement[], footnotes: DocNote[]): DocxDocumentModel {
  const section: SectionProps = {
    pageWidth: 400, pageHeight: 400,
    marginTop: 10, marginRight: 10, marginBottom: 10, marginLeft: 10,
    headerDistance: 4, footerDistance: 4, titlePage: false, evenAndOddHeaders: false,
    sectionStart: 'nextPage',
  } as SectionProps;
  return {
    section, body,
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    fontFamilyClasses: { [TEST_FONT]: 'roman' },
    footnotes,
  } as unknown as DocxDocumentModel;
}

async function renderPage0(doc: DocxDocumentModel): Promise<Call[]> {
  const { canvas, calls } = makeRecordingCanvas();
  await renderDocumentToCanvas(doc, canvas, 0, { dpr: 1, width: 400 });
  return calls;
}

describe('footnote referenced from a table cell (ECMA-376 §17.11.10)', () => {
  const footnotes: DocNote[] = [
    { id: 'fn1', content: [para([textRun('NOTE')]) as unknown as BodyElement] },
  ];

  it('draws the footnote block for a reference that lives inside a table cell', async () => {
    // The ONLY footnote reference on the page sits in a table cell.
    const cellPara = para([
      textRun('CELL'),
      textRun('', { noteRef: { kind: 'footnote', id: 'fn1' }, vertAlign: 'super' }),
    ]);
    const table = tableOf([row([cell([cellPara as unknown as CellElement])])]);
    const body: BodyElement[] = [
      para([textRun('BODY')]) as unknown as BodyElement,
      { type: 'table', ...table } as unknown as BodyElement,
    ];
    const calls = await renderPage0(docWith(body, footnotes));

    // The footnote content must be painted somewhere on the page.
    const noteY = calls.filter((c) => c.text === 'NOTE').map((c) => c.y);
    expect(noteY.length).toBeGreaterThan(0);
    // And it must sit at the BOTTOM of the page (below the body/table content).
    const cellY = Math.max(...calls.filter((c) => c.text === 'CELL').map((c) => c.y));
    expect(Math.max(...noteY)).toBeGreaterThan(cellY);
  });

  it('reserves body space so the footnote does not overlap table content', async () => {
    // A tall footnote referenced from a cell must shrink the body area so the
    // note (drawn at the bottom margin) never overlaps the table content above.
    const tallNote: DocNote[] = [
      {
        id: 'fn1',
        content: Array.from({ length: 6 }, () => para([textRun('NOTELINE')]) as unknown as BodyElement),
      },
    ];
    const cellPara = para([
      textRun('CELL'),
      textRun('', { noteRef: { kind: 'footnote', id: 'fn1' }, vertAlign: 'super' }),
    ]);
    const table = tableOf([row([cell([cellPara as unknown as CellElement])])]);
    const body: BodyElement[] = [
      { type: 'table', ...table } as unknown as BodyElement,
    ];
    const calls = await renderPage0(docWith(body, tallNote));

    const noteY = calls.filter((c) => c.text === 'NOTELINE').map((c) => c.y);
    const cellY = calls.filter((c) => c.text === 'CELL').map((c) => c.y);
    expect(noteY.length).toBeGreaterThan(0);
    expect(cellY.length).toBeGreaterThan(0);
    // Every footnote line sits below the cell text — no overlap.
    expect(Math.min(...noteY)).toBeGreaterThan(Math.max(...cellY));
  });

  it('descends into a NESTED table (§17.4.7) to find the reference', async () => {
    // The reference lives in a cell of a table nested inside another table's
    // cell — the collector must recurse through both levels.
    const innerPara = para([
      textRun('INNER'),
      textRun('', { noteRef: { kind: 'footnote', id: 'fn1' }, vertAlign: 'super' }),
    ]);
    const innerTable = tableOf([row([cell([innerPara as unknown as CellElement])])]);
    const outerCell = cell([{ type: 'table', ...innerTable } as unknown as CellElement]);
    const outerTable = tableOf([row([outerCell])]);
    const body: BodyElement[] = [
      { type: 'table', ...outerTable } as unknown as BodyElement,
    ];
    const calls = await renderPage0(docWith(body, footnotes));

    const noteY = calls.filter((c) => c.text === 'NOTE').map((c) => c.y);
    expect(noteY.length).toBeGreaterThan(0);
  });

  it('does not draw footnotes for a document with no references (byte-identical guard)', async () => {
    // A table document that references NO footnote leaves the note undrawn — the
    // fix is gated behind an actual reference, so footnote-free docs are unchanged.
    const table = tableOf([row([cell([para([textRun('CELL')]) as unknown as CellElement])])]);
    const body: BodyElement[] = [{ type: 'table', ...table } as unknown as BodyElement];
    const calls = await renderPage0(docWith(body, footnotes));
    expect(calls.filter((c) => c.text === 'NOTE').length).toBe(0);
  });

  it('remeasures a footnote PAGE field with its section display number', () => {
    const notePage = para([{
      type: 'field', fieldType: 'page', instruction: 'PAGE', fallbackText: '?',
      bold: false, italic: false, underline: false, strikethrough: false,
      fontSize: 10, color: null, fontFamily: TEST_FONT, background: null, vertAlign: null,
    }]);
    const reference = para([
      textRun('REFERENCE'),
      textRun('', { noteRef: { kind: 'footnote', id: 'fn-page' }, vertAlign: 'super' }),
    ]);
    const model = docWith(
      [reference as unknown as BodyElement],
      [{ id: 'fn-page', content: [notePage as unknown as BodyElement] }],
    );
    model.section.pageNumType = { start: 10, fmt: 'upperRoman' };
    const { canvas } = makeRecordingCanvas();
    const ctx = canvas.getContext('2d') as CanvasRenderingContext2D;
    const base = createLayoutServices(model, { measureContext: ctx });
    const shapedTexts: string[] = [];
    const services = Object.freeze({
      ...base,
      text: Object.freeze({
        ...base.text,
        shape(request: Parameters<typeof base.text.shape>[0]) {
          shapedTexts.push(request.text);
          return base.text.shape(request);
        },
      }),
    });

    const globals = globalThis as unknown as { OffscreenCanvas?: unknown };
    const previousOffscreenCanvas = globals.OffscreenCanvas;
    globals.OffscreenCanvas = class {
      getContext() { return ctx; }
    };
    try {
      paginateDocument(model, services);
    } finally {
      if (previousOffscreenCanvas === undefined) delete globals.OffscreenCanvas;
      else globals.OffscreenCanvas = previousOffscreenCanvas;
    }

    expect(shapedTexts).toContain('X');
  });

  it('uses the final split-table page that owns the footnote reserve', () => {
    const notePage = para([{
      type: 'field', fieldType: 'page', instruction: 'PAGE', fallbackText: '?',
      bold: false, italic: false, underline: false, strikethrough: false,
      fontSize: 10, color: null, fontFamily: TEST_FONT, background: null, vertAlign: null,
    }]);
    const rows = Array.from({ length: 12 }, (_, rowIndex) => row([cell([para([
      textRun(`ROW${rowIndex}`),
      ...(rowIndex === 0
        ? [textRun('', { noteRef: { kind: 'footnote', id: 'fn-page' }, vertAlign: 'super' })]
        : []),
    ]) as unknown as CellElement])]));
    const model = docWith(
      [{ type: 'table', ...tableOf(rows) } as unknown as BodyElement],
      [{ id: 'fn-page', content: [notePage as unknown as BodyElement] }],
    );
    model.section.pageHeight = 80;
    model.section.pageNumType = { start: 10, fmt: 'decimal' };
    const { canvas } = makeRecordingCanvas();
    const ctx = canvas.getContext('2d') as CanvasRenderingContext2D;
    const base = createLayoutServices(model, { measureContext: ctx });
    const shapedTexts: string[] = [];
    const services = Object.freeze({
      ...base,
      text: Object.freeze({
        ...base.text,
        shape(request: Parameters<typeof base.text.shape>[0]) {
          shapedTexts.push(request.text);
          return base.text.shape(request);
        },
      }),
    });
    const globals = globalThis as unknown as { OffscreenCanvas?: unknown };
    const previousOffscreenCanvas = globals.OffscreenCanvas;
    globals.OffscreenCanvas = class {
      getContext() { return ctx; }
    };
    let pages: ReturnType<typeof paginateDocument>;
    try {
      pages = paginateDocument(model, services);
    } finally {
      if (previousOffscreenCanvas === undefined) delete globals.OffscreenCanvas;
      else globals.OffscreenCanvas = previousOffscreenCanvas;
    }

    expect(pages.length).toBeGreaterThan(1);
    expect(shapedTexts).toContain(String(9 + pages.length));
  });
});
