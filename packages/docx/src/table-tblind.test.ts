import { describe, it, expect } from 'vitest';
import { renderDocumentToCanvas } from './renderer.js';
import type {
  DocxDocumentModel,
  DocTable,
  DocTableRow,
  DocTableCell,
  DocParagraph,
  SectionProps,
} from './types';

/**
 * ECMA-376 §17.4.50 `<w:tblInd>` — indentation added before the table's LEADING
 * edge (left in LTR, right in RTL/`bidiVisual`), shifting the table into the text
 * margin. A NEGATIVE value pulls the table OUTWARD past the leading margin toward
 * the page edge. Applies only when the resolved `jc` is left/leading.
 *
 * These tests pin the physical x-origin the renderer resolves for each direction
 * by capturing every `moveTo`/`lineTo` x-coordinate (cell borders are stroked via
 * `strokeCrispSegment` → `moveTo`/`lineTo`) and reading the table's left/right
 * extent from them. A vertical border is nudged ≤0.5 px by the crispness snap
 * (`crispOffset`), so the extent assertions allow a 0.75 px tolerance.
 */

/** Assert `actual` is within 0.75 px of `expected` (absorbs the ≤0.5 px crisp
 *  snap on a thin vertical border). */
function expectNear(actual: number, expected: number, msg: string): void {
  expect(Math.abs(actual - expected), `${msg} (got ${actual}, want ~${expected})`).toBeLessThanOrEqual(0.75);
}

function makeRecordingCanvas(): { canvas: HTMLCanvasElement; xs: number[] } {
  let font = '10px serif';
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
  const xs: number[] = [];
  let transform = { scaleX: 1, translateX: 0 };
  const transformStack: Array<typeof transform> = [];
  const record = (x: number) => xs.push(transform.translateX + x * transform.scaleX);
  const ctx = {
    get font() { return font; },
    set font(v: string) { font = v; },
    letterSpacing: '0px',
    measureText: (s: string) => {
      const p = px();
      return {
        width: [...s].length * p,
        fontBoundingBoxAscent: p * 0.8,
        fontBoundingBoxDescent: p * 0.2,
        actualBoundingBoxAscent: p * 0.8,
        actualBoundingBoxDescent: p * 0.2,
      } as TextMetrics;
    },
    save() { transformStack.push({ ...transform }); },
    restore() { transform = transformStack.pop() ?? { scaleX: 1, translateX: 0 }; },
    beginPath() {}, closePath() {},
    moveTo(x: number) { record(x); }, lineTo(x: number) { record(x); },
    stroke() {}, fill() {}, fillRect() {}, strokeRect() {},
    rect() {}, clip() {},
    scale(x: number) { transform.scaleX *= x; },
    translate(x: number) { transform.translateX += x * transform.scaleX; },
    setLineDash() {}, clearRect() {}, arc() {},
    quadraticCurveTo() {}, bezierCurveTo() {},
    createLinearGradient() { return { addColorStop() {} }; },
    drawImage() {}, fillText() {}, strokeText() {},
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    textAlign: 'left' as CanvasTextAlign, direction: 'ltr' as CanvasDirection,
    globalAlpha: 1, lineCap: 'butt' as CanvasLineCap, lineJoin: 'miter' as CanvasLineJoin,
  };
  const canvas = { width: 0, height: 0, style: {} as Record<string, string>, getContext: () => ctx };
  (ctx as unknown as { canvas: unknown }).canvas = canvas;
  return { canvas: canvas as unknown as HTMLCanvasElement, xs };
}

function solidBorder() { return { width: 0.5, color: '000000', style: 'single' }; }
function allSolid() {
  const b = solidBorder();
  return { top: b, bottom: b, left: b, right: b, insideH: b, insideV: b };
}
function bodyParagraph(text: string): DocParagraph {
  return {
    alignment: 'left',
    indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null,
    numbering: null, tabStops: [],
    runs: text === '' ? [] : [{
      type: 'text', text,
      bold: false, italic: false, underline: false, strikethrough: false,
      fontSize: 10, color: null, fontFamily: 'Times New Roman', fontFamilyEastAsia: 'Times New Roman',
      isLink: false, background: null, vertAlign: null, hyperlink: null,
    }],
    defaultFontSize: 10, defaultFontFamily: 'Times New Roman',
    widowControl: false,
  } as unknown as DocParagraph;
}

/** A 1×1 solid-border table (single column `colW` pt wide) with the given
 *  `tblInd` (pt) and bidiVisual flag, on a 200×200 pt page with 20 pt margins
 *  (content [20, 180], width 160). */
function tableDoc(colW: number, tblInd: number | undefined, bidiVisual: boolean): DocxDocumentModel {
  const cell: DocTableCell = {
    content: [{ type: 'paragraph', ...bodyParagraph('x') }],
    colSpan: 1, vMerge: null, borders: allSolid(), background: null, vAlign: 'top',
    widthPt: colW,
  } as unknown as DocTableCell;
  const row: DocTableRow = {
    cells: [cell], rowHeight: null, rowHeightRule: 'auto', isHeader: false,
  } as unknown as DocTableRow;
  const table: DocTable = {
    colWidths: [colW], rows: [row], borders: allSolid(),
    cellMarginTop: 0, cellMarginBottom: 0, cellMarginLeft: 0, cellMarginRight: 0,
    jc: 'left', tblInd, bidiVisual, widthPt: colW,
  } as unknown as DocTable;
  return {
    section: {
      pageWidth: 200, pageHeight: 200,
      marginTop: 20, marginRight: 20, marginBottom: 20, marginLeft: 20,
      headerDistance: 0, footerDistance: 0, titlePage: false, evenAndOddHeaders: false,
    } as SectionProps,
    body: [{ type: 'table', ...table }],
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    fontFamilyClasses: { 'Times New Roman': 'roman' },
  } as unknown as DocxDocumentModel;
}

describe('§17.4.50 tblInd — table indent from the leading margin', () => {
  it('LTR: negative tblInd pulls the LEFT origin outward past the left margin', async () => {
    // scale=1 (width 200 = pageWidth). content=[20,180]. colW=160 fits content.
    // No indent: left origin = contentX = 20. With tblInd=-10 → left origin = 10.
    const noInd = makeRecordingCanvas();
    await renderDocumentToCanvas(tableDoc(160, undefined, false), noInd.canvas, 0, { dpr: 1, width: 200 });
    expectNear(Math.min(...noInd.xs), 20, 'LTR base left origin = contentX');

    const withInd = makeRecordingCanvas();
    await renderDocumentToCanvas(tableDoc(160, -10, false), withInd.canvas, 0, { dpr: 1, width: 200 });
    expectNear(Math.min(...withInd.xs), 10, 'LTR left origin = contentX + tblInd = 20 + (-10)');
  });

  it('RTL (bidiVisual): negative tblInd pushes the RIGHT leading edge into the right margin', async () => {
    // No indent, bidiVisual, colW=160=content: table fills [20,180]; right edge 180.
    const noInd = makeRecordingCanvas();
    await renderDocumentToCanvas(tableDoc(160, undefined, true), noInd.canvas, 0, { dpr: 1, width: 200 });
    expectNear(Math.max(...noInd.xs), 180, 'RTL base right edge = contentRight');

    // tblInd=-10, bidiVisual, colW=160. The table keeps its full 160 width and its
    // leading (RIGHT) edge sits 10 pt PAST the right margin: rightEdge =
    // contentX + contentW - tblInd = 20 + 160 - (-10) = 190. Left origin = 190-160 = 30.
    const withInd = makeRecordingCanvas();
    await renderDocumentToCanvas(tableDoc(160, -10, true), withInd.canvas, 0, { dpr: 1, width: 200 });
    expectNear(Math.max(...withInd.xs), 190, 'RTL right edge = contentRight - tblInd = 190');
    expectNear(Math.min(...withInd.xs), 30, 'RTL left origin = rightEdge - tableW = 30');
  });

  it('applies positive tblInd after right alignment like Word ([MS-OI29500] 2.1.155)', async () => {
    // ECMA-376 §17.4.50 says jc='right' ignores tblInd, but Word explicitly
    // deviates: resolve right alignment first, then translate by the indent.
    const doc = tableDoc(100, 20, false);
    (doc.body[0] as unknown as DocTable).jc = 'right';
    const rec = makeRecordingCanvas();
    await renderDocumentToCanvas(doc, rec.canvas, 0, { dpr: 1, width: 200 });
    expectNear(Math.max(...rec.xs), 200, 'Word shifts the right-aligned table by tblInd');
  });
});
