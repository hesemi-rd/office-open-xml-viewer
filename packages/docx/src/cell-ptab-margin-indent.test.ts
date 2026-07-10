import { describe, it, expect } from 'vitest';
import { renderDocumentToCanvas } from './renderer.js';
import type {
  BodyElement,
  CellElement,
  DocParagraph,
  DocRun,
  DocTable,
  DocTableCell,
  DocTableRow,
  DocxDocumentModel,
  SectionProps,
} from './types';

// ─────────────────────────────────────────────────────────────────────────────
// Finding 4 — LTR table-cell paragraph: `<w:ptab w:relativeTo="margin">` resolves
// against the text margin (paraW + right indent), matching the paint side.
//
// ECMA-376 §17.3.3.23 (ptab) + §17.18.73 (ST_PTabRelativeTo): `relativeTo="margin"`
// positions the absolute tab against the TEXT-MARGIN box, which is independent of
// the paragraph's own left/right indents; `relativeTo="indent"` positions it
// against the (indented) content box.
//
// The old cell-specific measurer deliberately passed `marginRightPx = paraW`
// (NOT paraW + indRight) for LTR cell paragraphs, documented as a deferred ptab
// limitation. The unified placement-aware `measureParagraph` now passes
// `paragraphWidthPt + physicalIndentRightPt` — the SAME `marginRightPx = paraW +
// indRight` the paint side (renderParagraph) uses — so a cell paragraph's measured
// line geometry and its painted geometry resolve the margin ptab identically. When
// the paginator stamps a cell paragraph's scale-1 lines, the paint pass reuses
// them, so a measure/paint disagreement here would surface directly in the painted
// x. These tests pin the unified margin+indRight semantics end to end.
// ─────────────────────────────────────────────────────────────────────────────

interface FillCall { text: string; x: number; }

function makeRecordingCanvas(): { canvas: HTMLCanvasElement; fills: FillCall[] } {
  let font = '10px serif';
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
  const fills: FillCall[] = [];
  const ctx = {
    get font() { return font; },
    set font(v: string) { font = v; },
    letterSpacing: '0px',
    measureText: (s: string) => {
      const p = px();
      return {
        width: [...s].length * p,
        fontBoundingBoxAscent: p * 0.8, fontBoundingBoxDescent: p * 0.2,
        actualBoundingBoxAscent: p * 0.8, actualBoundingBoxDescent: p * 0.2,
      } as TextMetrics;
    },
    save() {}, restore() {}, beginPath() {}, closePath() {},
    moveTo() {}, lineTo() {}, stroke() {}, fill() {}, fillRect() {}, strokeRect() {},
    rect() {}, clip() {}, scale() {}, translate() {}, setLineDash() {}, clearRect() {}, arc() {},
    quadraticCurveTo() {}, bezierCurveTo() {},
    createLinearGradient() { return { addColorStop() {} }; },
    drawImage() {},
    fillText(text: string, x: number) { fills.push({ text, x }); },
    strokeText() {},
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    textAlign: 'left' as CanvasTextAlign, direction: 'ltr' as CanvasDirection,
    globalAlpha: 1, lineCap: 'butt' as CanvasLineCap, lineJoin: 'miter' as CanvasLineJoin,
  };
  const canvas = { width: 0, height: 0, style: {} as Record<string, string>, getContext: () => ctx };
  (ctx as unknown as { canvas: unknown }).canvas = canvas;
  return { canvas: canvas as unknown as HTMLCanvasElement, fills };
}

function textRun(text: string): DocRun {
  return {
    type: 'text', text,
    bold: false, italic: false, underline: false, strikethrough: false,
    fontSize: 10, color: null, fontFamily: 'Times New Roman', fontFamilyEastAsia: 'Times New Roman',
    isLink: false, background: null, vertAlign: null, hyperlink: null,
  } as unknown as DocRun;
}

function ptabRun(alignment: 'left' | 'center' | 'right', relativeTo: 'margin' | 'indent'): DocRun {
  return { type: 'ptab', alignment, relativeTo, leader: 'none', fontSize: 10 } as unknown as DocRun;
}

function cellPara(runs: DocRun[], indentRight: number): CellElement {
  return {
    type: 'paragraph',
    alignment: 'left',
    indentLeft: 0, indentRight, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null,
    numbering: null, tabStops: [],
    runs,
    defaultFontSize: 10, defaultFontFamily: 'Times New Roman', widowControl: false,
  } as unknown as CellElement;
}

const PAGE_W = 300;
const FS = 10; // glyph width px; scale = 1 px/pt (canvas width == pageWidth)

function docWithCellPara(el: CellElement): DocxDocumentModel {
  const cell = {
    content: [el],
    colSpan: 1, vMerge: null,
    borders: { top: null, bottom: null, left: null, right: null, insideH: null, insideV: null },
    background: null, vAlign: 'top', widthPt: PAGE_W,
  } as DocTableCell;
  const row = { cells: [cell], rowHeight: null, rowHeightRule: 'auto', isHeader: false } as DocTableRow;
  const table = {
    colWidths: [PAGE_W], rows: [row],
    borders: { top: null, bottom: null, left: null, right: null, insideH: null, insideV: null },
    cellMarginTop: 0, cellMarginBottom: 0, cellMarginLeft: 0, cellMarginRight: 0, jc: 'left',
  } as DocTable;
  return {
    section: {
      pageWidth: PAGE_W, pageHeight: 400,
      marginTop: 0, marginRight: 0, marginBottom: 0, marginLeft: 0,
      headerDistance: 0, footerDistance: 0, titlePage: false, evenAndOddHeaders: false,
    } as SectionProps,
    body: [{ type: 'table', ...table } as BodyElement],
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    fontFamilyClasses: { 'Times New Roman': 'roman' },
  } as unknown as DocxDocumentModel;
}

async function render(el: CellElement): Promise<FillCall[]> {
  const { canvas, fills } = makeRecordingCanvas();
  await renderDocumentToCanvas(docWithCellPara(el), canvas, 0, { dpr: 1, width: PAGE_W });
  return fills;
}

describe('table-cell ptab (§17.3.3.23) resolves margin against paraW + right indent', () => {
  // Cell content box = full 300 pt (zero cell margins). Paragraph has a 40 pt right
  // indent, so paraW = 260 and the TEXT MARGIN right edge = paraW + indRight = 300
  // (the cell content-box edge), independent of the indent.
  const INDENT_RIGHT = 40;

  it('right ptab relativeTo="margin" right-aligns to the cell text margin, ignoring the right indent', async () => {
    const fills = await render(cellPara([ptabRun('right', 'margin'), textRun('99')], INDENT_RIGHT));
    const f = fills.find((c) => c.text === '99');
    expect(f, '"99" must be drawn').toBeDefined();
    // Margin right edge = paraW(260) + indRight(40) = 300 → 2-glyph number ends
    // there ⇒ starts at 300 − 20 = 280. The old cell measurer would have resolved
    // the margin at paraW (260), landing it at 240; the unified marginRightPx makes
    // the measured stamp and the painted line agree at the true margin.
    expect(f!.x + 2 * FS).toBeCloseTo(PAGE_W, 3);
  });

  it('right ptab relativeTo="indent" aligns to the INDENTED content box (contrast)', async () => {
    const fills = await render(cellPara([ptabRun('right', 'indent'), textRun('99')], INDENT_RIGHT));
    const f = fills.find((c) => c.text === '99');
    expect(f, '"99" must be drawn').toBeDefined();
    // Content box right edge = paraW = 260 (right indent excluded) ⇒ 260 − 20 = 240.
    // This confirms the two relativeTo modes are distinguished in a cell: margin
    // includes the right indent, indent does not.
    expect(f!.x + 2 * FS).toBeCloseTo(PAGE_W - INDENT_RIGHT, 3);
  });

  it('center ptab relativeTo="margin" centers on the cell text-margin midpoint', async () => {
    const fills = await render(cellPara([ptabRun('center', 'margin'), textRun('AB')], INDENT_RIGHT));
    const f = fills.find((c) => c.text === 'AB');
    expect(f, '"AB" must be drawn').toBeDefined();
    // Margin box [0, 300] midpoint = 150 → 2-glyph text starts at 150 − (2·FS)/2 = 140.
    expect(f!.x).toBeCloseTo(PAGE_W / 2 - (2 * FS) / 2, 3);
  });
});
