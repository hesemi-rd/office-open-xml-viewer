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
 * ECMA-376 §17.4.80 + §17.18.37 ST_HeightRule:
 *
 *   `exact` — the row height is exactly @val; content taller than that is
 *   clipped *vertically* so it does not bleed into adjacent rows.
 *
 * The clip is on the Y axis only. Horizontal clipping has no spec basis: the
 * row's left/right extent is controlled by the table grid and cell margins, not
 * by hRule.
 *
 * The historical bug pinned here: `renderCell` used to clip the full bounding
 * box `(x, y, w, h)` of an `exact` cell. When a nested inner table's left
 * border landed exactly on the outer cell's left edge (e.g. outer
 * `tcMar.left=0` + outer `tblCellMar.left=0` + inner `tblInd=0`), a half-pixel
 * 0.5 pt stroke straddled the clip boundary and only the right half survived —
 * the left edge of the inner frame visibly vanished.
 *
 * This test pins that the clip rect is **Y-axis only**: x spans the whole
 * canvas, so any vertical line drawn at the cell's left edge is preserved while
 * tall content is still clipped to the row's Y band.
 */

/** Recording 2D context. Captures every `rect()` call so we can pick out the
 *  one issued by the `clipExact` branch (immediately before `clip()`). */
interface RectCall { x: number; y: number; w: number; h: number; }
function makeRecordingCanvas(): {
  canvas: HTMLCanvasElement;
  rectCalls: RectCall[];
  /** Mutable counter — read via `.count` so the closure captures by reference. */
  clipCounter: { count: number };
} {
  let font = '10px serif';
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
  const rectCalls: RectCall[] = [];
  const clipCounter = { count: 0 };
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
    save() {}, restore() {}, beginPath() {}, closePath() {},
    moveTo() {}, lineTo() {}, stroke() {}, fill() {}, fillRect() {},
    strokeRect() {},
    rect(x: number, y: number, w: number, h: number) {
      rectCalls.push({ x, y, w, h });
    },
    clip() { clipCounter.count += 1; },
    scale() {}, translate() {},
    setLineDash() {}, clearRect() {}, arc() {},
    quadraticCurveTo() {}, bezierCurveTo() {},
    createLinearGradient() { return { addColorStop() {} }; },
    drawImage() {}, fillText() {}, strokeText() {},
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    textAlign: 'left' as CanvasTextAlign, direction: 'ltr' as CanvasDirection,
    globalAlpha: 1, lineCap: 'butt' as CanvasLineCap, lineJoin: 'miter' as CanvasLineJoin,
  };
  const canvas = {
    width: 0,
    height: 0,
    style: {} as Record<string, string>,
    getContext: () => ctx,
  };
  // Real CanvasRenderingContext2D has a `.canvas` back-reference. The renderer
  // reads it (e.g. `ctx.canvas.width` in the §17.4.80 Y-axis clip) so wire it
  // up here too.
  (ctx as unknown as { canvas: unknown }).canvas = canvas;
  return { canvas: canvas as unknown as HTMLCanvasElement, rectCalls, clipCounter };
}

function emptyBorders() {
  return { top: null, bottom: null, left: null, right: null, insideH: null, insideV: null };
}
function solidBorder() {
  return { width: 0.5, color: '000000', style: 'single' };
}
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

/** Build a 2-level nested table: outer 1×1 row with `hRule="exact"` and zero
 *  cell margins; inner 1×1 table with `tblInd=0` (default) so its left border
 *  lands on the outer cell's left edge — the exact geometry that exposed the
 *  half-pixel masking bug on sample-2's "販売員/発注書番号/支払条件" block. */
function nestedTableDoc(): DocxDocumentModel {
  const innerCell: DocTableCell = {
    content: [{ type: 'paragraph', ...bodyParagraph('inner') }],
    colSpan: 1,
    vMerge: null,
    borders: allSolid(),
    background: null,
    vAlign: 'top',
    widthPt: 100,
  } as unknown as DocTableCell;

  const innerRow: DocTableRow = {
    cells: [innerCell],
    rowHeight: null,
    rowHeightRule: 'auto',
    isHeader: false,
  } as unknown as DocTableRow;

  const innerTable: DocTable = {
    colWidths: [100],
    rows: [innerRow],
    borders: allSolid(),
    cellMarginTop: 0,
    cellMarginBottom: 0,
    cellMarginLeft: 0,
    cellMarginRight: 0,
    jc: 'left',
  } as unknown as DocTable;

  const outerCell: DocTableCell = {
    content: [
      // Word emits a leading paragraph then the nested table then a trailing
      // structural empty paragraph (§17.4.7).
      { type: 'paragraph', ...bodyParagraph('') },
      { type: 'table', ...innerTable },
      { type: 'paragraph', ...bodyParagraph('') },
    ],
    colSpan: 1,
    vMerge: null,
    borders: allSolid(),
    background: null,
    vAlign: 'top',
    widthPt: 100,
    marginTop: 0, marginBottom: 0, marginLeft: 0, marginRight: 0,
  } as unknown as DocTableCell;

  // hRule=exact + a tight @val so the clipExact branch fires (cell.vMerge !== true).
  const outerRow: DocTableRow = {
    cells: [outerCell],
    rowHeight: 40, // pt — exact row height
    rowHeightRule: 'exact',
    isHeader: false,
  } as unknown as DocTableRow;

  const outerTable: DocTable = {
    colWidths: [100],
    rows: [outerRow],
    borders: allSolid(),
    cellMarginTop: 0,
    cellMarginBottom: 0,
    cellMarginLeft: 0,
    cellMarginRight: 0,
    jc: 'left',
  } as unknown as DocTable;

  return {
    section: {
      pageWidth: 200, pageHeight: 200,
      marginTop: 0, marginRight: 0, marginBottom: 0, marginLeft: 0,
      headerDistance: 0, footerDistance: 0,
      titlePage: false, evenAndOddHeaders: false,
    } as SectionProps,
    body: [{ type: 'table', ...outerTable }],
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    fontFamilyClasses: { 'Times New Roman': 'roman' },
  } as unknown as DocxDocumentModel;
}

describe('renderCell clipExact (§17.4.80 + §17.18.37) — Y-axis only clip', () => {
  it('clip rect spans the full canvas width so a nested table\'s left border survives', async () => {
    const { canvas, rectCalls, clipCounter } = makeRecordingCanvas();
    // scale = 1 px/pt (width 200, pageWidth 200 ⇒ scale=1).
    await renderDocumentToCanvas(nestedTableDoc(), canvas, 0, { dpr: 1, width: 200 });

    // The clipExact branch must fire at least once (clip() called) so a rect
    // immediately preceding it carries the clipExact bbox.
    expect(clipCounter.count, 'clipExact must engage for hRule="exact"').toBeGreaterThan(0);

    // Find a rect whose Y band corresponds to the exact row (Y=0, height ~ 40
    // px since scale=1). Multiple rect() calls may happen, so search by Y/H.
    //
    // POST-FIX: the clipExact rect must be Y-axis only — x === 0 and
    // width === canvas.width. The cell sits at x=0/width=100 here, so if the
    // old (full-bbox) clip is in effect, the matching rect has w=100, and this
    // assertion fails. With the fix in place the rect has w=canvas.width (200).
    const canvasW = canvas.width;
    const clipRowRect = rectCalls.find(
      (r) => Math.abs(r.y - 0) < 1e-6 && Math.abs(r.h - 40) < 1e-6,
    );
    expect(clipRowRect, 'a rect at the exact row Y band must exist').toBeDefined();
    expect(clipRowRect!.x, 'clipExact rect.x must be 0 (canvas left)').toBeCloseTo(0, 6);
    expect(clipRowRect!.w, 'clipExact rect.w must span the full canvas').toBeCloseTo(canvasW, 6);
    // Sanity: the canvas IS wider than the outer cell (200 px vs 100 px), so a
    // w === canvas.width clip cannot coincidentally equal the cell width.
    expect(canvasW).toBeGreaterThan(100);
  });

  it('clipExact still bounds Y so taller content cannot bleed into adjacent rows', async () => {
    const { canvas, rectCalls, clipCounter } = makeRecordingCanvas();
    await renderDocumentToCanvas(nestedTableDoc(), canvas, 0, { dpr: 1, width: 200 });

    expect(clipCounter.count).toBeGreaterThan(0);
    // The Y-band of the clipExact rect must match the row (y=0..40 here).
    const clipRowRect = rectCalls.find(
      (r) => Math.abs(r.y - 0) < 1e-6 && Math.abs(r.h - 40) < 1e-6,
    );
    expect(clipRowRect).toBeDefined();
    expect(clipRowRect!.y).toBeCloseTo(0, 6);
    expect(clipRowRect!.h).toBeCloseTo(40, 6);
  });
});
