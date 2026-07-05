import { describe, it, expect } from 'vitest';
import { renderDocumentToCanvas } from './renderer.js';
import type {
  BodyElement,
  DocParagraph,
  DocTable,
  DocTableCell,
  DocTableRow,
  DocxDocumentModel,
  DocxTextRun,
  SectionProps,
} from './types';

// ECMA-376 §17.3.2.6 — `<w:color w:val="auto"/>` (or no color anywhere in the
// style hierarchy) means the consumer picks a color that CONTRASTS with the
// background BEHIND the run. sample-28 p.17: a table cell with a near-black fill
// (`<w:tcPr><w:shd w:fill="0C0C0C"/>`) holds a run with NO color, so Word paints
// the text WHITE. The run has no run-level `<w:shd>`, so the effective background
// is the CELL shading — which must be folded into the auto-contrast decision.
//
// These tests capture the canvas fillStyle at each fillText so the glyph color is
// asserted directly (autoContrastColor's own black/white math is unit-tested in
// core/shape/paint.test.ts; here we prove the CELL background reaches it).

interface DrawnGlyph { text: string; style: string }

function makeRecordingCanvas(): { canvas: HTMLCanvasElement; glyphs: DrawnGlyph[] } {
  const glyphs: DrawnGlyph[] = [];
  let font = '16px Arial';
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '16');
  const ctx = {
    get font() { return font; },
    set font(v: string) { font = v; },
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
    moveTo() {}, lineTo() {}, stroke() {}, fill() {}, clip() {}, rect() {},
    scale() {}, translate() {}, setLineDash() {}, drawImage() {}, clearRect() {},
    arc() {}, quadraticCurveTo() {}, bezierCurveTo() {},
    createLinearGradient() { return { addColorStop() {} }; },
    fillRect() {}, strokeRect() {}, strokeText() {},
    fillText(text: string) {
      if (text.trim()) glyphs.push({ text, style: String(this.fillStyle).toUpperCase() });
    },
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    textAlign: 'left' as CanvasTextAlign, direction: 'ltr' as CanvasDirection,
    globalAlpha: 1, lineCap: 'butt' as CanvasLineCap, lineJoin: 'miter' as CanvasLineJoin,
  };
  const canvas = {
    width: 0, height: 0, style: {} as Record<string, string>,
    getContext: () => ctx,
  };
  return { canvas: canvas as unknown as HTMLCanvasElement, glyphs };
}

function textRun(text: string, extra: Partial<DocxTextRun> = {}): DocxTextRun {
  return {
    text, bold: false, italic: false, underline: false, strikethrough: false,
    fontSize: 16, color: null, fontFamily: 'Arial', fontFamilyEastAsia: 'Arial',
    isLink: false, background: null, vertAlign: null, hyperlink: null,
    ...extra,
  };
}

function cellPara(run: DocxTextRun, shading?: string | null): DocParagraph {
  return {
    alignment: 'left', indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null, tabStops: [],
    runs: [{ type: 'text', ...run } as DocParagraph['runs'][number]],
    defaultFontSize: 16, defaultFontFamily: 'Arial', widowControl: false,
    shading: shading ?? null,
  } as DocParagraph;
}

function cell(content: DocParagraph[], background: string | null): DocTableCell {
  return {
    content: content.map((p) => ({ type: 'paragraph', ...p })) as unknown as DocTableCell['content'],
    colSpan: 1, vMerge: null,
    borders: { top: null, bottom: null, left: null, right: null, insideH: null, insideV: null },
    background,
    vAlign: 'top', widthPt: null,
  } as DocTableCell;
}

function tableDoc(cells: DocTableCell[]): DocxDocumentModel {
  const rows: DocTableRow[] = [{ cells, rowHeight: null, rowHeightRule: null, isHeader: false } as unknown as DocTableRow];
  const table: DocTable = {
    colWidths: cells.map(() => 100),
    rows,
    borders: { top: null, bottom: null, left: null, right: null, insideH: null, insideV: null },
    cellMarginTop: 0, cellMarginBottom: 0, cellMarginLeft: 0, cellMarginRight: 0,
    jc: 'left',
  } as unknown as DocTable;
  return {
    section: {
      pageWidth: 400, pageHeight: 400,
      marginTop: 0, marginRight: 0, marginBottom: 0, marginLeft: 0,
      headerDistance: 0, footerDistance: 0, titlePage: false, evenAndOddHeaders: false,
    } as SectionProps,
    body: [{ type: 'table', ...table } as unknown as BodyElement],
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    fontFamilyClasses: { Arial: 'swiss' },
  } as unknown as DocxDocumentModel;
}

async function renderGlyphs(cells: DocTableCell[]): Promise<DrawnGlyph[]> {
  const { canvas, glyphs } = makeRecordingCanvas();
  await renderDocumentToCanvas(tableDoc(cells), canvas, 0, { dpr: 1, width: 400 });
  return glyphs;
}

describe('auto text color folds in table-cell shading (§17.3.2.6, sample-28 p.17)', () => {
  it('paints an auto-color run WHITE inside a near-black cell (fill=0C0C0C)', async () => {
    // The sample-28 shape: a `w:color`-less run (colorAuto) inside a cell whose
    // only shading is the near-black `<w:tcPr><w:shd w:fill="0C0C0C"/>`. Without
    // folding the cell fill into the effective background the run stayed black
    // (invisible). It must resolve to white.
    const glyphs = await renderGlyphs([
      cell([cellPara(textRun('X', { colorAuto: true }))], '0C0C0C'),
    ]);
    const g = glyphs.find((x) => x.text.includes('X'));
    expect(g).toBeDefined();
    expect(g?.style).toBe('#FFFFFF');
  });

  it('keeps an auto-color run BLACK inside a light cell (fill=D9D9D9)', async () => {
    // The neighbouring light-grey header cells: auto stays black against a light
    // fill (contrast pick is a no-flip). Guards against over-inverting.
    const glyphs = await renderGlyphs([
      cell([cellPara(textRun('Y', { colorAuto: true }))], 'D9D9D9'),
    ]);
    const g = glyphs.find((x) => x.text.includes('Y'));
    expect(g).toBeDefined();
    expect(g?.style).toBe('#000000');
  });

  it('lets an EXPLICIT run color win over the cell-shading contrast pick', async () => {
    // A run with an explicit color is not "auto" — the cell fill must NOT override
    // it (only auto folds in the background). Red stays red on black.
    const glyphs = await renderGlyphs([
      cell([cellPara(textRun('Z', { color: 'FF0000' }))], '0C0C0C'),
    ]);
    const g = glyphs.find((x) => x.text.includes('Z'));
    expect(g).toBeDefined();
    expect(g?.style).toBe('#FF0000');
  });

  it('prefers RUN shading over cell shading for the auto contrast (run fill wins)', async () => {
    // A run-level light `<w:shd>` inside a dark cell: the background immediately
    // behind the glyphs is the RUN fill, so auto contrasts against THAT (stays
    // black on the light run fill), not the cell.
    const glyphs = await renderGlyphs([
      cell([cellPara(textRun('W', { colorAuto: true, background: 'FFFFFF' }))], '0C0C0C'),
    ]);
    const g = glyphs.find((x) => x.text.includes('W'));
    expect(g).toBeDefined();
    expect(g?.style).toBe('#000000');
  });
});
