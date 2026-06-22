import { describe, it, expect } from 'vitest';
import { renderDocumentToCanvas } from './renderer.js';
import type { DocParagraph, DocxDocumentModel, SectionProps } from './types.js';

// ECMA-376 §17.9.8 `<w:lvlJc>` — a list level's marker justification. "right"
// (period-aligned roman/decimal numerals) right-aligns the marker so its RIGHT
// edge sits at the hanging-indent reference (firstLineX = indentLeft + the
// negative firstLine indent); different-width markers ("i." vs "iii.") then share
// that right edge and their periods line up. We previously drew every marker
// LEFT-aligned, so the periods staggered (sample-11's "i./ii./iii./iv." list).
//
// Recording canvas measures each glyph at fontSize px, so "i." = 20 px and
// "iii." = 40 px — distinct widths that expose the alignment.

interface FillCall { text: string; x: number }

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
    save() {}, restore() {}, beginPath() {}, closePath() {}, moveTo() {}, lineTo() {},
    stroke() {}, fill() {}, fillRect() {}, strokeRect() {}, rect() {}, clip() {}, scale() {},
    translate() {}, setLineDash() {}, clearRect() {}, arc() {}, quadraticCurveTo() {}, bezierCurveTo() {},
    createLinearGradient() { return { addColorStop() {} }; }, drawImage() {},
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

function romanItem(markerText: string, body: string, jc: string): DocParagraph {
  return {
    alignment: 'left', indentLeft: 57.6, indentRight: 0, indentFirst: -18,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null, tabStops: [],
    numbering: {
      numId: 7, level: 0, format: 'lowerRoman', text: markerText, indentLeft: 57.6,
      tab: 18, suff: 'tab', jc, fontFamily: 'Times New Roman', fontFamilyEastAsia: 'Times New Roman',
    },
    runs: [{
      type: 'text', text: body, bold: false, italic: false, underline: false, strikethrough: false,
      fontSize: 10, color: null, fontFamily: 'Times New Roman', fontFamilyEastAsia: 'Times New Roman',
      isLink: false, background: null, vertAlign: null, hyperlink: null,
    }],
    defaultFontSize: 10, defaultFontFamily: 'Times New Roman', widowControl: false,
  } as unknown as DocParagraph;
}

function doc(paras: DocParagraph[]): DocxDocumentModel {
  return {
    section: {
      pageWidth: 400, pageHeight: 400,
      marginTop: 0, marginRight: 0, marginBottom: 0, marginLeft: 0,
      headerDistance: 0, footerDistance: 0, titlePage: false, evenAndOddHeaders: false,
    } as SectionProps,
    body: paras.map((p) => ({ type: 'paragraph', ...p })),
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    fontFamilyClasses: { 'Times New Roman': 'roman' },
  } as unknown as DocxDocumentModel;
}

describe('list marker lvlJc (§17.9.8)', () => {
  // firstLineX = indentLeft + firstLine = 57.6 + (−18) = 39.6 (margins 0).
  const FIRST_LINE_X = 39.6;

  it('right: different-width markers share a right edge (period-aligned)', async () => {
    const { canvas, fills } = makeRecordingCanvas();
    await renderDocumentToCanvas(doc([romanItem('i.', 'One', 'right'), romanItem('iii.', 'Three', 'right')]), canvas, 0, { dpr: 1, width: 400 });
    const mi = fills.find((f) => f.text === 'i.');
    const miii = fills.find((f) => f.text === 'iii.');
    expect(mi, 'marker "i." drawn').toBeDefined();
    expect(miii, 'marker "iii." drawn').toBeDefined();
    // Right edge = x + glyphs × 10.
    expect(mi!.x + 2 * 10).toBeCloseTo(FIRST_LINE_X, 3);
    expect(miii!.x + 4 * 10).toBeCloseTo(FIRST_LINE_X, 3);
    // Narrow "i." therefore starts to the RIGHT of wide "iii.".
    expect(mi!.x).toBeGreaterThan(miii!.x);
    // Body stays at indentLeft (the right-aligned marker fits the hanging indent).
    const body = fills.find((f) => f.text === 'One');
    expect(body!.x).toBeCloseTo(57.6, 3);
  });

  it('left (default): markers share a LEFT edge', async () => {
    const { canvas, fills } = makeRecordingCanvas();
    await renderDocumentToCanvas(doc([romanItem('i.', 'One', 'left'), romanItem('iii.', 'Three', 'left')]), canvas, 0, { dpr: 1, width: 400 });
    const mi = fills.find((f) => f.text === 'i.');
    const miii = fills.find((f) => f.text === 'iii.');
    expect(mi!.x).toBeCloseTo(FIRST_LINE_X, 3);
    expect(miii!.x).toBeCloseTo(FIRST_LINE_X, 3);
  });
});
