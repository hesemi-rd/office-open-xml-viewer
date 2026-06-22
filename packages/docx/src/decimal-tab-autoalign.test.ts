import { describe, it, expect } from 'vitest';
import { renderDocumentToCanvas } from './renderer.js';
import type { DocParagraph, DocxDocumentModel, SectionProps } from './types.js';

// Word runtime behavior (NOT in ECMA-376 §17.3.1.37 / §17.18.84): a paragraph
// whose first tab stop is a `decimal` tab and whose content is a bare number
// with NO explicit tab character still aligns that number to the decimal tab —
// the built-in "Decimal Aligned" style on table number cells. sample-11's
// College table: 110 / 24 / 998 right-align on the decimal tab even though none
// carries a tab character. This test pins the auto-alignment: two numbers of
// different digit counts must share the same right edge (the decimal tab), and a
// number with NO decimal tab must stay left-aligned (the gate is specific).

interface FillCall {
  text: string;
  x: number;
}

function makeRecordingCanvas(): { canvas: HTMLCanvasElement; fills: FillCall[] } {
  let font = '10px serif';
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
  const fills: FillCall[] = [];
  const ctx = {
    get font() {
      return font;
    },
    set font(v: string) {
      font = v;
    },
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

function numberPara(text: string, withDecimalTab: boolean): DocParagraph {
  return {
    alignment: 'left',
    indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null,
    numbering: null,
    tabStops: withDecimalTab ? [{ pos: 36, alignment: 'decimal', leader: 'none' }] : [],
    runs: [{
      type: 'text', text,
      bold: false, italic: false, underline: false, strikethrough: false,
      fontSize: 10, color: null, fontFamily: 'Times New Roman', fontFamilyEastAsia: 'Times New Roman',
      isLink: false, background: null, vertAlign: null, hyperlink: null,
    }],
    defaultFontSize: 10, defaultFontFamily: 'Times New Roman', widowControl: false,
  } as unknown as DocParagraph;
}

function doc(paras: DocParagraph[]): DocxDocumentModel {
  return {
    section: {
      pageWidth: 300, pageHeight: 400,
      marginTop: 0, marginRight: 0, marginBottom: 0, marginLeft: 0,
      headerDistance: 0, footerDistance: 0, titlePage: false, evenAndOddHeaders: false,
    } as SectionProps,
    body: paras.map((p) => ({ type: 'paragraph', ...p })),
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    fontFamilyClasses: { 'Times New Roman': 'roman' },
  } as unknown as DocxDocumentModel;
}

describe('decimal-tab auto-alignment for tab-less numbers (Word behavior)', () => {
  const TAB_POS = 36; // pt; scale = 1 px/pt (width 300, pageWidth 300)
  const FS = 10; // each glyph is FS px wide in the recording canvas

  it('right-aligns numbers of different digit counts at the decimal tab', async () => {
    const { canvas, fills } = makeRecordingCanvas();
    await renderDocumentToCanvas(doc([numberPara('110', true), numberPara('24', true)]), canvas, 0, { dpr: 1, width: 300 });

    const f110 = fills.find((f) => f.text === '110');
    const f24 = fills.find((f) => f.text === '24');
    expect(f110, '"110" must be drawn').toBeDefined();
    expect(f24, '"24" must be drawn').toBeDefined();
    // Right edge = x + glyphCount * FS. Both must land on the decimal tab (36 pt).
    expect(f110!.x + 3 * FS).toBeCloseTo(TAB_POS, 4);
    expect(f24!.x + 2 * FS).toBeCloseTo(TAB_POS, 4);
    // 2-digit "24" therefore starts to the RIGHT of 3-digit "110".
    expect(f24!.x).toBeGreaterThan(f110!.x);
  });

  it('leaves a number with no decimal tab left-aligned (gate is specific)', async () => {
    const { canvas, fills } = makeRecordingCanvas();
    await renderDocumentToCanvas(doc([numberPara('110', false), numberPara('24', false)]), canvas, 0, { dpr: 1, width: 300 });
    const f110 = fills.find((f) => f.text === '110');
    const f24 = fills.find((f) => f.text === '24');
    // No decimal tab ⇒ both left-align at the margin (x = 0), right edges differ.
    expect(f110!.x).toBeCloseTo(0, 4);
    expect(f24!.x).toBeCloseTo(0, 4);
  });
});
