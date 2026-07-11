import { describe, it, expect } from 'vitest';
import { renderDocumentToCanvas } from './renderer.js';
import type { DocParagraph, DocxDocumentModel, SectionProps } from './types.js';

// ECMA-376 §17.3.1.12 (hanging) + §17.3.1.38 (a hanging indent implicitly creates
// a tab stop at indentLeft) + §17.9.28 (`<w:suff>`, default "tab"): in a
// hanging-indent numbered paragraph the marker sits in the hanging margin and the
// suff=tab that follows it advances the BODY to the indentLeft tab stop, so the
// first line's text region matches the continuation lines' region. This holds
// regardless of the paragraph's base direction — §17.3.1.6 makes w:ind (and its
// hanging first-line component) logical under w:bidi, so a `w:bidi` (RTL) paragraph
// mirrors the whole construction to the physical RIGHT: the body's start (right)
// edge lands at the paragraph's start indent, and the marker's right edge sits
// `w:hanging` further out.
//
// Regression: RTL let the raw negative first-line indent (−hanging) flow into the
// first line, widening it by `hanging`, so the body's right edge — and the marker
// beyond it — were pushed one `hanging` too far toward the margin.
//
// The recording canvas measures each glyph at fontSize px, so widths are exact and
// the drawn x of the body and the marker can be pinned.

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

// A single RTL word — no spaces so it is one segment; the recording canvas ignores
// shaping and measures it as codePoints × fontSize.
const RTL_BODY = 'سلام'; // 4 code points → 40 px at fontSize 10

function bulletItem(overrides: Partial<DocParagraph> = {}): DocParagraph {
  return {
    alignment: 'left', // start-aligned; under RTL this is the physical RIGHT edge
    bidi: true,
    indentLeft: 57.6, indentRight: 0, indentFirst: -18, // hanging 18 pt
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null, tabStops: [],
    numbering: {
      numId: 3, level: 0, format: 'bullet', text: '•', indentLeft: 57.6,
      tab: 18, suff: 'tab', jc: 'left', fontFamily: 'Times New Roman', fontFamilyEastAsia: 'Times New Roman',
    },
    runs: [{
      type: 'text', text: RTL_BODY, bold: false, italic: false, underline: false, strikethrough: false,
      fontSize: 10, color: null, fontFamily: 'Times New Roman', fontFamilyEastAsia: 'Times New Roman',
      isLink: false, background: null, vertAlign: null, hyperlink: null,
    }],
    defaultFontSize: 10, defaultFontFamily: 'Times New Roman', widowControl: false,
    ...overrides,
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

describe('RTL numbered hanging indent (§17.9.28 suff=tab)', () => {
  // Page 400, margins 0 ⇒ contentX = 0, contentW = 400.
  // RTL swaps physical indents: physicalLeft = indentRight = 0,
  // physicalRight = indentLeft = 57.6. paraX = 0, paraW = 400 − 57.6 = 342.4.
  // The start (right) edge of the text region — where continuation lines end and
  // where the first-line body must also end — is paraX + paraW = 342.4.
  const START_EDGE = 342.4;
  const HANGING = 18; // = numbering.tab
  const BODY_W = 40; // 4 code points × 10 px

  it('first-line body ends at the start indent, marker hangs one w:hanging beyond', async () => {
    const { canvas, fills } = makeRecordingCanvas();
    await renderDocumentToCanvas(doc([bulletItem()]), canvas, 0, { dpr: 1, width: 400 });

    const bodyFill = fills.find((f) => f.text === RTL_BODY);
    const markerFill = fills.find((f) => f.text === '•');
    expect(bodyFill, 'RTL body drawn').toBeDefined();
    expect(markerFill, 'bullet marker drawn').toBeDefined();

    // The body's right edge sits AT the paragraph's start indent — it does NOT
    // hang out by `hanging`. (Pre-fix it landed at START_EDGE + HANGING.)
    const bodyRightEdge = bodyFill!.x + BODY_W;
    expect(bodyRightEdge).toBeCloseTo(START_EDGE, 2);

    // The marker's right edge sits exactly `w:hanging` past the body's start edge
    // (the hanging margin the marker+tab occupy).
    const markerRightEdge = markerFill!.x + 10; // 1 code point × 10 px
    expect(markerRightEdge).toBeCloseTo(START_EDGE + HANGING, 2);
  });
});
