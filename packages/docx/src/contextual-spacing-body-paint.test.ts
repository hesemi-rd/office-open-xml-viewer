import { describe, it, expect } from 'vitest';
import { paginateDocument, renderDocumentToCanvas } from './renderer.js';
import type { BodyElement, DocParagraph, DocxDocumentModel, SectionProps } from './types';

// ECMA-376 §17.3.1.9 contextualSpacing — BODY paginator + paint integration pin
// for the Word-adjudicated per-side semantics (issue #1015, sample-57 ground
// truth). The unit tests cover the shared kernel through the cell/text-box
// helpers; this test drives the real body path end-to-end (paginateDocument →
// prebuiltPages → renderDocumentToCanvas) so a regression in EITHER the
// paginator's gap arithmetic or the paint pass's mirrored recompute (they must
// stay in lockstep) moves a painted baseline and fails here.
//
// Measurement is differential, like the sample-57 fixture: a control document
// whose two paragraphs carry zero spacing yields the pure single-line pitch L;
// each case document's baseline delta minus L is the inter-paragraph gap in pt
// (scale 1: canvas width == page width).
//
// Adjudicated table (a = prev.spaceAfter, b = curr.spaceBefore):
//   no toggle             → max(a, b)
//   prev-only toggle      → max(b − a, 0)
//   curr-only toggle      → a
//   both toggle           → 0
//   different styles      → max(a, b) (no suppression)

interface Call { text: string; y: number; }

/** Minimal recording canvas with LINEAR, scale-proportional metrics. */
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

(globalThis as unknown as { OffscreenCanvas: unknown }).OffscreenCanvas = class {
  getContext() { return makeRecordingCanvas().canvas.getContext('2d'); }
};

function para(
  text: string,
  over: Partial<DocParagraph> & { styleId?: string | null },
): DocParagraph {
  return {
    type: 'paragraph', alignment: 'left',
    indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null,
    numbering: null, tabStops: [],
    runs: [{
      type: 'text', text, bold: false, italic: false, underline: false,
      strikethrough: false, fontSize: 10, color: null, fontFamily: 'Times New Roman',
      fontFamilyEastAsia: '', isLink: false, background: null, vertAlign: null, hyperlink: null,
    } as DocParagraph['runs'][number]],
    defaultFontSize: 10, defaultFontFamily: 'Times New Roman', widowControl: false,
    ...over,
  } as unknown as DocParagraph;
}

function doc(body: BodyElement[]): DocxDocumentModel {
  const section: SectionProps = {
    pageWidth: 400, pageHeight: 400,
    marginTop: 10, marginRight: 10, marginBottom: 10, marginLeft: 10,
    headerDistance: 4, footerDistance: 4, titlePage: false, evenAndOddHeaders: false,
    sectionStart: 'nextPage', columns: null,
  } as SectionProps;
  return {
    section,
    body,
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    fontFamilyClasses: { 'Times New Roman': 'roman' },
    footnotes: [],
  } as unknown as DocxDocumentModel;
}

/** Paint a two-paragraph body at scale 1 and return yCURR − yPREV (baseline pt). */
async function baselineDelta(prev: DocParagraph, curr: DocParagraph): Promise<number> {
  const model = doc([prev, curr] as unknown as BodyElement[]);
  const pages = paginateDocument(model);
  expect(pages.length).toBe(1);
  const { canvas, calls } = makeRecordingCanvas();
  await renderDocumentToCanvas(model, canvas, 0, { dpr: 1, width: 400, prebuiltPages: pages });
  const yPrev = calls.find((c) => c.text.includes('PREV'))?.y;
  const yCurr = calls.find((c) => c.text.includes('CURR'))?.y;
  expect(yPrev).toBeTypeOf('number');
  expect(yCurr).toBeTypeOf('number');
  return (yCurr as number) - (yPrev as number);
}

describe('§17.3.1.9 contextualSpacing — body paginate+paint per-side gaps (issue #1015)', () => {
  it('paints the adjudicated six-case gap table', async () => {
    // Control: zero spacing everywhere → pure single-line pitch L (gap 0).
    const L = await baselineDelta(
      para('PREV', { styleId: 'CtxPair' }),
      para('CURR', { styleId: 'CtxPair' }),
    );
    expect(L).toBeGreaterThan(0);

    const gap = async (
      prevOver: Partial<DocParagraph> & { styleId?: string | null },
      currOver: Partial<DocParagraph> & { styleId?: string | null },
    ): Promise<number> =>
      (await baselineDelta(para('PREV', prevOver), para('CURR', currOver))) - L;

    // Case 5 — no toggle: gap = max(10, 12) = 12.
    expect(await gap(
      { styleId: 'CtxPair', spaceAfter: 10 },
      { styleId: 'CtxPair', spaceBefore: 12 },
    )).toBeCloseTo(12, 5);

    // Case 1 — PREV-only toggle: gap = max(12 − 10, 0) = 2 (spec worked example).
    expect(await gap(
      { styleId: 'CtxPair', spaceAfter: 10, contextualSpacing: true },
      { styleId: 'CtxPair', spaceBefore: 12 },
    )).toBeCloseTo(2, 5);

    // Case 2 — CURR-only toggle: gap = prev.spaceAfter = 10 (Word-measured).
    expect(await gap(
      { styleId: 'CtxPair', spaceAfter: 10 },
      { styleId: 'CtxPair', spaceBefore: 12, contextualSpacing: true },
    )).toBeCloseTo(10, 5);

    // Case 3a — both toggle: gap = 0.
    expect(await gap(
      { styleId: 'CtxPair', spaceAfter: 10, contextualSpacing: true },
      { styleId: 'CtxPair', spaceBefore: 12, contextualSpacing: true },
    )).toBeCloseTo(0, 5);

    // Case 3b — both toggle, asymmetric 4/12: still 0.
    expect(await gap(
      { styleId: 'CtxPair', spaceAfter: 4, contextualSpacing: true },
      { styleId: 'CtxPair', spaceBefore: 12, contextualSpacing: true },
    )).toBeCloseTo(0, 5);

    // Case 4 — both toggle, DIFFERENT styles: no suppression, gap = 12.
    expect(await gap(
      { styleId: 'CtxPair', spaceAfter: 10, contextualSpacing: true },
      { styleId: 'CtxOther', spaceBefore: 12, contextualSpacing: true },
    )).toBeCloseTo(12, 5);
  });
});
