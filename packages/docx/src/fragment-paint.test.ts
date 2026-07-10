import { describe, it, expect, beforeAll } from 'vitest';
import { renderDocumentToCanvas, paginateDocument } from './renderer.js';
import type { BodyElement, DocParagraph, DocxDocumentModel, PaginatedBodyElement, SectionProps } from './types';

// ─────────────────────────────────────────────────────────────────────────────
// PR 5 Task 13 — body fragment paint purity.
//
// A migrated body paragraph paints from its stored measured fragment
// (fragment-paint.ts). At the paint scale of 1 the stored scale-1 geometry needs no
// rescale, so the paint pass must draw the paragraph's lines WITHOUT calling
// measureText at all — no line layout, no segment measurement, no remeasurement.
//
// This is proved end-to-end through the real production flow: pages are paginated
// with a normal OffscreenCanvas metric, then painted at scale 1 onto a canvas whose
// measureText THROWS. If any part of the migrated paragraph paint tried to measure,
// the render would throw; instead it completes and draws the paragraph text.
// ─────────────────────────────────────────────────────────────────────────────

interface Call { text: string; x: number; y: number; }

/** Pagination-side canvas with a normal linear glyph metric. */
function makeMeasuringCtx(): CanvasRenderingContext2D {
  let font = '10px serif';
  const ctx = {
    get font() { return font; },
    set font(v: string) { font = v; },
    letterSpacing: '0px',
    measureText: (s: string) => {
      const p = parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
      const per = p * 0.5;
      return {
        width: [...s].length * per,
        fontBoundingBoxAscent: p * 0.8, fontBoundingBoxDescent: p * 0.2,
        actualBoundingBoxAscent: p * 0.8, actualBoundingBoxDescent: p * 0.2,
      } as TextMetrics;
    },
    save() {}, restore() {}, beginPath() {}, closePath() {}, moveTo() {}, lineTo() {},
    stroke() {}, fill() {}, fillRect() {}, strokeRect() {}, clip() {}, rect() {},
    scale() {}, translate() {}, rotate() {}, setLineDash() {}, clearRect() {}, arc() {},
    quadraticCurveTo() {}, bezierCurveTo() {}, createLinearGradient() { return { addColorStop() {} }; },
    drawImage() {}, fillText() {}, strokeText() {},
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    textAlign: 'left' as CanvasTextAlign, direction: 'ltr' as CanvasDirection,
  };
  return ctx as unknown as CanvasRenderingContext2D;
}

beforeAll(() => {
  (globalThis as unknown as { OffscreenCanvas: unknown }).OffscreenCanvas = class {
    getContext() { return makeMeasuringCtx(); }
  };
});

/** Paint-side recording canvas whose measureText THROWS — any measurement during
 *  paint fails the test loudly. Records every text draw for the content assertion. */
function makeThrowingPaintCanvas(): { canvas: HTMLCanvasElement; calls: Call[]; measured: () => number } {
  let font = '10px serif';
  const calls: Call[] = [];
  let measured = 0;
  const ctx = {
    get font() { return font; },
    set font(v: string) { font = v; },
    letterSpacing: '0px',
    measureText: (_s: string) => {
      measured++;
      throw new Error('measureText must not be called during fragment paint');
    },
    save() {}, restore() {}, beginPath() {}, closePath() {},
    moveTo() {}, lineTo() {}, stroke() {}, fill() {}, fillRect() {},
    strokeRect() {}, clip() {}, rect() {}, scale() {}, translate() {}, rotate() {},
    setLineDash() {}, clearRect() {}, arc() {}, quadraticCurveTo() {},
    bezierCurveTo() {}, createLinearGradient() { return { addColorStop() {} }; },
    drawImage() {},
    fillText(s: string, x: number, y: number) { calls.push({ text: s, x, y }); },
    strokeText(s: string, x: number, y: number) { calls.push({ text: s, x, y }); },
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    textAlign: 'left' as CanvasTextAlign, direction: 'ltr' as CanvasDirection,
    globalAlpha: 1, lineCap: 'butt' as CanvasLineCap, lineJoin: 'miter' as CanvasLineJoin,
  };
  const canvas = { width: 0, height: 0, style: {} as Record<string, string>, getContext: () => ctx };
  return { canvas: canvas as unknown as HTMLCanvasElement, calls, measured: () => measured };
}

function para(text: string, over: Partial<DocParagraph> = {}): DocParagraph {
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

function doc(body: BodyElement[], pageHeight = 400): DocxDocumentModel {
  const section: SectionProps = {
    pageWidth: 200, pageHeight,
    marginTop: 10, marginRight: 10, marginBottom: 10, marginLeft: 10,
    headerDistance: 4, footerDistance: 4, titlePage: false, evenAndOddHeaders: false,
    sectionStart: 'nextPage', columns: null,
  } as SectionProps;
  return {
    section, body,
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    fontFamilyClasses: { 'Times New Roman': 'roman' },
    footnotes: [],
  } as unknown as DocxDocumentModel;
}

describe('fragment paint purity (PR 5 Task 13)', () => {
  it('paints a premeasured body paragraph at scale 1 without ever calling measureText', async () => {
    const model = doc([para('hello world one two three') as unknown as BodyElement]);
    const pages = paginateDocument(model); // measured with the normal OffscreenCanvas
    const paint = makeThrowingPaintCanvas();

    // Paint scale 1 (render width == page width). A migrated paragraph draws its
    // stored fragment lines; nothing measures.
    await expect(
      renderDocumentToCanvas(model, paint.canvas, 0, { dpr: 1, width: 200, prebuiltPages: pages }),
    ).resolves.not.toThrow();

    expect(paint.measured()).toBe(0);
    // Non-vacuity: the paragraph's words were actually drawn.
    const drewText = paint.calls.some((c) => c.text.includes('hello'));
    expect(drewText).toBe(true);
  });

  it('paints a paragraph that SPLITS across pages from fragments, still measure-free', async () => {
    // A long paragraph over a short page splits; each continuation slice paints from
    // the shared measured fragment window without remeasuring.
    const long = Array.from({ length: 120 }, () => 'w').join(' ');
    const model = doc([para(long) as unknown as BodyElement], 60);
    const pages = paginateDocument(model);
    expect(pages.length).toBeGreaterThan(1);
    const split = pages.some((pg) => pg.some((el) => (el as PaginatedBodyElement).lineSlice));
    expect(split).toBe(true);

    for (let p = 0; p < pages.length; p++) {
      const paint = makeThrowingPaintCanvas();
      await expect(
        renderDocumentToCanvas(model, paint.canvas, p, { dpr: 1, width: 200, prebuiltPages: pages }),
      ).resolves.not.toThrow();
      expect(paint.measured()).toBe(0);
      expect(paint.calls.length).toBeGreaterThan(0);
    }
  });
});
