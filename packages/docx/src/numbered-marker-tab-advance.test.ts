import { describe, it, expect } from 'vitest';
import { renderDocumentToCanvas } from './renderer.js';
import type { DocParagraph, DocxDocumentModel, SectionProps } from './types.js';

// ECMA-376 §17.9.6 (suff="tab") + §17.3.1.37 (a tab never moves backward): a
// numbered paragraph's marker is followed by a tab that advances the body to the
// numbering's indentLeft stop — UNLESS the marker overruns that stop (a wide
// multi-level number like "1.1.1." whose glyphs exceed the hanging indent), in
// which case the tab advances to the next stop PAST the marker and the body
// follows. sample-11's "1.1.1. Three" collided because we pinned the body at
// indentLeft; Word advances "Three" to the next default tab.
//
// Recording canvas measures every glyph at fontSize px, so a 4-glyph "1.1."
// marker is 40 px — wider than the 18 px hanging indent here — forcing the
// overrun path. Default tab grid = 36 pt.

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

/** A 3rd-level numbered paragraph: marker "1.1." (overruns), body "Body". */
function numberedPara(markerText: string, indentLeftPt: number, hangingPt: number): DocParagraph {
  return {
    alignment: 'left', indentLeft: indentLeftPt, indentRight: 0, indentFirst: -hangingPt,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null, tabStops: [],
    numbering: {
      numId: 1, level: 2, format: 'decimal', text: markerText, indentLeft: indentLeftPt,
      tab: hangingPt, suff: 'tab', fontFamily: 'Times New Roman', fontFamilyEastAsia: 'Times New Roman',
    },
    runs: [{
      type: 'text', text: 'Body', bold: false, italic: false, underline: false, strikethrough: false,
      fontSize: 10, color: null, fontFamily: 'Times New Roman', fontFamilyEastAsia: 'Times New Roman',
      isLink: false, background: null, vertAlign: null, hyperlink: null,
    }],
    defaultFontSize: 10, defaultFontFamily: 'Times New Roman', widowControl: false,
  } as unknown as DocParagraph;
}

function doc(p: DocParagraph): DocxDocumentModel {
  return {
    section: {
      pageWidth: 400, pageHeight: 400,
      marginTop: 0, marginRight: 0, marginBottom: 0, marginLeft: 0,
      headerDistance: 0, footerDistance: 0, titlePage: false, evenAndOddHeaders: false,
    } as SectionProps,
    body: [{ type: 'paragraph', ...p }],
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    fontFamilyClasses: { 'Times New Roman': 'roman' },
  } as unknown as DocxDocumentModel;
}

describe('numbered list — body clears an over-wide marker (§17.9.6 / §17.3.1.37)', () => {
  it('advances the body past the marker to the next tab stop, not onto indentLeft', async () => {
    const { canvas, fills } = makeRecordingCanvas();
    // indentLeft 72 pt, hanging 18 pt ⇒ marker budget 18 px; "1.1." is 40 px and
    // overruns. Marker draws at paraX+indFirst = 72−18 = 54, ends at 94. The body
    // must advance to the next 36-pt tab past 94 ⇒ 108, NOT stay at 72.
    await renderDocumentToCanvas(doc(numberedPara('1.1.', 72, 18)), canvas, 0, { dpr: 1, width: 400 });
    const marker = fills.find((f) => f.text === '1.1.');
    const body = fills.find((f) => f.text === 'Body');
    expect(marker, 'marker drawn').toBeDefined();
    expect(body, 'body drawn').toBeDefined();
    const markerEnd = marker!.x + 4 * 10; // 4 glyphs × 10 px
    expect(marker!.x).toBeCloseTo(54, 3);
    expect(markerEnd).toBeCloseTo(94, 3);
    // Body advanced to the next default tab (108), clearing the marker.
    expect(body!.x).toBeCloseTo(108, 3);
    expect(body!.x, 'body must not overlap the marker').toBeGreaterThanOrEqual(markerEnd);
  });

  it('leaves the body at indentLeft when the marker fits the hanging indent', async () => {
    const { canvas, fills } = makeRecordingCanvas();
    // "1." is 20 px; hanging 36 px ⇒ fits. Body stays at indentLeft (72), the
    // normal suff=tab behavior (no overrun advance).
    await renderDocumentToCanvas(doc(numberedPara('1.', 72, 36)), canvas, 0, { dpr: 1, width: 400 });
    const body = fills.find((f) => f.text === 'Body');
    expect(body!.x).toBeCloseTo(72, 3);
  });
});
