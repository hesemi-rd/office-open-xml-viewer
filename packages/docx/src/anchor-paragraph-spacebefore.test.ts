import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import { renderDocumentToCanvas } from './renderer.js';
import type { BodyElement, DocParagraph, DocxDocumentModel, ImageRun, SectionProps } from './types';

// ECMA-376 §20.4.3.5 ST_RelFromV "paragraph": a `<wp:positionV relativeFrom=
// "paragraph">` float is positioned "relative to the paragraph which contains the
// drawing anchor" — its TOP edge, BEFORE the paragraph's spaceBefore. Word anchors
// the float at the paragraph top, NOT at the post-spaceBefore text area, and makes
// no distinction between wrap modes. The renderer previously anchored WRAP floats
// (square/tight/topAndBottom) at the post-spaceBefore text top while wrapNone floats
// used the pre-spaceBefore top — so a wrap float on a paragraph with spaceBefore sat
// `spaceBefore` pt too low (sample-12's figure: anchor paragraph spaceBefore=12 pt,
// the image dropped 12 pt and ate the whitespace above its caption). This pins the
// drawn float Y at the paragraph's pre-spaceBefore top for a wrapSquare image.

interface DrawCall { y: number; }

function makeRecordingCanvas(): { canvas: HTMLCanvasElement; draws: DrawCall[] } {
  let font = '10px serif';
  const draws: DrawCall[] = [];
  const ctx = {
    get font() { return font; },
    set font(v: string) { font = v; },
    letterSpacing: '0px',
    measureText: (s: string) => {
      const p = parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
      return {
        width: [...s].length * p,
        fontBoundingBoxAscent: p * 0.8, fontBoundingBoxDescent: p * 0.2,
        actualBoundingBoxAscent: p * 0.8, actualBoundingBoxDescent: p * 0.2,
      } as TextMetrics;
    },
    save() {}, restore() {}, beginPath() {}, closePath() {},
    moveTo() {}, lineTo() {}, stroke() {}, fill() {}, fillRect() {},
    strokeRect() {}, clip() {}, rect() {}, scale() {}, translate() {},
    setLineDash() {}, clearRect() {}, arc() {}, quadraticCurveTo() {},
    bezierCurveTo() {}, createLinearGradient() { return { addColorStop() {} }; },
    // No srcRect ⇒ 5-arg form drawImage(img, dx, dy, dw, dh): dy is the 3rd arg.
    drawImage(_img: unknown, _dx: number, dy: number) { draws.push({ y: dy }); },
    fillText() {}, strokeText() {},
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    textAlign: 'left' as CanvasTextAlign, direction: 'ltr' as CanvasDirection,
    globalAlpha: 1, lineCap: 'butt' as CanvasLineCap, lineJoin: 'miter' as CanvasLineJoin,
  };
  const canvas = { width: 0, height: 0, style: {} as Record<string, string>, getContext: () => ctx };
  return { canvas: canvas as unknown as HTMLCanvasElement, draws };
}

function squareFloat(): ImageRun {
  return {
    type: 'image', imagePath: 'word/media/image1.png', mimeType: 'image/png',
    widthPt: 80, heightPt: 40,
    anchor: true, anchorYFromPara: true, anchorYRelativeFrom: 'paragraph',
    anchorYPt: 0, anchorXPt: 0, anchorXFromMargin: true, anchorXRelativeFrom: 'margin',
    wrapMode: 'square', wrapSide: 'bothSides',
  } as unknown as ImageRun;
}

function paraWithFloat(spaceBefore: number): DocParagraph {
  return {
    type: 'paragraph', alignment: 'left',
    indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore, spaceAfter: 0, lineSpacing: null,
    numbering: null, tabStops: [],
    runs: [{ ...squareFloat() } as DocParagraph['runs'][number],
           { type: 'text', text: 'x', bold: false, italic: false, underline: false,
             strikethrough: false, fontSize: 11, color: null, fontFamily: 'Times New Roman',
             fontFamilyEastAsia: '', isLink: false, background: null, vertAlign: null,
             hyperlink: null } as DocParagraph['runs'][number]],
    defaultFontSize: 11, defaultFontFamily: 'Times New Roman', widowControl: false,
  } as unknown as DocParagraph;
}

function docOf(p: DocParagraph): DocxDocumentModel {
  return {
    section: {
      pageWidth: 400, pageHeight: 800,
      marginTop: 0, marginRight: 0, marginBottom: 0, marginLeft: 0,
      headerDistance: 0, footerDistance: 0, titlePage: false, evenAndOddHeaders: false,
    } as SectionProps,
    body: [p as unknown as BodyElement],
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    fontFamilyClasses: { 'Times New Roman': 'roman' },
  } as unknown as DocxDocumentModel;
}

describe('paragraph-anchored wrap float Y (§20.4.3.5)', () => {
  beforeEach(() => {
    vi.stubGlobal(
      'createImageBitmap',
      vi.fn(async () => ({ width: 2, height: 2, close: () => {} }) as unknown as ImageBitmap),
    );
  });
  afterEach(() => vi.unstubAllGlobals());

  it('anchors a wrapSquare paragraph float at the pre-spaceBefore paragraph top', async () => {
    const { canvas, draws } = makeRecordingCanvas();
    await renderDocumentToCanvas(docOf(paraWithFloat(24)), canvas, 0, {
      dpr: 1, width: 400, // scale = 1 (px per pt)
      fetchImage: async (path: string, mime: string) => new Blob([new Uint8Array([1])], { type: mime }),
    });
    // First paragraph, marginTop=0, anchorYPt=0. Pre-spaceBefore top = 0; the
    // float must draw at y≈0, NOT y≈24 (the post-spaceBefore text area).
    expect(draws.length).toBeGreaterThan(0);
    expect(draws[0].y).toBeCloseTo(0, 2);
  });
});
