import { describe, it, expect } from 'vitest';
import { renderDocumentToCanvas } from './renderer.js';
import type { BodyElement, DocParagraph, DocxDocumentModel, SectionProps, ShapeRun } from './types';

// ECMA-376 §20.4.2.10 — a wp:anchor with behindDoc="0" floats IN FRONT of the
// inline text/image flow. Since the flow is painted in document order, a
// front-anchored shape in an EARLY paragraph must NOT be overpainted by a LATER
// paragraph's content (sample-13: the "Journal homepage" text box, anchored to
// the first paragraph, was hidden behind the inline masthead banner that
// follows it). The renderer defers front floats to a per-page top layer; this
// test pins that ordering by recording the sequence of fillText calls.

function makeRecordingCanvas(): { canvas: HTMLCanvasElement; texts: string[] } {
  let font = '10px serif';
  const texts: string[] = [];
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
    fillText(s: string) { texts.push(s); },
    strokeText(s: string) { texts.push(s); },
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    textAlign: 'left' as CanvasTextAlign, direction: 'ltr' as CanvasDirection,
    globalAlpha: 1, lineCap: 'butt' as CanvasLineCap, lineJoin: 'miter' as CanvasLineJoin,
  };
  const canvas = { width: 0, height: 0, style: {} as Record<string, string>, getContext: () => ctx };
  return { canvas: canvas as unknown as HTMLCanvasElement, texts };
}

function textPara(text: string): DocParagraph {
  return {
    type: 'paragraph', alignment: 'left',
    indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null,
    numbering: null, tabStops: [],
    runs: [{
      type: 'text', text, bold: false, italic: false, underline: false,
      strikethrough: false, fontSize: 11, color: null, fontFamily: 'Times New Roman',
      fontFamilyEastAsia: '', isLink: false, background: null, vertAlign: null, hyperlink: null,
    } as DocParagraph['runs'][number]],
    defaultFontSize: 11, defaultFontFamily: 'Times New Roman', widowControl: false,
  } as unknown as DocParagraph;
}

// A paragraph whose only run is a front-anchored (behindDoc unset) text box that
// overlaps the following paragraph's flow.
function shapePara(markerText: string): DocParagraph {
  const shape = {
    type: 'shape',
    widthPt: 200, heightPt: 40,
    anchorXPt: 0, anchorYPt: 0,
    anchorXFromMargin: false, anchorYFromPara: true,
    anchorXRelativeFrom: 'column', anchorYRelativeFrom: 'paragraph',
    presetGeometry: 'rect',
    wrapMode: 'none',
    textAnchor: 't',
    textInsetL: 0, textInsetT: 0, textInsetR: 0, textInsetB: 0,
    textBlocks: [{ text: markerText, fontSizePt: 11, fontFamily: 'Times New Roman', alignment: 'left' }],
  } as unknown as ShapeRun;
  return {
    type: 'paragraph', alignment: 'left',
    indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null,
    numbering: null, tabStops: [],
    runs: [shape as unknown as DocParagraph['runs'][number]],
    defaultFontSize: 11, defaultFontFamily: 'Times New Roman', widowControl: false,
  } as unknown as DocParagraph;
}

function zShape(markerText: string, zOrder: number): ShapeRun {
  return {
    type: 'shape',
    widthPt: 200, heightPt: 40,
    anchorXPt: 0, anchorYPt: 0,
    anchorXFromMargin: false, anchorYFromPara: true,
    anchorXRelativeFrom: 'column', anchorYRelativeFrom: 'paragraph',
    presetGeometry: 'rect',
    wrapMode: 'none',
    zOrder,
    subpaths: [],
    fill: { fillType: 'solid', color: 'FFFFFF' },
    stroke: null,
    textAnchor: 't',
    textInsetL: 0, textInsetT: 0, textInsetR: 0, textInsetB: 0,
    textBlocks: [{ text: markerText, fontSizePt: 11, fontFamily: 'Times New Roman', alignment: 'left' }],
  } as unknown as ShapeRun;
}

function twoShapePara(): DocParagraph {
  return {
    type: 'paragraph', alignment: 'left',
    indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null,
    numbering: null, tabStops: [],
    runs: [
      zShape('HIGH_Z', 20) as unknown as DocParagraph['runs'][number],
      zShape('LOW_Z', 10) as unknown as DocParagraph['runs'][number],
    ],
    defaultFontSize: 11, defaultFontFamily: 'Times New Roman', widowControl: false,
  } as unknown as DocParagraph;
}

describe('front float z-order (§20.4.2.10)', () => {
  it('a front-anchored shape paints ON TOP of a following paragraph (after it)', async () => {
    const section: SectionProps = {
      pageWidth: 400, pageHeight: 600,
      marginTop: 10, marginRight: 10, marginBottom: 10, marginLeft: 10,
      headerDistance: 4, footerDistance: 4, titlePage: false, evenAndOddHeaders: false,
    } as SectionProps;
    const doc = {
      section,
      body: [
        shapePara('SHAPE_FRONT') as unknown as BodyElement,
        textPara('LATER_BODY') as unknown as BodyElement,
      ],
      headers: { default: null, first: null, even: null },
      footers: { default: null, first: null, even: null },
      fontFamilyClasses: { 'Times New Roman': 'roman' },
    } as unknown as DocxDocumentModel;

    const { canvas, texts } = makeRecordingCanvas();
    await renderDocumentToCanvas(doc, canvas, 0, { dpr: 1, width: 400 });

    const bodyIdx = texts.indexOf('LATER_BODY');
    const shapeIdx = texts.indexOf('SHAPE_FRONT');
    expect(bodyIdx).toBeGreaterThanOrEqual(0);
    expect(shapeIdx).toBeGreaterThanOrEqual(0);
    // The front shape, anchored to the FIRST paragraph, must be painted AFTER the
    // later body text so it lands on top (deferred to the page's front layer).
    expect(shapeIdx).toBeGreaterThan(bodyIdx);
  });

  it('orders front shapes by wp:anchor relativeHeight, not paragraph run order', async () => {
    const section: SectionProps = {
      pageWidth: 400, pageHeight: 600,
      marginTop: 10, marginRight: 10, marginBottom: 10, marginLeft: 10,
      headerDistance: 4, footerDistance: 4, titlePage: false, evenAndOddHeaders: false,
    } as SectionProps;
    const doc = {
      section,
      body: [twoShapePara() as unknown as BodyElement],
      headers: { default: null, first: null, even: null },
      footers: { default: null, first: null, even: null },
      fontFamilyClasses: { 'Times New Roman': 'roman' },
    } as unknown as DocxDocumentModel;

    const { canvas, texts } = makeRecordingCanvas();
    await renderDocumentToCanvas(doc, canvas, 0, { dpr: 1, width: 400 });

    expect(texts.indexOf('LOW_Z')).toBeGreaterThanOrEqual(0);
    expect(texts.indexOf('HIGH_Z')).toBeGreaterThanOrEqual(0);
    expect(texts.indexOf('HIGH_Z')).toBeGreaterThan(texts.indexOf('LOW_Z'));
  });
});
