import { describe, it, expect } from 'vitest';
import { renderDocumentToCanvas } from './renderer.js';
import type {
  BodyElement,
  DocParagraph,
  DocxDocumentModel,
  SectionProps,
  ShapeRun,
  ShapeText,
} from './types';

// sample-13's journal masthead "Journal homepage: https://…" lives inside a
// DrawingML anchored text box (`<mc:Choice Requires="wps"><wp:anchor><wps:wsp
// txBox=1><wps:txbx>`), noFill + noLine, anchored to the FIRST body paragraph
// (positionV relativeFrom="paragraph", wrapNone, behindDoc=0). The shape carries
// no panel, so if its txbx text is not drawn the whole line vanishes. The parser
// emits it as a `{type:'shape'}` run with `textBlocks`; `renderShapeText` draws
// such blocks (renderer.textbox-image.test.ts). This test exercises the
// END-TO-END pipeline (renderDocumentToCanvas → renderAnchorImagesAndShapes →
// renderAnchorShape → renderShapeText) to verify the anchored text box on the
// first paragraph actually reaches the draw.

interface FillTextEvent { text: string; x: number; y: number }

function makeRecordingCanvas(): { canvas: HTMLCanvasElement; fillTexts: FillTextEvent[] } {
  let font = '11px serif';
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '11');
  const fillTexts: FillTextEvent[] = [];
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
    save() {}, restore() {}, beginPath() {}, closePath() {},
    moveTo() {}, lineTo() {}, stroke() {}, fill() {}, clip() {}, rect() {},
    scale() {}, translate() {}, rotate() {}, setLineDash() {}, drawImage() {}, clearRect() {},
    arc() {}, quadraticCurveTo() {}, bezierCurveTo() {},
    createLinearGradient() { return { addColorStop() {} }; },
    fillRect() {}, strokeRect() {},
    fillText(text: string, x: number, y: number) { fillTexts.push({ text, x, y }); },
    strokeText() {},
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    textAlign: 'left' as CanvasTextAlign, direction: 'ltr' as CanvasDirection,
    globalAlpha: 1, lineCap: 'butt' as CanvasLineCap, lineJoin: 'miter' as CanvasLineJoin,
  };
  const canvas = { width: 0, height: 0, style: {} as Record<string, string>, getContext: () => ctx };
  return { canvas: canvas as unknown as HTMLCanvasElement, fillTexts };
}

/** The sample-13 masthead text box: anchored, wrapNone, noFill/noLine, with the
 *  "Journal homepage: URL" line as a single rich-text block. */
function homepageTextbox(): ShapeRun {
  const block: ShapeText = {
    text: '     Journal homepage:  https://reference-global.com/journal/MSR',
    fontSizePt: 11,
    alignment: 'left',
    runs: [
      { text: '     ', fontSizePt: 11 },
      { text: 'Journal homepage:  ', fontSizePt: 11 },
      { text: 'https://reference-global.com/journal/MSR', fontSizePt: 11, color: '1F4E79' },
    ],
  };
  return {
    type: 'shape',
    zOrder: 0, subpaths: [], presetGeometry: 'rect',
    fill: null, stroke: null,
    behindDoc: false,
    wrapMode: 'none',
    widthPt: 494.35, heightPt: 74.85,
    anchorXPt: 2.7, anchorXFromMargin: false, anchorXRelativeFrom: 'column',
    anchorYPt: 11.5, anchorYFromPara: true, anchorYRelativeFrom: 'paragraph',
    textBlocks: [block], textAnchor: 't',
    textInsetL: 7.2, textInsetT: 3.6, textInsetR: 7.2, textInsetB: 3.6,
  } as unknown as ShapeRun;
}

function para(runs: DocParagraph['runs']): DocParagraph {
  return {
    alignment: 'left',
    indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null,
    numbering: null, tabStops: [],
    runs,
    defaultFontSize: 11, defaultFontFamily: 'Arial',
    widowControl: false,
  } as unknown as DocParagraph;
}

async function renderDoc(body: BodyElement[]): Promise<FillTextEvent[]> {
  const { canvas, fillTexts } = makeRecordingCanvas();
  const doc = {
    section: {
      pageWidth: 595, pageHeight: 842,
      marginTop: 70, marginRight: 47, marginBottom: 70, marginLeft: 47,
      headerDistance: 21, footerDistance: 21, titlePage: false, evenAndOddHeaders: false,
    } as SectionProps,
    body,
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    fontFamilyClasses: { Arial: 'swiss' },
  } as unknown as DocxDocumentModel;
  await renderDocumentToCanvas(doc, canvas, 0, { dpr: 1, width: 595 });
  return fillTexts;
}

describe('anchored text box on the first paragraph (sample-13 journal masthead)', () => {
  it('draws the txbx body text end-to-end (renderDocumentToCanvas)', async () => {
    // First paragraph carries ONLY the anchored masthead text box (its own text
    // flow is empty), exactly like sample-13's cover paragraph.
    const fillTexts = await renderDoc([
      { type: 'paragraph', ...para([homepageTextbox() as unknown as DocParagraph['runs'][number]]) } as BodyElement,
      { type: 'paragraph', ...para([{ type: 'text', text: 'Body follows.', bold: false, italic: false, underline: false, strikethrough: false, fontSize: 11, color: null, fontFamily: 'Arial', fontFamilyEastAsia: 'Arial', isLink: false, background: null, vertAlign: null, hyperlink: null } as DocParagraph['runs'][number]]) } as BodyElement,
    ]);
    const ink = fillTexts.map((c) => c.text).join('').replace(/\s/g, '');
    expect(ink).toContain('Journalhomepage:');
    expect(ink).toContain('https://reference-global.com/journal/MSR');
  });
});
