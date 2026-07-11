import { describe, expect, it } from 'vitest';
import { computePages, renderDocumentToCanvas } from './renderer.js';
import type {
  BodyElement,
  DocParagraph,
  DocxDocumentModel,
  SectionProps,
  ShapeRun,
  ShapeText,
} from './types';

interface FillTextEvent { text: string; x: number; y: number }

function makeRecordingCanvas(): { canvas: HTMLCanvasElement; fillTexts: FillTextEvent[] } {
  let font = '20px serif';
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '20');
  const fillTexts: FillTextEvent[] = [];
  const ctx = {
    get font() { return font; },
    set font(value: string) { font = value; },
    letterSpacing: '0px',
    measureText: (text: string) => {
      const size = px();
      return {
        width: [...text].length * size,
        fontBoundingBoxAscent: size * 0.8,
        fontBoundingBoxDescent: size * 0.2,
        actualBoundingBoxAscent: size * 0.8,
        actualBoundingBoxDescent: size * 0.2,
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

function para(runs: DocParagraph['runs'] = []): DocParagraph {
  return {
    alignment: 'left',
    indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null,
    numbering: null, tabStops: [], runs,
    defaultFontSize: 20, defaultFontFamily: 'Arial',
    widowControl: false,
  } as unknown as DocParagraph;
}

function dividerTextbox(): ShapeRun {
  const block: ShapeText = {
    text: 'Divider title',
    fontSizePt: 24,
    alignment: 'center',
    runs: [{ text: 'Divider title', fontSizePt: 24 }],
  };
  return {
    type: 'shape',
    zOrder: 0, subpaths: [], presetGeometry: 'rect',
    fill: null, stroke: null,
    behindDoc: false, wrapMode: 'none',
    widthPt: 120, heightPt: 50,
    anchorXPt: 0, anchorXFromMargin: false,
    anchorXRelativeFrom: 'page', anchorXAlign: 'center',
    anchorYPt: 5, anchorYFromPara: true, anchorYRelativeFrom: 'paragraph',
    textBlocks: [block], textAnchor: 'ctr',
    textInsetL: 0, textInsetT: 0, textInsetR: 0, textInsetB: 0,
  } as unknown as ShapeRun;
}

function inlineDividerBackground(): DocParagraph['runs'][number] {
  return {
    type: 'image',
    imagePath: 'word/media/divider.png', mimeType: 'image/png',
    widthPt: 160, heightPt: 50,
    anchor: false,
  } as DocParagraph['runs'][number];
}

function documentModel(body: BodyElement[]): DocxDocumentModel {
  return {
    section: {
      pageWidth: 200, pageHeight: 140,
      marginTop: 20, marginRight: 20, marginBottom: 20, marginLeft: 20,
      headerDistance: 0, footerDistance: 0,
      titlePage: false, evenAndOddHeaders: false,
      sectionStart: 'nextPage',
    } as SectionProps,
    body,
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    fontFamilyClasses: { Arial: 'swiss' },
  } as unknown as DocxDocumentModel;
}

describe('anchored text box on a section-mark paragraph', () => {
  it('draws the text box on the overflow page before the next-page section starts', async () => {
    // ECMA-376 Part 1 §17.6.18 stores a non-final section's sectPr in its last
    // paragraph; §17.6.22 makes the following section start on the next page.
    // The same paragraph can still contain the §20.4.2.3 floating object whose
    // §20.4.2.10/.11 position is page-centered and paragraph-relative, with
    // §20.4.2.15 wrapNone. Model that paragraph immediately before the marker.
    const sectionMarkParagraph = {
      type: 'paragraph',
      ...para([
        dividerTextbox() as unknown as DocParagraph['runs'][number],
        inlineDividerBackground(),
      ]),
    } as BodyElement;
    const body: BodyElement[] = [
      ...Array.from({ length: 4 }, () => ({ type: 'paragraph', ...para() }) as BodyElement),
      sectionMarkParagraph,
      { type: 'sectionBreak', kind: 'nextPage', columns: null } as BodyElement,
      { type: 'paragraph', ...para() } as BodyElement,
    ];
    const { canvas, fillTexts } = makeRecordingCanvas();
    const model = documentModel(body);
    const context = canvas.getContext('2d') as CanvasRenderingContext2D;
    const pages = computePages(body, model.section, context, model.fontFamilyClasses);

    expect(pages).toHaveLength(3);
    await renderDocumentToCanvas(model, canvas, 1, {
      dpr: 1,
      width: 200,
      prebuiltPages: pages,
    });

    expect(fillTexts.map((event) => event.text).join('')).toContain('Divider title');

    // The anchor belongs to the section-mark paragraph on the overflow page
    // (§20.4.2.3: the object is positioned relative to its anchor paragraph),
    // so neither the first page of the section nor the first page of the next
    // section may paint it.
    for (const otherPage of [0, 2]) {
      const other = makeRecordingCanvas();
      await renderDocumentToCanvas(model, other.canvas, otherPage, {
        dpr: 1,
        width: 200,
        prebuiltPages: pages,
      });
      expect(other.fillTexts.map((event) => event.text).join('')).not.toContain('Divider title');
    }
  });
});
