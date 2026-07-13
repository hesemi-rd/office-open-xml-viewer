import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { renderDocumentToCanvas } from './renderer.js';
import type {
  BodyElement,
  DocParagraph,
  DocxDocumentModel,
  NumberingInfo,
  ParagraphBorders,
  SectionProps,
} from './types.js';

// ECMA-376 Part 1 §17.9.7: lvlJc positions numbering text relative to the
// paragraph text margin; numbering text includes numerals, symbols and graphics.
// §17.3.1.24: pBdr applies to the paragraph. Therefore the resolved marker ink
// must remain inside the paragraph border even when right/center lvlJc shifts the
// marker left of the raw hanging-indent reference.

interface FillCall { text: string; x: number; }
interface VerticalStroke { x: number; width: number; color: string; }
interface ImageCall { x: number; w: number; }

function makeRecordingCanvas(): {
  canvas: HTMLCanvasElement;
  fills: FillCall[];
  verticalStrokes: VerticalStroke[];
  images: ImageCall[];
} {
  let font = '10px serif';
  let strokeStyle = '#000';
  let lineWidth = 1;
  let path: { x: number; y: number }[] = [];
  const fills: FillCall[] = [];
  const verticalStrokes: VerticalStroke[] = [];
  const images: ImageCall[] = [];
  const fontPx = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
  const ctx = {
    get font() { return font; },
    set font(value: string) { font = value; },
    get strokeStyle() { return strokeStyle; },
    set strokeStyle(value: string) { strokeStyle = value; },
    get lineWidth() { return lineWidth; },
    set lineWidth(value: number) { lineWidth = value; },
    letterSpacing: '0px',
    measureText(text: string) {
      const px = fontPx();
      return {
        width: [...text].length * px,
        fontBoundingBoxAscent: px * 0.8,
        fontBoundingBoxDescent: px * 0.2,
        actualBoundingBoxAscent: px * 0.8,
        actualBoundingBoxDescent: px * 0.2,
      } as TextMetrics;
    },
    save() {}, restore() {},
    beginPath() { path = []; }, closePath() {},
    moveTo(x: number, y: number) { path.push({ x, y }); },
    lineTo(x: number, y: number) { path.push({ x, y }); },
    stroke() {
      for (let i = 1; i < path.length; i++) {
        if (path[i - 1].x === path[i].x) {
          verticalStrokes.push({ x: path[i].x, width: lineWidth, color: strokeStyle });
        }
      }
    },
    fill() {}, fillRect() {}, strokeRect() {}, clip() {}, rect() {},
    scale() {}, translate() {}, rotate() {}, setLineDash() {}, clearRect() {}, arc() {},
    quadraticCurveTo() {}, bezierCurveTo() {},
    createLinearGradient() { return { addColorStop() {} }; },
    drawImage(_image: unknown, x: number, _y: number, w: number) { images.push({ x, w }); },
    fillText(text: string, x: number) { fills.push({ text, x }); },
    strokeText() {},
    fillStyle: '#000',
    textAlign: 'left' as CanvasTextAlign,
    direction: 'ltr' as CanvasDirection,
    globalAlpha: 1,
    lineCap: 'butt' as CanvasLineCap,
    lineJoin: 'miter' as CanvasLineJoin,
  };
  const canvas = {
    width: 0,
    height: 0,
    style: {} as Record<string, string>,
    getContext: () => ctx,
  };
  (ctx as unknown as { canvas: unknown }).canvas = canvas;
  return { canvas: canvas as unknown as HTMLCanvasElement, fills, verticalStrokes, images };
}

const ONE_BY_ONE_GIF = new Uint8Array([
  0x47, 0x49, 0x46, 0x38, 0x39, 0x61, 0x01, 0x00, 0x01, 0x00, 0x80, 0x00, 0x00,
  0xff, 0xff, 0xff, 0x00, 0x00, 0x00, 0x21, 0xf9, 0x04, 0x01, 0x00, 0x00, 0x00,
  0x00, 0x2c, 0x00, 0x00, 0x00, 0x00, 0x01, 0x00, 0x01, 0x00, 0x00, 0x02, 0x02,
  0x44, 0x01, 0x00, 0x3b,
]);

const BORDER_COLOR = '123456';
const borderEdge = () => ({ style: 'single', color: BORDER_COLOR, width: 1, space: 0 });
const borders = (): ParagraphBorders => ({
  top: borderEdge(),
  bottom: borderEdge(),
  left: borderEdge(),
  right: borderEdge(),
  between: null,
});

function numberedParagraph(jc: 'right' | 'center'): DocParagraph {
  const numbering: NumberingInfo = {
    numId: 1,
    level: 0,
    format: 'decimal',
    text: '1.',
    indentLeft: 72,
    tab: 18,
    suff: 'tab',
    jc,
    fontFamily: 'Times New Roman',
  };
  return {
    type: 'paragraph',
    alignment: 'left',
    indentLeft: 72,
    indentRight: 0,
    indentFirst: -18,
    spaceBefore: 0,
    spaceAfter: 0,
    lineSpacing: null,
    numbering,
    tabStops: [],
    borders: borders(),
    runs: [{
      type: 'text',
      text: 'Body',
      bold: false,
      italic: false,
      underline: false,
      strikethrough: false,
      fontSize: 10,
      color: null,
      fontFamily: 'Times New Roman',
      fontFamilyEastAsia: '',
      isLink: false,
      background: null,
      vertAlign: null,
      hyperlink: null,
    }],
    defaultFontSize: 10,
    defaultFontFamily: 'Times New Roman',
    widowControl: false,
  } as DocParagraph;
}

function rtlNumberedParagraph(jc: 'right' | 'center'): DocParagraph {
  const paragraph = numberedParagraph(jc);
  paragraph.bidi = true;
  paragraph.numbering!.tab = 36;
  (paragraph.runs[0] as { text: string }).text = 'سلام';
  return paragraph;
}

function documentOf(paragraph: DocParagraph): DocxDocumentModel {
  return {
    section: {
      pageWidth: 400,
      pageHeight: 400,
      marginTop: 0,
      marginRight: 0,
      marginBottom: 0,
      marginLeft: 0,
      headerDistance: 0,
      footerDistance: 0,
      titlePage: false,
      evenAndOddHeaders: false,
    } as SectionProps,
    body: [paragraph as unknown as BodyElement],
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    fontFamilyClasses: { 'Times New Roman': 'roman' },
  } as unknown as DocxDocumentModel;
}

describe('paragraph border contains justified numbering marker', () => {
  beforeEach(() => {
    vi.stubGlobal(
      'createImageBitmap',
      vi.fn(async () => ({ width: 1, height: 1, close() {} }) as unknown as ImageBitmap),
    );
  });
  afterEach(() => vi.unstubAllGlobals());

  it.each(['right', 'center'] as const)('contains an LTR %s-justified marker', async (jc) => {
    const { canvas, fills, verticalStrokes } = makeRecordingCanvas();
    await renderDocumentToCanvas(documentOf(numberedParagraph(jc)), canvas, 0, {
      dpr: 1,
      width: 400,
    });

    const marker = fills.find(({ text }) => text === '1.');
    expect(marker).toBeDefined();
    const borderSides = verticalStrokes.filter(({ color }) => color === `#${BORDER_COLOR}`);
    expect(borderSides).toHaveLength(2);
    const leftBorder = borderSides.reduce((left, stroke) => stroke.x < left.x ? stroke : left);
    const borderOuterLeft = leftBorder.x - leftBorder.width / 2;

    expect(marker!.x).toBeGreaterThanOrEqual(borderOuterLeft - 1e-9);
  });

  it.each(['right', 'center'] as const)('contains a mirrored RTL %s-justified marker', async (jc) => {
    const { canvas, fills, verticalStrokes } = makeRecordingCanvas();
    await renderDocumentToCanvas(documentOf(rtlNumberedParagraph(jc)), canvas, 0, {
      dpr: 1,
      width: 400,
    });

    const marker = fills.find(({ text }) => text === '1.');
    expect(marker).toBeDefined();
    const borderSides = verticalStrokes.filter(({ color }) => color === `#${BORDER_COLOR}`);
    expect(borderSides).toHaveLength(2);
    const rightBorder = borderSides.reduce((right, stroke) => stroke.x > right.x ? stroke : right);
    const borderOuterRight = rightBorder.x + rightBorder.width / 2;
    const markerRight = marker!.x + 20;

    expect(markerRight).toBeLessThanOrEqual(borderOuterRight + 1e-9);
  });

  it('contains the actual decoded picture-bullet draw box', async () => {
    const paragraph = numberedParagraph('right');
    paragraph.numbering = {
      ...paragraph.numbering!,
      text: '',
      picBulletImagePath: 'word/media/border-picture-bullet.gif',
      picBulletMimeType: 'image/gif',
      picBulletWidthPt: 24,
      picBulletHeightPt: 12,
    };
    const { canvas, images, verticalStrokes } = makeRecordingCanvas();
    await renderDocumentToCanvas(documentOf(paragraph), canvas, 0, {
      dpr: 1,
      width: 400,
      fetchImage: async (_path, mime) => new Blob([ONE_BY_ONE_GIF], { type: mime }),
    });

    expect(images).toHaveLength(1);
    const borderSides = verticalStrokes.filter(({ color }) => color === `#${BORDER_COLOR}`);
    expect(borderSides).toHaveLength(2);
    const leftBorder = borderSides.reduce((left, stroke) => stroke.x < left.x ? stroke : left);
    const rightBorder = borderSides.reduce((right, stroke) => stroke.x > right.x ? stroke : right);
    const borderOuterLeft = leftBorder.x - leftBorder.width / 2;
    const borderOuterRight = rightBorder.x + rightBorder.width / 2;

    expect(images[0].x).toBeGreaterThanOrEqual(borderOuterLeft - 1e-9);
    expect(images[0].x + images[0].w).toBeLessThanOrEqual(borderOuterRight + 1e-9);
  });
});
