import { describe, it, expect } from 'vitest';
import { renderDocumentToCanvas } from './renderer.js';
import type { DocParagraph, DocxDocumentModel, SectionProps, DocRun } from './types.js';

// ECMA-376 §17.18.44 — TRUE kashida justification. For jc=lowKashida /
// mediumKashida / highKashida the line slack of an Arabic justified line is
// filled by inserting U+0640 tatweel at valid Arabic joining points (elongating
// the words), NOT by widening inter-word spaces (the `both` behaviour). These
// end-to-end tests drive the renderer with the mock canvas and assert that the
// painted glyph strings gain tatweels under kashida jc and DO NOT under `both`.

const TATWEEL = 'ـ'; // U+0640
const BEH = 'ب'; // U+0628 dual-joining — a 4-beh word has 3 interior kashida points

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
    fontKerning: 'auto' as CanvasFontKerning,
  };
  const canvas = { width: 0, height: 0, style: {} as Record<string, string>, getContext: () => ctx };
  (ctx as unknown as { canvas: unknown }).canvas = canvas;
  return { canvas: canvas as unknown as HTMLCanvasElement, fills };
}

const PAGE_W = 300;
const FS = 10;

function textRun(text: string): DocRun {
  return {
    type: 'text', text,
    bold: false, italic: false, underline: false, strikethrough: false,
    fontSize: FS, color: null, fontFamily: 'Times New Roman', fontFamilyEastAsia: 'Times New Roman',
    isLink: false, background: null, vertAlign: null, hyperlink: null,
  } as unknown as DocRun;
}

// Two 4-beh Arabic words on line 1; a long filler forces the wrap so line 1 is a
// NON-LAST (therefore justified) line. Each 4-beh word has 3 interior kashida
// insertion points.
const W = BEH.repeat(4);
const LONG = 'w'.repeat(28); // 280px — cannot share line 1.
function jcPara(alignment: string): DocParagraph {
  return {
    alignment,
    indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null,
    numbering: null,
    tabStops: [],
    runs: [textRun(`${W} ${W} ${LONG}`)],
    defaultFontSize: FS, defaultFontFamily: 'Times New Roman', widowControl: false,
  } as unknown as DocParagraph;
}

function doc(p: DocParagraph): DocxDocumentModel {
  return {
    section: {
      pageWidth: PAGE_W, pageHeight: 400,
      marginTop: 0, marginRight: 0, marginBottom: 0, marginLeft: 0,
      headerDistance: 0, footerDistance: 0, titlePage: false, evenAndOddHeaders: false,
    } as SectionProps,
    body: [{ type: 'paragraph', ...p }],
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    fontFamilyClasses: { 'Times New Roman': 'roman' },
  } as unknown as DocxDocumentModel;
}

async function render(alignment: string) {
  const { canvas, fills } = makeRecordingCanvas();
  await renderDocumentToCanvas(doc(jcPara(alignment)), canvas, 0, { dpr: 1, width: PAGE_W });
  return fills;
}

const tatweelsDrawn = (f: FillCall[]) =>
  f.reduce((n, c) => n + [...c.text].filter((ch) => ch === TATWEEL).length, 0);

describe('jc kashida inserts U+0640 tatweel (§17.18.44)', () => {
  for (const jc of ['lowKashida', 'mediumKashida', 'highKashida']) {
    it(`${jc} elongates the Arabic word with tatweels`, async () => {
      const f = await render(jc);
      expect(tatweelsDrawn(f)).toBeGreaterThan(0);
    });
  }

  it('highKashida inserts at least as many tatweels as lowKashida (more aggressive)', async () => {
    const low = tatweelsDrawn(await render('lowKashida'));
    const high = tatweelsDrawn(await render('highKashida'));
    expect(high).toBeGreaterThanOrEqual(low);
    expect(low).toBeGreaterThan(0);
  });

  it('ordinary `both` justification inserts NO tatweel (widens spaces instead)', async () => {
    expect(tatweelsDrawn(await render('both'))).toBe(0);
  });

  it('left alignment (no justification) inserts NO tatweel', async () => {
    expect(tatweelsDrawn(await render('left'))).toBe(0);
  });
});
