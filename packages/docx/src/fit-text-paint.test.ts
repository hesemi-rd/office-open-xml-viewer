import { describe, expect, it } from 'vitest';
import { renderDocumentToCanvas, type DocxTextRunInfo } from './renderer.js';
import type {
  BodyElement,
  DocParagraph,
  DocxDocumentModel,
  DocxTextRun,
  SectionProps,
} from './types.js';

interface FillCall {
  text: string;
  letterSpacing: number;
  scaleX: number;
}

function makeRecordingCanvas(): { canvas: HTMLCanvasElement; fills: FillCall[] } {
  let font = '12px serif';
  let letterSpacing = '0px';
  let scaleX = 1;
  const stack: number[] = [];
  const fills: FillCall[] = [];
  const ctx = {
    get font() { return font; },
    set font(value: string) { font = value; },
    get letterSpacing() { return letterSpacing; },
    set letterSpacing(value: string) { letterSpacing = value; },
    fontKerning: 'auto',
    measureText(text: string) {
      const px = Number(/([\d.]+)px/.exec(font)?.[1] ?? 12);
      return {
        width: [...text].length * px,
        fontBoundingBoxAscent: px * 0.8,
        fontBoundingBoxDescent: px * 0.2,
        actualBoundingBoxAscent: px * 0.8,
        actualBoundingBoxDescent: px * 0.2,
      } as TextMetrics;
    },
    save() { stack.push(scaleX); },
    restore() { scaleX = stack.pop() ?? 1; },
    scale(x: number) { scaleX *= x; },
    translate() {},
    fillText(text: string) {
      fills.push({ text, letterSpacing: Number.parseFloat(letterSpacing), scaleX });
    },
    beginPath() {}, closePath() {}, moveTo() {}, lineTo() {}, stroke() {}, fill() {},
    fillRect() {}, strokeRect() {}, clip() {}, rect() {}, setLineDash() {},
    drawImage() {}, clearRect() {}, arc() {}, quadraticCurveTo() {}, bezierCurveTo() {},
    createLinearGradient() { return { addColorStop() {} }; },
    strokeText() {},
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    textAlign: 'left' as CanvasTextAlign, direction: 'ltr' as CanvasDirection,
    globalAlpha: 1, lineCap: 'butt' as CanvasLineCap, lineJoin: 'miter' as CanvasLineJoin,
  };
  const canvas = { width: 0, height: 0, style: {}, getContext: () => ctx };
  return { canvas: canvas as unknown as HTMLCanvasElement, fills };
}

function textRun(text: string, extra: Partial<DocxTextRun> = {}): DocxTextRun {
  return {
    text,
    bold: false,
    italic: false,
    underline: false,
    strikethrough: false,
    fontSize: 12,
    color: null,
    fontFamily: 'serif',
    isLink: false,
    background: null,
    vertAlign: null,
    ...extra,
    hyperlink: extra.hyperlink ?? null,
  };
}

function paragraph(runs: DocxTextRun[]): BodyElement {
  const para: DocParagraph = {
    alignment: 'left',
    indentLeft: 0,
    indentRight: 0,
    indentFirst: 0,
    spaceBefore: 0,
    spaceAfter: 0,
    lineSpacing: null,
    numbering: null,
    tabStops: [],
    runs: runs.map((run) => ({ type: 'text', ...run })),
    defaultFontSize: 12,
    defaultFontFamily: 'serif',
    widowControl: false,
  };
  return { type: 'paragraph', ...para };
}

function section(): SectionProps {
  return {
    pageWidth: 600,
    pageHeight: 400,
    marginTop: 0,
    marginRight: 0,
    marginBottom: 0,
    marginLeft: 0,
    headerDistance: 0,
    footerDistance: 0,
    titlePage: false,
    evenAndOddHeaders: false,
  } as SectionProps;
}

async function render(runs: DocxTextRun[]): Promise<{ fills: FillCall[]; boxes: DocxTextRunInfo[] }> {
  const { canvas, fills } = makeRecordingCanvas();
  const boxes: DocxTextRunInfo[] = [];
  const model = {
    section: section(),
    body: [paragraph(runs)],
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
  } as unknown as DocxDocumentModel;
  await renderDocumentToCanvas(model, canvas, 0, {
    dpr: 1,
    width: 600,
    onTextRun: (box) => boxes.push(box),
  });
  return { fills, boxes };
}

describe('ECMA-376 §17.3.2.14 fitText paint', () => {
  it('draws the cross-run resolved gap and reports the fixed-width boxes', async () => {
    const fit = { fitTextVal: 2400, fitTextId: -1431456512, charSpacing: 4.8 };
    const { fills, boxes } = await render([
      textRun('氏名又は', fit),
      textRun('名', fit),
      textRun('称', fit),
    ]);

    expect(fills.map((fill) => fill.letterSpacing)).toEqual([9.6, 9.6, 9.6]);
    expect(boxes.reduce((sum, box) => sum + box.w, 0)).toBeCloseTo(120, 9);
    expect(boxes.map((box) => box.x)).toEqual([0, 86.4, 108]);
  });

  it('composes the gap with w:w in the scaled draw frame', async () => {
    const fit = { fitTextVal: 2400, fitTextId: -1431456510, charScale: 0.66, charSpacing: 0.9 };
    const { fills, boxes } = await render([
      textRun('並びに法人の場合は', fit),
      textRun('代表者の氏名', fit),
    ]);
    const perGap = (120 - 15 * 12 * 0.66) / 14;

    expect(fills).toHaveLength(2);
    expect(fills.every((fill) => Math.abs(fill.scaleX - 0.66) < 1e-9)).toBe(true);
    expect(fills[0].letterSpacing).toBeCloseTo(perGap / 0.66, 9);
    expect(boxes.reduce((sum, box) => sum + box.w, 0)).toBeCloseTo(120, 9);
  });
});
