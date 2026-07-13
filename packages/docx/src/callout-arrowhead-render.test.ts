import { describe, expect, it } from 'vitest';
import { renderDocumentToCanvas } from './renderer.js';
import type { BodyElement, DocParagraph, DocxDocumentModel, SectionProps, ShapeRun } from './types';

function makeRecordingCanvas(): { canvas: HTMLCanvasElement; fills: number; strokes: number } {
  let fills = 0;
  let strokes = 0;
  let font = '11px serif';
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '11');
  const ctx = {
    get font() { return font; },
    set font(v: string) { font = v; },
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
    moveTo() {}, lineTo() {}, bezierCurveTo() {}, quadraticCurveTo() {},
    arc() {}, ellipse() {}, rect() {}, clip() {},
    translate() {}, rotate() {}, scale() {}, setLineDash() {},
    clearRect() {}, fillRect() {}, strokeRect() {}, drawImage() {},
    stroke() { strokes += 1; },
    fill() { fills += 1; },
    fillText() {}, strokeText() {},
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    lineCap: 'butt' as CanvasLineCap,
    lineJoin: 'miter' as CanvasLineJoin,
    textAlign: 'left' as CanvasTextAlign,
    textBaseline: 'alphabetic' as CanvasTextBaseline,
    direction: 'ltr' as CanvasDirection,
    globalAlpha: 1,
  };
  const canvas = { width: 0, height: 0, style: {} as Record<string, string>, getContext: () => ctx };
  return {
    canvas: canvas as unknown as HTMLCanvasElement,
    get fills() { return fills; },
    get strokes() { return strokes; },
  };
}

function para(runs: DocParagraph['runs']): BodyElement {
  return {
    type: 'paragraph',
    alignment: 'left',
    indentLeft: 0,
    indentRight: 0,
    indentFirst: 0,
    spaceBefore: 0,
    spaceAfter: 0,
    lineSpacing: null,
    numbering: null,
    tabStops: [],
    runs,
    defaultFontSize: 11,
    defaultFontFamily: 'Arial',
    widowControl: false,
  } as unknown as BodyElement;
}

function calloutWithTailArrow(): ShapeRun & { type: 'shape' } {
  return {
    type: 'shape',
    zOrder: 0,
    subpaths: [],
    presetGeometry: 'accentBorderCallout2',
    adjValues: [null, null, null, null, null, null, null, null],
    fill: null,
    stroke: '000000',
    strokeWidth: 1.5,
    tailEnd: { type: 'triangle', w: 'med', len: 'med' },
    widthPt: 120,
    heightPt: 80,
    anchorXPt: 40,
    anchorYPt: 40,
    anchorXFromMargin: false,
    anchorYFromPara: true,
    anchorXRelativeFrom: 'column',
    anchorYRelativeFrom: 'paragraph',
    wrapMode: 'none',
    behindDoc: false,
  } as unknown as ShapeRun & { type: 'shape' };
}

function plainLine(): ShapeRun & { type: 'shape' } {
  return {
    type: 'shape',
    zOrder: 0,
    subpaths: [],
    presetGeometry: 'line',
    fill: null,
    stroke: '000000',
    strokeWidth: 0.5,
    widthPt: 120,
    heightPt: 0.75,
    anchorXPt: 40,
    anchorYPt: 40,
    anchorXFromMargin: false,
    anchorYFromPara: true,
    anchorXRelativeFrom: 'column',
    anchorYRelativeFrom: 'paragraph',
    wrapMode: 'none',
    behindDoc: false,
  } as unknown as ShapeRun & { type: 'shape' };
}

function docWith(body: BodyElement[]): DocxDocumentModel {
  return {
    section: {
      pageWidth: 300,
      pageHeight: 220,
      marginTop: 20,
      marginRight: 20,
      marginBottom: 20,
      marginLeft: 20,
      headerDistance: 10,
      footerDistance: 10,
      titlePage: false,
      evenAndOddHeaders: false,
    } as SectionProps,
    body,
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    fontFamilyClasses: { Arial: 'swiss' },
  } as unknown as DocxDocumentModel;
}

describe('callout line-end rendering', () => {
  it('strokes an undecorated line exactly once', async () => {
    const rec = makeRecordingCanvas();

    await renderDocumentToCanvas(
      docWith([para([plainLine()])]),
      rec.canvas,
      0,
      { dpr: 1, width: 300 },
    );

    expect(rec.strokes).toBe(1);
  });

  it('draws a triangle tailEnd on the callout leader tip', async () => {
    const rec = makeRecordingCanvas();

    await renderDocumentToCanvas(
      docWith([para([calloutWithTailArrow()])]),
      rec.canvas,
      0,
      { dpr: 1, width: 300 },
    );

    expect(rec.fills).toBeGreaterThan(0);
    // accentBorderCallout2: text-box border + accent bar + retracted leader.
    expect(rec.strokes).toBe(3);
  });
});
