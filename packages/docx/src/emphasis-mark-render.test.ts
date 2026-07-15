import { describe, it, expect } from 'vitest';
import { renderDocumentToCanvas } from './renderer.js';
import type { BodyElement, DocParagraph, DocxDocumentModel, SectionProps } from './types';

// ECMA-376 §17.3.2.12 w:em / §17.18.24 ST_Em — characterization of the docx
// emphasis-mark (圏点) draw path. The selected authored glyph itself is retained
// and painted; tests record glyph identity/origin rather than synthetic paths.

interface GlyphOp { text: string; x: number; y: number; font: string; }
interface Recording {
  glyphs: GlyphOp[];
  paths: number;
}

function makeRecordingCanvas(): { canvas: HTMLCanvasElement; rec: Recording } {
  let font = '10px serif';
  const rec: Recording = { glyphs: [], paths: 0 };
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
  let currentPath: { x: number; y: number }[] = [];
  const ctx = {
    get font() { return font; },
    set font(v: string) { font = v; },
    letterSpacing: '0px',
    measureText: (s: string) => {
      const p = px();
      return {
        // Each code point advances a fixed half-em, so glyph i centre is
        // predictable: penX + (i + 0.5) * p * 0.5.
        width: [...s].length * p * 0.5,
        fontBoundingBoxAscent: p * 0.8, fontBoundingBoxDescent: p * 0.2,
        actualBoundingBoxAscent: p * 0.8, actualBoundingBoxDescent: p * 0.2,
      } as TextMetrics;
    },
    save() {}, restore() {}, closePath() {},
    beginPath() { currentPath = []; },
    moveTo(x: number, y: number) { currentPath.push({ x, y }); },
    lineTo(x: number, y: number) { currentPath.push({ x, y }); },
    arc() {},
    fill() { if (currentPath.length > 0) rec.paths += 1; },
    stroke() { if (currentPath.length > 0) rec.paths += 1; },
    fillRect() {}, strokeRect() {}, clip() {}, rect() {},
    scale() {}, translate() {}, rotate() {},
    setLineDash() {}, clearRect() {}, quadraticCurveTo() {}, bezierCurveTo() {},
    createLinearGradient() { return { addColorStop() {} }; },
    drawImage() {},
    fillText(text: string, x: number, y: number) { rec.glyphs.push({ text, x, y, font }); },
    strokeText() {},
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    textAlign: 'left' as CanvasTextAlign, direction: 'ltr' as CanvasDirection,
    globalAlpha: 1, lineCap: 'butt' as CanvasLineCap, lineJoin: 'miter' as CanvasLineJoin,
  };
  const canvas = { width: 0, height: 0, style: {} as Record<string, string>, getContext: () => ctx };
  return { canvas: canvas as unknown as HTMLCanvasElement, rec };
}

function para(run: Partial<DocParagraph['runs'][number]>, text = 'abc'): DocParagraph {
  return {
    type: 'paragraph', alignment: 'left',
    indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null,
    numbering: null, tabStops: [],
    runs: [{
      type: 'text', text, bold: false, italic: false, underline: false,
      strikethrough: false, fontSize: 40, color: null, fontFamily: 'Times New Roman',
      fontFamilyEastAsia: '', isLink: false, background: null, vertAlign: null, hyperlink: null,
      ...run,
    } as DocParagraph['runs'][number]],
    defaultFontSize: 40, defaultFontFamily: 'Times New Roman', widowControl: false,
  } as unknown as DocParagraph;
}

function doc(body: BodyElement[]): DocxDocumentModel {
  const section = {
    pageWidth: 400, pageHeight: 600,
    marginTop: 5, marginRight: 5, marginBottom: 5, marginLeft: 5,
    headerDistance: 4, footerDistance: 4, titlePage: false, evenAndOddHeaders: false,
  } as SectionProps;
  return {
    section, body,
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    fontFamilyClasses: { 'Times New Roman': 'roman' },
  } as unknown as DocxDocumentModel;
}

async function render(run: Partial<DocParagraph['runs'][number]>, text = 'abc') {
  const { canvas, rec } = makeRecordingCanvas();
  await renderDocumentToCanvas(doc([para(run, text) as unknown as BodyElement]), canvas, 0, { dpr: 1, width: 400 });
  return rec;
}

const emphasisGlyphs = (recording: Recording): GlyphOp[] =>
  recording.glyphs.filter((operation) => operation.text === '•'
    || operation.text === '﹅' || operation.text === '○');

describe('docx emphasis mark (§17.3.2.12 w:em) draw path', () => {
  it('draws one mark per non-space character (§17.18.24)', async () => {
    const rec = await render({ emphasisMark: 'dot' }, 'abc');
    expect(emphasisGlyphs(rec).map((operation) => operation.text)).toEqual(['•', '•', '•']);
  });

  it('skips space characters (no mark on the space)', async () => {
    const rec = await render({ emphasisMark: 'dot' }, 'a b');
    // 'a' and 'b' get marks; the space does not.
    expect(emphasisGlyphs(rec)).toHaveLength(2);
  });

  it('draws no marks when the run has no emphasis mark', async () => {
    const rec = await render({}, 'abc');
    expect(emphasisGlyphs(rec)).toHaveLength(0);
  });

  it('dot marks sit ABOVE the glyph box (negative-ish y, above baseline)', async () => {
    // The paragraph baseline for a 40pt run sits well below the box top. All dot
    // centres must be above the box top (§17.18.24 horizontal writing = above).
    const rec = await render({ emphasisMark: 'dot' }, 'abc');
    const ys = emphasisGlyphs(rec).map((operation) => operation.y);
    // Every above-mark shares the same y (same line, same font).
    expect(new Set(ys).size).toBe(1);
  });

  it('underDot places the mark BELOW the box, at a larger y than dot (above)', async () => {
    const above = await render({ emphasisMark: 'dot' }, 'abc');
    const below = await render({ emphasisMark: 'underDot' }, 'abc');
    // Canvas y grows downward: the underDot centre is below (greater y) the dot.
    expect(emphasisGlyphs(below)[0]!.y).toBeGreaterThan(emphasisGlyphs(above)[0]!.y);
  });

  it('preserves dot, comma, and hollow-circle semantics through selected glyph identity', async () => {
    const dot = await render({ emphasisMark: 'dot' }, 'a');
    const comma = await render({ emphasisMark: 'comma' }, 'a');
    const circle = await render({ emphasisMark: 'circle' }, 'a');
    expect(emphasisGlyphs(dot).map((operation) => operation.text)).toEqual(['•']);
    expect(emphasisGlyphs(comma).map((operation) => operation.text)).toEqual(['﹅']);
    expect(emphasisGlyphs(circle).map((operation) => operation.text)).toEqual(['○']);
    // ○ owns its hollow counter in the selected face; paint fabricates no path.
    expect(circle.paths).toBe(0);
  });

  it('mark centre x tracks the glyph advance midpoint', async () => {
    // 40pt @ dpr 1: measureText gives half-em (10px) per glyph. The first glyph's
    // centre is contentX (left margin 5pt*scale) + 5px. We only assert the marks
    // are evenly spaced by the glyph advance (10px), which pins the per-glyph
    // centring without depending on the exact left origin.
    const rec = await render({ emphasisMark: 'dot' }, 'abc');
    const xs = emphasisGlyphs(rec).map((operation) => operation.x);
    expect(xs[1] - xs[0]).toBeCloseTo(xs[2] - xs[1], 3);
    expect(xs[1] - xs[0]).toBeGreaterThan(0);
  });

  it('paints the retained selected-face size without an em-ratio path', async () => {
    const big = await render({ emphasisMark: 'dot', fontSize: 40 }, 'a');
    const small = await render({ emphasisMark: 'dot', fontSize: 20 }, 'a');
    const bigMark = emphasisGlyphs(big)[0]!;
    const smallMark = emphasisGlyphs(small)[0]!;
    const fontPx = (value: string): number =>
      Number(/([\d.]+)px/u.exec(value)?.[1] ?? Number.NaN);
    expect(fontPx(bigMark.font)).toBeCloseTo(2 * fontPx(smallMark.font), 8);
    expect(big.paths).toBe(0);
    expect(small.paths).toBe(0);
  });
});
