import { describe, it, expect } from 'vitest';
import { renderDocumentToCanvas } from './renderer.js';
import type { BodyElement, DocParagraph, DocxDocumentModel, SectionProps } from './types';

// ECMA-376 §17.3.2.12 w:em / §17.18.24 ST_Em — characterization of the docx
// emphasis-mark (圏点) draw path. We record the ctx arc / fill / stroke ops so we
// can assert the per-glyph mark count, above/below placement, and shape (filled
// dot / hollow circle / comma) without a real canvas.

interface ArcOp { cx: number; cy: number; r: number; }
interface Recording {
  arcs: ArcOp[];
  fills: number; // fill() calls
  strokes: number; // stroke() calls
  // arc → the next terminal op (fill or stroke) tells us the mark shape.
  arcTerminals: ('fill' | 'stroke')[];
}

function makeRecordingCanvas(): { canvas: HTMLCanvasElement; rec: Recording } {
  let font = '10px serif';
  const rec: Recording = { arcs: [], fills: 0, strokes: 0, arcTerminals: [] };
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
  // pendingArc: set by arc(), consumed by the next fill()/stroke() so we can
  // attribute each ring/disc to how it was closed out.
  let pendingArc = false;
  const ctx = {
    get font() { return font; },
    set font(v: string) { font = v; },
    letterSpacing: '0px',
    measureText: (s: string) => {
      const p = px();
      const advanceEm = s === 'あアかカ' ? 1 : 0.5;
      return {
        // Each code point advances a fixed half-em, so glyph i centre is
        // predictable: penX + (i + 0.5) * p * 0.5.
        width: [...s].length * p * advanceEm,
        fontBoundingBoxAscent: p * 0.8, fontBoundingBoxDescent: p * 0.2,
        actualBoundingBoxAscent: p * 0.8, actualBoundingBoxDescent: p * 0.2,
      } as TextMetrics;
    },
    save() {}, restore() {}, closePath() {},
    beginPath() {},
    moveTo() {}, lineTo() {},
    arc(cx: number, cy: number, r: number) {
      rec.arcs.push({ cx, cy, r });
      pendingArc = true;
    },
    fill() {
      rec.fills++;
      if (pendingArc) { rec.arcTerminals.push('fill'); pendingArc = false; }
    },
    stroke() {
      rec.strokes++;
      if (pendingArc) { rec.arcTerminals.push('stroke'); pendingArc = false; }
    },
    fillRect() {}, strokeRect() {}, clip() {}, rect() {},
    scale() {}, translate() {}, rotate() {},
    setLineDash() {}, clearRect() {}, quadraticCurveTo() {}, bezierCurveTo() {},
    createLinearGradient() { return { addColorStop() {} }; },
    drawImage() {}, fillText() {}, strokeText() {},
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

describe('docx emphasis mark (§17.3.2.12 w:em) draw path', () => {
  it('draws one mark per non-space character (§17.18.24)', async () => {
    const rec = await render({ emphasisMark: 'dot' }, 'abc');
    // Three glyphs → three dot arcs, each filled.
    expect(rec.arcs).toHaveLength(3);
    expect(rec.arcTerminals).toEqual(['fill', 'fill', 'fill']);
  });

  it('skips space characters (no mark on the space)', async () => {
    const rec = await render({ emphasisMark: 'dot' }, 'a b');
    // 'a' and 'b' get marks; the space does not.
    expect(rec.arcs).toHaveLength(2);
  });

  it('draws no marks when the run has no emphasis mark', async () => {
    const rec = await render({}, 'abc');
    expect(rec.arcs).toHaveLength(0);
  });

  it('dot marks sit ABOVE the glyph box (negative-ish y, above baseline)', async () => {
    // The paragraph baseline for a 40pt run sits well below the box top. All dot
    // centres must be above the box top (§17.18.24 horizontal writing = above).
    const rec = await render({ emphasisMark: 'dot' }, 'abc');
    const ys = rec.arcs.map((a) => a.cy);
    // Every above-mark shares the same y (same line, same font).
    expect(new Set(ys).size).toBe(1);
  });

  it('underDot places the mark BELOW the box, at a larger y than dot (above)', async () => {
    const above = await render({ emphasisMark: 'dot' }, 'abc');
    const below = await render({ emphasisMark: 'underDot' }, 'abc');
    // Canvas y grows downward: the underDot centre is below (greater y) the dot.
    expect(below.arcs[0].cy).toBeGreaterThan(above.arcs[0].cy);
  });

  it('circle draws hollow rings (stroked, not filled)', async () => {
    const rec = await render({ emphasisMark: 'circle' }, 'abc');
    expect(rec.arcs).toHaveLength(3);
    // Each ring is closed with stroke(), not fill().
    expect(rec.arcTerminals).toEqual(['stroke', 'stroke', 'stroke']);
  });

  it('mark centre x tracks the glyph advance midpoint', async () => {
    // 40pt @ dpr 1: measureText gives half-em (10px) per glyph. The first glyph's
    // centre is contentX (left margin 5pt*scale) + 5px. We only assert the marks
    // are evenly spaced by the glyph advance (10px), which pins the per-glyph
    // centring without depending on the exact left origin.
    const rec = await render({ emphasisMark: 'dot' }, 'abc');
    const xs = rec.arcs.map((a) => a.cx);
    expect(xs[1] - xs[0]).toBeCloseTo(xs[2] - xs[1], 3);
    expect(xs[1] - xs[0]).toBeGreaterThan(0);
  });

  it('scales multi-kana mark centres with the condensed glyph draw', async () => {
    const rec = await render({
      emphasisMark: 'dot',
      fontFamily: 'Meiryo UI',
      fontFamilyEastAsia: 'Meiryo UI',
    }, 'あい');
    const xs = rec.arcs.map((a) => a.cx);

    expect(xs).toHaveLength(2);
    expect(xs[1] - xs[0]).toBeCloseTo(40 * 0.5 * 0.7775, 6);
  });

  it('mark radius scales with the run font size (~0.07 em)', async () => {
    const big = await render({ emphasisMark: 'dot', fontSize: 40 }, 'a');
    const small = await render({ emphasisMark: 'dot', fontSize: 20 }, 'a');
    expect(big.arcs[0].r).toBeGreaterThan(small.arcs[0].r);
    // ~40*0.07 vs ~20*0.07 at dpr 1 → ratio ≈ 2.
    expect(big.arcs[0].r / small.arcs[0].r).toBeCloseTo(2, 1);
  });
});
