import { describe, it, expect } from 'vitest';
import { renderDocumentToCanvas } from './renderer.js';
import type { DocParagraph, DocxDocumentModel, SectionProps, DocRun } from './types.js';

// ECMA-376 §17.18.44 `thaiDistribute` ("Thai Language Justification") — issue
// #959. Thai has no inter-word spaces, so `both`/`distribute` reach no gap on a
// continuous-Thai line and leave it ragged. `thaiDistribute` instead distributes
// the line's slack across every GRAPHEME-CLUSTER boundary (a combining vowel/tone
// mark stays glued to its base), justifying the non-final lines to the right text
// margin. Verified against the Word-exported adjudication fixture:
//   • both / distribute over continuous Thai: interior gaps stay natural, ragged.
//   • thaiDistribute: every inter-cluster gap widens uniformly → line fills the
//     margin; the paragraph's FINAL line stays flush-left / ragged (like `both`).
// These end-to-end tests drive the docx renderer with a recording canvas (each
// code point measured at a fixed px) and read back the drawn glyph x-positions.

interface FillCall {
  text: string;
  x: number;
  y: number;
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
    fillText(text: string, x: number, y: number) { fills.push({ text, x, y }); },
    strokeText() {},
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    textAlign: 'left' as CanvasTextAlign, direction: 'ltr' as CanvasDirection,
    globalAlpha: 1, lineCap: 'butt' as CanvasLineCap, lineJoin: 'miter' as CanvasLineJoin,
  };
  const canvas = { width: 0, height: 0, style: {} as Record<string, string>, getContext: () => ctx };
  (ctx as unknown as { canvas: unknown }).canvas = canvas;
  return { canvas: canvas as unknown as HTMLCanvasElement, fills };
}

const PAGE_W = 240;
const FS = 10; // px per code point; scale = 1 px/pt, zero margins

// A continuous Thai sentence (no spaces) long enough to wrap to several lines in
// the narrow column. "การเขียนภาษาไทย…" repeated.
const THAI = 'การเขียนภาษาไทยไม่ใช้ช่องว่างระหว่างคำแต่ใช้ช่องว่างเฉพาะเมื่อจบประโยค';

function textRun(text: string): DocRun {
  return {
    type: 'text', text,
    bold: false, italic: false, underline: false, strikethrough: false,
    fontSize: FS, color: null, fontFamily: 'Leelawadee UI', fontFamilyEastAsia: 'Leelawadee UI',
    isLink: false, background: null, vertAlign: null, hyperlink: null,
  } as unknown as DocRun;
}

function thaiPara(alignment: string): DocParagraph {
  return {
    alignment,
    indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null,
    numbering: null,
    tabStops: [],
    runs: [textRun(THAI)],
    defaultFontSize: FS, defaultFontFamily: 'Leelawadee UI', widowControl: false,
  } as unknown as DocParagraph;
}

function doc(p: DocParagraph): DocxDocumentModel {
  return {
    section: {
      pageWidth: PAGE_W, pageHeight: 800,
      marginTop: 0, marginRight: 0, marginBottom: 0, marginLeft: 0,
      headerDistance: 0, footerDistance: 0, titlePage: false, evenAndOddHeaders: false,
    } as SectionProps,
    body: [{ type: 'paragraph', ...p }],
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
    fontFamilyClasses: { 'Leelawadee UI': 'sansSerif' },
  } as unknown as DocxDocumentModel;
}

async function render(alignment: string): Promise<FillCall[]> {
  const { canvas, fills } = makeRecordingCanvas();
  await renderDocumentToCanvas(doc(thaiPara(alignment)), canvas, 0, { dpr: 1, width: PAGE_W });
  return fills;
}

/** Group drawn glyph pieces by baseline y (one group per rendered line) and
 *  return each line's rightmost drawn edge px, ordered top→bottom. */
function lineRightEdges(fills: FillCall[]): number[] {
  const byY = new Map<number, number>();
  for (const f of fills) {
    const key = Math.round(f.y);
    const right = f.x + [...f.text].length * FS;
    byY.set(key, Math.max(byY.get(key) ?? -Infinity, right));
  }
  return [...byY.entries()].sort((a, b) => a[0] - b[0]).map(([, r]) => r);
}

describe('thaiDistribute justifies continuous Thai at cluster granularity (§17.18.44, #959)', () => {
  it('fills non-final lines to the right margin under thaiDistribute', async () => {
    const edges = lineRightEdges(await render('thaiDistribute'));
    expect(edges.length).toBeGreaterThan(1);
    // Every line EXCEPT the last reaches the right text margin (justified).
    for (let i = 0; i < edges.length - 1; i++) {
      expect(edges[i], `line ${i} right edge`).toBeCloseTo(PAGE_W, 1);
    }
    // The paragraph's FINAL line stays ragged (natural width < margin).
    expect(edges[edges.length - 1]).toBeLessThan(PAGE_W - FS);
  });

  it('leaves the same Thai ragged under `both` (no cluster distribution)', async () => {
    const edges = lineRightEdges(await render('both'));
    expect(edges.length).toBeGreaterThan(1);
    // No line — not even a non-final one — is stretched to the margin: Thai has no
    // inter-word space and `both` never splits clusters, so each line ends at its
    // natural width, short of the margin.
    for (let i = 0; i < edges.length; i++) {
      expect(edges[i], `both line ${i} right edge`).toBeLessThan(PAGE_W);
    }
  });

  it('leaves the same Thai ragged under `distribute` too (unit is the token, not the cluster)', async () => {
    const edges = lineRightEdges(await render('distribute'));
    expect(edges.length).toBeGreaterThan(1);
    for (let i = 0; i < edges.length; i++) {
      expect(edges[i], `distribute line ${i} right edge`).toBeLessThan(PAGE_W);
    }
  });
});
