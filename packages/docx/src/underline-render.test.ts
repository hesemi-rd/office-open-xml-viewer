import { describe, it, expect } from 'vitest';
import { renderDocumentToCanvas } from './renderer.js';
import type { BodyElement, DocParagraph, DocxDocumentModel, SectionProps } from './types';

// Characterization of the docx run-underline draw path (§17.3.2.40 `<w:u>`),
// which delegates styled underlines to core.drawUnderline (§20.1.10.82) while
// keeping the plain single rule byte-stable. We record the ctx stroke ops so we
// can assert the geometry / dash / colour without a real canvas.

interface StrokeOp { op: string; args: (number | string)[]; }

function makeRecordingCanvas(): { canvas: HTMLCanvasElement; ops: StrokeOp[]; dashes: number[][] } {
  let font = '10px serif';
  const ops: StrokeOp[] = [];
  const dashes: number[][] = [];
  let strokeStyle = '#000';
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
  const ctx = {
    get font() { return font; },
    set font(v: string) { font = v; },
    get strokeStyle() { return strokeStyle; },
    set strokeStyle(v: string) { strokeStyle = v; },
    letterSpacing: '0px',
    measureText: (s: string) => {
      const p = px();
      const lowLine = s === '_';
      return {
        width: [...s].length * p * 0.5,
        fontBoundingBoxAscent: p * 0.8, fontBoundingBoxDescent: p * 0.2,
        actualBoundingBoxAscent: lowLine ? 0 : p * 0.8,
        actualBoundingBoxDescent: lowLine ? p * 0.05 : p * 0.2,
      } as TextMetrics;
    },
    save() {}, restore() {}, beginPath() {}, closePath() {},
    moveTo(x: number, y: number) { ops.push({ op: 'moveTo', args: [x, y] }); },
    lineTo(x: number, y: number) { ops.push({ op: 'lineTo', args: [x, y] }); },
    stroke() { ops.push({ op: 'stroke', args: [strokeStyle] }); },
    fill() {}, fillRect() {},
    strokeRect() {}, clip() {}, rect() {}, scale() {}, translate() {}, rotate() {},
    setLineDash(d: number[]) { dashes.push(d.slice()); },
    clearRect() {}, arc() {}, quadraticCurveTo() {},
    bezierCurveTo() {}, createLinearGradient() { return { addColorStop() {} }; },
    drawImage() {},
    fillText() {},
    strokeText() {},
    fillStyle: '#000', lineWidth: 1,
    textAlign: 'left' as CanvasTextAlign, direction: 'ltr' as CanvasDirection,
    globalAlpha: 1, lineCap: 'butt' as CanvasLineCap, lineJoin: 'miter' as CanvasLineJoin,
  };
  const canvas = { width: 0, height: 0, style: {} as Record<string, string>, getContext: () => ctx };
  return { canvas: canvas as unknown as HTMLCanvasElement, ops, dashes };
}

function para(run: Partial<DocParagraph['runs'][number]>): DocParagraph {
  return {
    type: 'paragraph', alignment: 'left',
    indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null,
    numbering: null, tabStops: [],
    runs: [{
      type: 'text', text: 'abc', bold: false, italic: false, underline: true,
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

async function render(run: Partial<DocParagraph['runs'][number]>) {
  const { canvas, ops, dashes } = makeRecordingCanvas();
  await renderDocumentToCanvas(doc([para(run) as unknown as BodyElement]), canvas, 0, { dpr: 1, width: 400 });
  return { ops, dashes };
}

// Underline strokes are horizontal (moveTo.y == the following lineTo.y).
function horizontalStrokes(ops: StrokeOp[]): { y: number; x1: number; x2: number }[] {
  const out: { y: number; x1: number; x2: number }[] = [];
  for (let i = 0; i < ops.length - 1; i++) {
    if (ops[i].op === 'moveTo' && ops[i + 1].op === 'lineTo') {
      const [x1, y1] = ops[i].args as number[];
      const [x2, y2] = ops[i + 1].args as number[];
      if (Math.abs(y1 - y2) < 0.001) out.push({ y: y1, x1, x2 });
    }
  }
  return out;
}

describe('docx run underline (§17.3.2.40) draw path', () => {
  it('plain single underline draws exactly one straight horizontal rule (byte-stable)', async () => {
    const styled = await render({ underline: true });
    const strokes = horizontalStrokes(styled.ops);
    // The single rule is one horizontal segment under the run. (Strike-through is
    // absent here, so any horizontal stroke is the underline.)
    expect(strokes.length).toBe(1);
    // No non-empty dash pattern is applied for the plain single rule.
    expect(styled.dashes.every((d) => d.length === 0)).toBe(true);
  });

  it('dotted style routes through core with a non-empty dash pattern', async () => {
    const { dashes } = await render({ underline: true, underlineStyle: 'dotted' });
    // core.drawUnderline sets a non-empty dash for dotted; the single path never does.
    expect(dashes.some((d) => d.length > 0)).toBe(true);
  });

  it('double style (dbl) draws two parallel horizontal rules', async () => {
    const { ops } = await render({ underline: true, underlineStyle: 'double' });
    const strokes = horizontalStrokes(ops);
    expect(strokes.length).toBe(2);
    // Two distinct y rows straddling the anchor.
    expect(strokes[0].y).not.toBeCloseTo(strokes[1].y, 3);
  });

  it('wave style traces a multi-point polyline (many non-horizontal segments)', async () => {
    const { ops } = await render({ underline: true, underlineStyle: 'wave' });
    const lineTos = ops.filter((o) => o.op === 'lineTo');
    // A sine polyline emits far more lineTo ops than a straight rule (1).
    expect(lineTos.length).toBeGreaterThan(5);
  });

  it('w:u@color overrides the underline stroke colour', async () => {
    const { ops } = await render({ underline: true, underlineStyle: 'dotted', underlineColor: 'ff0000', color: '0000ff' });
    // The styled underline strokes with the underline colour, not the glyph colour.
    const strokeColors = ops
      .filter((o) => o.op === 'stroke')
      .map((o) => String(o.args[0]).toLowerCase());
    expect(strokeColors).toContain('#ff0000');
  });

  it('w:u@color=auto follows the glyph colour (no override)', async () => {
    const { ops } = await render({ underline: true, underlineStyle: 'dotted', underlineColor: 'auto', color: '0000ff' });
    const strokeColors = ops
      .filter((o) => o.op === 'stroke')
      .map((o) => String(o.args[0]).toLowerCase());
    // auto → glyph colour (#0000ff), never the literal 'auto'.
    expect(strokeColors).toContain('#0000ff');
    expect(strokeColors).not.toContain('#auto');
  });
});
