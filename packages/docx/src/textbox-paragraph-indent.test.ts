import { describe, it, expect } from 'vitest';
import { renderShapeText } from './renderer.js';
import type { ShapeRun, ShapeText } from './types';

// ECMA-376 §17.3.1.12 — a text-box paragraph honors its `<w:ind>` left/right/
// first-line indent. The first line carries the (signed) first-line indent; a
// wrapping continuation line sits at the inner-left + left-indent. Alignment
// (center/right) is computed within the INDENTED region, not the full inner box.
// When all three indents are 0 the output is byte-identical to no indent.

interface Call { text: string; x: number; y: number; }

function makeRecordingCanvas(): { ctx: CanvasRenderingContext2D; calls: Call[] } {
  let font = '10px serif';
  const calls: Call[] = [];
  const ctx = {
    get font() { return font; },
    set font(v: string) { font = v; },
    letterSpacing: '0px',
    measureText: (s: string) => {
      const p = parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
      return {
        width: [...s].length * p * 0.5,
        fontBoundingBoxAscent: p * 0.8, fontBoundingBoxDescent: p * 0.2,
        actualBoundingBoxAscent: p * 0.8, actualBoundingBoxDescent: p * 0.2,
      } as TextMetrics;
    },
    save() {}, restore() {}, beginPath() {}, closePath() {},
    moveTo() {}, lineTo() {}, stroke() {}, fill() {}, fillRect() {},
    strokeRect() {}, clip() {}, rect() {}, scale() {}, translate() {}, rotate() {},
    setLineDash() {}, clearRect() {}, arc() {}, quadraticCurveTo() {},
    bezierCurveTo() {}, createLinearGradient() { return { addColorStop() {} }; },
    drawImage() {},
    fillText(s: string, x: number, y: number) { calls.push({ text: s, x, y }); },
    strokeText(s: string, x: number, y: number) { calls.push({ text: s, x, y }); },
    fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
    textAlign: 'left' as CanvasTextAlign, direction: 'ltr' as CanvasDirection,
    globalAlpha: 1, lineCap: 'butt' as CanvasLineCap, lineJoin: 'miter' as CanvasLineJoin,
  };
  return { ctx: ctx as unknown as CanvasRenderingContext2D, calls };
}

// A rect shape at (x=100,y=50,w=200,h=200) with NO insets so innerX=100,
// innerW=200. textBlocks is the single argument under test.
function shapeWith(blocks: ShapeText[]): ShapeRun {
  return {
    type: 'shape',
    presetGeometry: 'rect', wrapMode: 'none', textAnchor: 't',
    textInsetL: 0, textInsetT: 0, textInsetR: 0, textInsetB: 0,
    textBlocks: blocks,
  } as unknown as ShapeRun;
}

const SHAPE_X = 100, SHAPE_Y = 50, SHAPE_W = 200, SHAPE_H = 200;
const INNER_X = SHAPE_X; // no insets
const SCALE = 1;
const FONTS = { 'Times New Roman': 'roman' };

function render(blocks: ShapeText[]): Call[] {
  const { ctx, calls } = makeRecordingCanvas();
  renderShapeText(shapeWith(blocks), SHAPE_X, SHAPE_Y, SHAPE_W, SHAPE_H, ctx, SCALE, FONTS);
  return calls;
}

function leftBlock(text: string, extra: Partial<ShapeText> = {}): ShapeText {
  return { text, fontSizePt: 10, fontFamily: 'Times New Roman', alignment: 'left', ...extra } as unknown as ShapeText;
}

describe('text-box paragraph indent (§17.3.1.12)', () => {
  it('shifts a left-aligned first line right by indentLeft*scale', () => {
    const baseline = render([leftBlock('Hello')]);
    const indented = render([leftBlock('Hello', { indentLeft: 18 })]);
    const b = baseline.find((c) => c.text === 'Hello');
    const i = indented.find((c) => c.text === 'Hello');
    expect(b).toBeDefined();
    expect(i).toBeDefined();
    // No indent ⇒ x = innerX. With indentLeft=18pt (scale 1) ⇒ x = innerX + 18.
    expect((b as Call).x).toBeCloseTo(INNER_X, 5);
    expect((i as Call).x).toBeCloseTo(INNER_X + 18, 5);
  });

  it('adds the (positive) first-line indent on top of the left indent for the first line', () => {
    const calls = render([leftBlock('Hello', { indentLeft: 10, indentFirst: 8 })]);
    const i = calls.find((c) => c.text === 'Hello');
    expect(i).toBeDefined();
    // First line region-left = innerX + leftPx + firstPx = innerX + 10 + 8.
    expect((i as Call).x).toBeCloseTo(INNER_X + 18, 5);
  });

  it('keeps the continuation line at innerX+leftPx while the first line carries the first-line indent', () => {
    // 0.5px/char width (10px font). innerW=200, leftPx=20 ⇒ paraW=180 (=36 chars).
    // firstPx=+40 ⇒ firstLineW=140 (=28 chars). A long word stream wraps: the
    // first line holds fewer chars (narrower firstLineW), the continuation more.
    const words = Array.from({ length: 30 }, (_, n) => `w${n}`).join(' ');
    const calls = render([leftBlock(words, { indentLeft: 20, indentFirst: 40 })]);
    // Lines are emitted top-to-bottom; group fillText by y.
    const ys = [...new Set(calls.map((c) => c.y))].sort((a, b) => a - b);
    expect(ys.length).toBeGreaterThanOrEqual(2);
    const firstLineX = calls.find((c) => c.y === ys[0])?.x;
    const contLineX = calls.find((c) => c.y === ys[1])?.x;
    expect(firstLineX).toBeDefined();
    expect(contLineX).toBeDefined();
    // First line: innerX + leftPx + firstPx = 100 + 20 + 40 = 160.
    expect(firstLineX as number).toBeCloseTo(INNER_X + 20 + 40, 5);
    // Continuation: innerX + leftPx = 100 + 20 = 120 (no first-line indent).
    expect(contLineX as number).toBeCloseTo(INNER_X + 20, 5);
  });

  it('centers within the indented region for a ctr paragraph', () => {
    // ctr block, no indent: x_center0 = innerX + (innerW - w)/2.
    const word = 'Hi';
    const w = [...word].length * 10 * 0.5; // 0.5px/char @10px = 10px wide.
    const base = render([leftBlock(word, { alignment: 'center' })]);
    const indented = render([leftBlock(word, { alignment: 'center', indentLeft: 30, indentRight: 10 })]);
    const b = base.find((c) => c.text === word);
    const i = indented.find((c) => c.text === word);
    expect(b).toBeDefined();
    expect(i).toBeDefined();
    // No indent: centered in [innerX, innerW].
    expect((b as Call).x).toBeCloseTo(INNER_X + (SHAPE_W - w) / 2, 5);
    // Indented: region-left = innerX + 30, region-width = innerW - 30 - 10 = 160.
    const regionLeft = INNER_X + 30;
    const regionW = SHAPE_W - 30 - 10;
    expect((i as Call).x).toBeCloseTo(regionLeft + (regionW - w) / 2, 5);
  });

  it('is byte-identical to no-indent when all three indents are 0', () => {
    const a = render([leftBlock('Hello world this wraps maybe', { alignment: 'left' })]);
    const b = render([leftBlock('Hello world this wraps maybe', { alignment: 'left', indentLeft: 0, indentRight: 0, indentFirst: 0 })]);
    expect(b.map((c) => [c.text, c.x, c.y])).toEqual(a.map((c) => [c.text, c.x, c.y]));
  });
});
