import { describe, it, expect } from 'vitest';
import { acquireAndPaintShapeTextBox } from './retained-shape-textbox.test-support.js';
import type { ShapeRun, ShapeText, ShapeTextRun } from './types';

// ECMA-376 §20.1.4.1.17 `<wps:style><a:fontRef>` gives a text box a DEFAULT text
// color. A `<wps:txbx>` run that sets no explicit `<w:color>` inherits it before
// falling back to the document/theme default (black). sample-28's Arabic cover
// banner draws its color-less runs in the fontRef's `lt1` = white; without this
// fallback they render black on the dark panel (issue #821). These characterization
// tests capture the `fillStyle` active at each glyph draw to pin the precedence:
//   run color > shape fontRef default (defaultTextColor) > document default.

interface FillTextEvent { text: string; x: number; y: number; fillStyle: string }

/** Recording 2D context that captures the `fillStyle` at each fillText call.
 *  measureText width is code-point count × the current font's px so widths are
 *  deterministic (no real font needed). */
function makeRecordingCanvas(): { ctx: CanvasRenderingContext2D; fillTexts: FillTextEvent[] } {
  let font = '10px serif';
  let letterSpacing = '0px';
  let fillStyle = '#000000';
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
  const ls = () => parseFloat(/(-?\d+(?:\.\d+)?)px/.exec(letterSpacing)?.[1] ?? '0');
  const fillTexts: FillTextEvent[] = [];
  const ctx = {
    get font() { return font; },
    set font(v: string) { font = v; },
    get letterSpacing() { return letterSpacing; },
    set letterSpacing(v: string) { letterSpacing = v; },
    get fillStyle() { return fillStyle; },
    set fillStyle(v: string) { fillStyle = v; },
    measureText: (s: string) => {
      const p = px();
      const n = [...s].length;
      return {
        width: n * p + n * ls(),
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
    fillText(s: string, x: number, y: number) { fillTexts.push({ text: s, x, y, fillStyle }); },
    strokeText(s: string, x: number, y: number) { fillTexts.push({ text: s, x, y, fillStyle }); },
    strokeStyle: '#000', lineWidth: 1,
    textAlign: 'left' as CanvasTextAlign, direction: 'ltr' as CanvasDirection,
    globalAlpha: 1, lineCap: 'butt' as CanvasLineCap, lineJoin: 'miter' as CanvasLineJoin,
  };
  return { ctx: ctx as unknown as CanvasRenderingContext2D, fillTexts };
}

const SHAPE_X = 100, SHAPE_Y = 50, W = 400, H = 200, SCALE = 1;

function block(text: string, runColor?: string): ShapeText {
  return {
    text,
    fontSizePt: 10,
    fontFamily: 'serif',
    alignment: 'left',
    runs: [{ text, fontSizePt: 10, fontFamily: 'serif', color: runColor } as ShapeTextRun],
  } as unknown as ShapeText;
}

function shape(blocks: ShapeText[], defaultTextColor?: string): ShapeRun {
  return {
    type: 'shape',
    presetGeometry: 'rect', wrapMode: 'none', textAnchor: 't',
    textInsetL: 0, textInsetT: 0, textInsetR: 0, textInsetB: 0,
    textBlocks: blocks,
    defaultTextColor,
  } as unknown as ShapeRun;
}

function render(s: ShapeRun): FillTextEvent[] {
  const { ctx, fillTexts } = makeRecordingCanvas();
  acquireAndPaintShapeTextBox(s, SHAPE_X, SHAPE_Y, W, H, ctx, SCALE);
  return fillTexts;
}

describe('text-box fontRef default text color (§20.1.4.1.17, issue #821)', () => {
  it('draws a color-less run in the shape defaultTextColor (sample-28 white banner)', () => {
    // Arabic text, no run color, shape defaultTextColor = FFFFFF (fontRef lt1).
    const evs = render(shape([block('الملاحق')], 'FFFFFF'));
    expect(evs.length).toBeGreaterThan(0);
    for (const e of evs) {
      expect(e.fillStyle.toLowerCase()).toBe('#ffffff');
    }
  });

  it('a run with its own color overrides the shape defaultTextColor', () => {
    // Run carries an explicit red; the shape default (white) must NOT win.
    const evs = render(shape([block('X', 'FF0000')], 'FFFFFF'));
    expect(evs.length).toBeGreaterThan(0);
    for (const e of evs) {
      expect(e.fillStyle.toLowerCase()).toBe('#ff0000');
    }
  });

  it('falls back to black when neither the run nor the shape set a color', () => {
    // No defaultTextColor, no run color, no threaded document default ⇒ black.
    const evs = render(shape([block('X')]));
    expect(evs.length).toBeGreaterThan(0);
    for (const e of evs) {
      expect(e.fillStyle.toLowerCase()).toBe('#000000');
    }
  });
});
