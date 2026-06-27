import { describe, it, expect } from 'vitest';
import { layoutRichTextLines, drawResolvedRichLine } from './renderer.js';
import type { CellFont, Run } from './types.js';

// ECMA-376 §18.4.14 vertAlign / §22.9.2.17 ST_VerticalAlignRun: a run with
// vertAlign "superscript"/"subscript" renders at a reduced size (≈65%) shifted
// off the baseline. The NON-wrap rich path (PR #586) applies this; the WRAP rich
// path used to draw every `RichLine.segment` at its full `seg.font` size on the
// baseline — `layoutRichTextLines` measured segment widths at full size and both
// wrap draw loops ignored `vertAlign`. This file covers the wrap-path fix:
//   1. layout reserves the *reduced* width for a super/subscript segment (so
//      wrapping and x-advance match the drawn glyph), while line height still
//      uses the run's full size;
//   2. the shared `drawResolvedRichLine` draws a super/subscript segment with the
//      reduced font and a baseline shift (up for super, down for sub).

const BASE: CellFont = {
  bold: false,
  italic: false,
  underline: false,
  strike: false,
  size: 11,
  color: null,
  name: null,
};

function makeMeasureCtx(): CanvasRenderingContext2D {
  let font = '11px sans-serif';
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '11');
  return {
    get font() { return font; },
    set font(v: string) { font = v; },
    measureText: (s: string) => ({ width: [...s].length * px() }) as TextMetrics,
  } as unknown as CanvasRenderingContext2D;
}

function superRun(text: string, size: number): Run {
  return { text, font: { bold: false, italic: false, underline: false, strike: false, size, vertAlign: 'superscript' } };
}

describe('layoutRichTextLines reserves reduced width for super/subscript (§18.4.14)', () => {
  it('a superscript segment is measured narrower than the same text at full size', () => {
    // Wide maxWidth so nothing wraps: one token → one segment.
    const normal = layoutRichTextLines(makeMeasureCtx(), [{ text: 'X', font: { bold: false, italic: false, underline: false, strike: false, size: 20 } }] as Run[], BASE, 1, 100000);
    const sup = layoutRichTextLines(makeMeasureCtx(), [superRun('X', 20)], BASE, 1, 100000);

    const normalSeg = normal[0].segments[0];
    const supSeg = sup[0].segments[0];

    expect(supSeg.font.vertAlign).toBe('superscript');
    // The superscript glyph occupies ~65% of the full-size width.
    expect(supSeg.width).toBeLessThan(normalSeg.width);
    // Line height is unaffected — it still uses the run's full size (20pt).
    expect(sup[0].maxFontSize).toBe(20);
    expect(sup[0].maxFontSize).toBe(normal[0].maxFontSize);
  });
});

interface FillTextCall { text: string; x: number; y: number; fontPx: number; }

function makeRecordingCtx(): { ctx: CanvasRenderingContext2D; calls: FillTextCall[] } {
  let font = '11px sans-serif';
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '11');
  const calls: FillTextCall[] = [];
  const ctx = {
    get font() { return font; },
    set font(v: string) { font = v; },
    measureText: (s: string) => ({ width: [...s].length * px() }) as TextMetrics,
    fillText(text: string, x: number, y: number) { calls.push({ text, x, y, fontPx: px() }); },
    save() {}, restore() {}, beginPath() {}, moveTo() {}, lineTo() {}, stroke() {},
    fillStyle: '#000' as string,
    strokeStyle: '#000' as string,
    lineWidth: 1,
    textAlign: 'left' as CanvasTextAlign,
    textBaseline: 'alphabetic' as CanvasTextBaseline,
    direction: 'ltr' as 'ltr' | 'rtl',
  };
  return { ctx: ctx as unknown as CanvasRenderingContext2D, calls };
}

function seg(text: string, vertAlign?: 'superscript' | 'subscript'): { text: string; font: CellFont; width: number } {
  return { text, font: { ...BASE, size: 11, vertAlign }, width: text.length * 11 };
}

const TEXT_Y = 100;

function drawOne(s: { text: string; font: CellFont; width: number }): FillTextCall {
  const { ctx, calls } = makeRecordingCtx();
  // LTR, no bidi; 'top' baseline (the wrap path's per-line baseline).
  drawResolvedRichLine(ctx, [s], 0, TEXT_Y, 'top', 1, 1, {});
  return calls[0];
}

describe('drawResolvedRichLine draws super/subscript reduced + shifted (§18.4.14)', () => {
  it('a normal segment draws at full size on the baseline', () => {
    const c = drawOne(seg('A'));
    // 11pt → round(11 * 96/72) = 15px
    expect(c.fontPx).toBe(15);
    expect(c.y).toBe(TEXT_Y);
  });

  it('a superscript segment draws at reduced size, shifted UP', () => {
    const c = drawOne(seg('A', 'superscript'));
    expect(c.fontPx).toBeLessThan(15); // ~65% of 15
    expect(c.y).toBeLessThan(TEXT_Y);  // shifted up off the baseline
  });

  it('a subscript segment draws at reduced size, shifted DOWN', () => {
    const c = drawOne(seg('A', 'subscript'));
    expect(c.fontPx).toBeLessThan(15);
    expect(c.y).toBeGreaterThan(TEXT_Y); // shifted down
  });
});
