import { describe, it, expect } from 'vitest';
import { drawNonWrapRichText } from './renderer.js';
import type { CellFont, Run } from './types.js';

// Companion to nonwrap-rich-text-hard-break.test.ts. `drawNonWrapRichText` is the
// single entry point for NON-wrapped rich-text cells; both the in-viewport draw
// path and the off-screen-anchor merge pre-pass route through it so they render
// identically (the off-screen pre-pass previously drew a break-free rich cell as
// joined base-font text via one `ctx.fillText`, losing per-run fonts/colors — a
// fidelity gap vs the same cell on-screen). For a BREAK-FREE value the helper
// must draw each run with its own font/color (ECMA-376 §18.4.4 r / §18.4.7 rPr)
// and anchor the single line with the cell's alignV-dependent baseline
// ('top' / 'middle' / 'bottom'), matching the legacy in-viewport single-line
// path it now shares — NOT the 'top'-only multi-line model used for hard breaks.

const BASE: CellFont = {
  bold: false,
  italic: false,
  underline: false,
  strike: false,
  size: 11,
  color: null,
  name: null,
};

interface FillTextCall {
  text: string;
  x: number;
  y: number;
  font: string;
  baseline: CanvasTextBaseline;
  fillStyle: string;
}

function makeRecordingCtx(): { ctx: CanvasRenderingContext2D; calls: FillTextCall[] } {
  let font = '11px sans-serif';
  let baseline: CanvasTextBaseline = 'alphabetic';
  let fillStyle = '#000';
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '11');
  const calls: FillTextCall[] = [];
  const ctx = {
    get font() { return font; },
    set font(v: string) { font = v; },
    get textBaseline() { return baseline; },
    set textBaseline(v: CanvasTextBaseline) { baseline = v; },
    get fillStyle() { return fillStyle; },
    set fillStyle(v: string) { fillStyle = v; },
    measureText: (s: string) => ({ width: [...s].length * px() }) as TextMetrics,
    fillText(text: string, x: number, y: number) {
      calls.push({ text, x, y, font, baseline, fillStyle });
    },
    save() {},
    restore() {},
    beginPath() {},
    moveTo() {},
    lineTo() {},
    stroke() {},
    strokeStyle: '#000' as string,
    lineWidth: 1,
    textAlign: 'left' as CanvasTextAlign,
  };
  return { ctx: ctx as unknown as CanvasRenderingContext2D, calls };
}

const CELL_H = 10000;
const PADDING_Y = 2;

function geom(alignV: 'top' | 'center' | 'bottom') {
  return {
    alignH: 'left' as const,
    alignV,
    cx: 0,
    cy: 0,
    cellW: 10000,
    cellH: CELL_H,
    leftPad: 3,
    paddingX: 3,
    paddingY: PADDING_Y,
  };
}

function draw(runs: Run[], alignV: 'top' | 'center' | 'bottom' = 'bottom'): FillTextCall[] {
  const { ctx, calls } = makeRecordingCtx();
  drawNonWrapRichText(ctx, runs, BASE, geom(alignV), 1, 1);
  return calls;
}

function rf(color: string | null, size = 11) {
  return { bold: false, italic: false, underline: false, strike: false, size, color };
}

describe('non-wrap break-free rich text mirrors the in-viewport per-run single-line path', () => {
  it('draws each run with its own color (off-screen pre-pass no longer collapses to base font)', () => {
    const calls = draw([
      { text: 'AA', font: rf('#FF0000') },
      { text: 'BB', font: rf('#00FF00') },
    ]);
    expect(calls.map((c) => c.text)).toEqual(['AA', 'BB']);
    // Per-run color: the two runs paint in their own colors, not one base color.
    expect(calls[0].fillStyle).not.toBe(calls[1].fillStyle);
    expect(calls[0].fillStyle).toContain('255'); // red channel of #FF0000
    expect(calls[1].fillStyle).toContain('255'); // green channel of #00FF00
  });

  it('draws each run at its own font size', () => {
    const calls = draw([
      { text: 'small', font: rf(null, 11) },
      { text: 'BIG', font: rf(null, 22) },
    ]);
    expect(calls.map((c) => c.text)).toEqual(['small', 'BIG']);
    // Each fillText is preceded by ctx.font set to that run's size.
    expect(/\b15px\b/.test(calls[0].font)).toBe(true); // 11pt → round(11 * 96/72) = 15px
    expect(/\b29px\b/.test(calls[1].font)).toBe(true); // 22pt → round(22 * 96/72) = 29px
  });

  it('lays the two runs side by side on one line (single line, not stacked)', () => {
    const calls = draw([
      { text: 'foo', font: rf(null) },
      { text: 'bar', font: rf(null) },
    ]);
    expect(calls).toHaveLength(2);
    expect(calls[0].y).toBe(calls[1].y); // same baseline
    expect(calls[1].x).toBeGreaterThan(calls[0].x);
  });

  it('anchors the single line with the alignV-dependent baseline, not the top-only model', () => {
    // alignV='bottom' (the cell default): the legacy single-line path uses a
    // 'bottom' baseline at cy + cellH - paddingY — the off-screen pre-pass must
    // match it so an off-screen-anchored merge sits where the on-screen cell does.
    const bottom = draw([{ text: 'X', font: rf(null) }], 'bottom');
    expect(bottom).toHaveLength(1);
    expect(bottom[0].baseline).toBe('bottom');
    expect(bottom[0].y).toBe(CELL_H - PADDING_Y); // 9998, not a reserved-line top

    const top = draw([{ text: 'X', font: rf(null) }], 'top');
    expect(top[0].baseline).toBe('top');
    expect(top[0].y).toBe(PADDING_Y);

    const center = draw([{ text: 'X', font: rf(null) }], 'center');
    expect(center[0].baseline).toBe('middle');
    expect(center[0].y).toBe(CELL_H / 2);
  });
});
