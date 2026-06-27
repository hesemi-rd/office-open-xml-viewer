import { describe, it, expect } from 'vitest';
import { drawNonWrapRichText } from './renderer.js';
import type { CellFont, Run } from './types.js';

// ECMA-376 §18.8.1 (CT_CellAlignment) @wrapText governs only soft-wrapping. It
// says nothing about hard breaks — but Excel still renders a break authored with
// Alt+Enter (a literal LF, U+000A, preserved in the run text via §18.4.12 t
// xml:space="preserve"; the run is §18.4.4 r) as a SEPARATE line even with
// wrapText off. That is undocumented runtime behavior, matched for parity: the
// plain-text non-wrap path already splits on LF ("to match Excel's behavior"),
// while the rich-text (mixed-font runs) non-wrap path used to lay every run on
// one horizontal line (`runX += width`), silently dropping every hard break.
// This is the rich-text sibling of that plain-text behavior and of the wrap-path
// blank-line fix (PR #585).
//
// The cell path derives line height analytically (font size × 1.2 via vMetricPx),
// not from font metrics, so the line spacing is visible without an asymmetric
// font stub: each line — including a blank one from consecutive breaks — reserves
// the SAME 1.2-em single-line height, exactly as the wrap path and Excel do.

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
}

function makeRecordingCtx(): { ctx: CanvasRenderingContext2D; calls: FillTextCall[] } {
  let font = '11px sans-serif';
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '11');
  const calls: FillTextCall[] = [];
  const ctx = {
    get font() {
      return font;
    },
    set font(v: string) {
      font = v;
    },
    measureText: (s: string) => ({ width: [...s].length * px() }) as TextMetrics,
    fillText(text: string, x: number, y: number) {
      calls.push({ text, x, y });
    },
    // underline / strike decoration stubs (not exercised by these LTR plain runs)
    save() {},
    restore() {},
    beginPath() {},
    moveTo() {},
    lineTo() {},
    stroke() {},
    fillStyle: '#000' as string,
    strokeStyle: '#000' as string,
    lineWidth: 1,
    textAlign: 'left' as CanvasTextAlign,
    textBaseline: 'alphabetic' as CanvasTextBaseline,
  };
  return { ctx: ctx as unknown as CanvasRenderingContext2D, calls };
}

// A top-anchored, wrap:none cell. Top anchoring isolates the line-height
// contribution: each line's top y is the cumulative sum of the heights of the
// lines above it, so the gap between two drawn lines reports exactly how much
// the lines between them reserved.
const GEOM = {
  alignH: 'left' as const,
  alignV: 'top' as const,
  cx: 0,
  cy: 0,
  cellW: 10000, // wide enough that nothing soft-wraps (only LF splits)
  cellH: 10000,
  leftPad: 3,
  paddingX: 3,
  paddingY: 2,
};

function draw(runs: Run[]): FillTextCall[] {
  const { ctx, calls } = makeRecordingCtx();
  drawNonWrapRichText(ctx, runs, BASE, GEOM, 1, 1);
  return calls;
}

function run(text: string, font?: Run['font']): Run {
  return font ? { text, font } : { text };
}

describe('non-wrap rich text honors hard breaks (§18.8.1 wrapText off / §18.4.4 r)', () => {
  it('a single run with an embedded LF draws two lines', () => {
    const calls = draw([run('A\nB')]);
    expect(calls.map((c) => c.text)).toEqual(['A', 'B']);
    expect(calls[1].y).toBeGreaterThan(calls[0].y);
  });

  it('a hard break across two (mixed-font) runs draws two lines', () => {
    // The break sits at the boundary between two differently-fonted runs:
    // "Hello" on line 1, "World" on line 2. Previously both ran together on a
    // single horizontal line.
    const calls = draw([
      run('Hello\n', { bold: true, italic: false, underline: false, strike: false, size: 11 }),
      run('World', { bold: false, italic: false, underline: false, strike: false, size: 11 }),
    ]);
    expect(calls.map((c) => c.text)).toEqual(['Hello', 'World']);
    expect(calls[1].y).toBeGreaterThan(calls[0].y);
  });

  it('the two lines are spaced exactly one single-line height apart', () => {
    const calls = draw([run('A\nB')]);
    // vMetricPx(11, cs=1, 1.2) = round(11 * (96/72) * 1.2) = round(17.6) = 18
    const PT_TO_PX = 96 / 72;
    const lineH = Math.round(11 * PT_TO_PX * 1.2);
    expect(calls[1].y - calls[0].y).toBe(lineH);
  });

  it('a blank line from consecutive breaks reserves one line height', () => {
    // "A\n\nB": A on line 1, a BLANK line 2, B on line 3. The blank middle line
    // must push B down by one extra single-line height (gap = 2 line heights),
    // not collapse it back next to A.
    const ctrl = draw([run('A\nB')]);
    const subj = draw([run('A\n\nB')]);
    const gapCtrl = ctrl[1].y - ctrl[0].y;
    const gapSubj = subj[1].y - subj[0].y;
    expect(subj.map((c) => c.text)).toEqual(['A', 'B']);
    expect(gapSubj).toBe(gapCtrl * 2);
  });

  it('a cell with no hard break still draws on a single line (no regression)', () => {
    const calls = draw([
      run('foo', { bold: false, italic: false, underline: false, strike: false, size: 11 }),
      run('bar', { bold: true, italic: false, underline: false, strike: false, size: 11 }),
    ]);
    expect(calls.map((c) => c.text)).toEqual(['foo', 'bar']);
    // Same baseline → same y; the two runs sit side by side on one line.
    expect(calls[0].y).toBe(calls[1].y);
    expect(calls[1].x).toBeGreaterThan(calls[0].x);
  });
});
