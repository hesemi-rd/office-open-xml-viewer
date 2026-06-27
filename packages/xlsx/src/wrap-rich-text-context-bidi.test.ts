import { describe, it, expect } from 'vitest';
import { drawWrappedRichText, layoutRichTextLines } from './renderer.js';
import { computeLineVisualOrder } from './bidi-line.js';
import type { CellFont, Run } from './types.js';
import type { RichCellGeom } from './renderer.js';

// ECMA-376 §18.8.1 readingOrder: 0/absent = Context → UAX#9 first-strong (P1–P3).
// A soft wrap does NOT start a new bidi paragraph, so all soft-wrapped display
// lines of one logical paragraph share a base direction. But a HARD break (LF /
// Alt+Enter, preserved in the run text per §18.4.12 t @xml:space) DOES start a
// new paragraph — under Context reading order each LF-delimited paragraph must
// resolve its OWN base direction from its OWN first strong character.
//
// `drawWrappedRichText` used to resolve ONE cell-wide first-strong direction over
// the whole joined run text and apply it to every LF paragraph, so a wrapped cell
// whose paragraph 1 is LTR-first and paragraph 2 is RTL-first rendered paragraph 2
// under the wrong (paragraph-1) base direction. The non-wrap hard-break path
// (`drawMultiLineRichText` → `drawRichLine`) already resolves per LF line; this
// covers the wrap path.

const BASE: CellFont = {
  bold: false,
  italic: false,
  underline: false,
  strike: false,
  size: 11,
  color: null,
  name: null,
};

interface FillTextCall { text: string; x: number; y: number; }

function makeRecordingCtx(): { ctx: CanvasRenderingContext2D; calls: FillTextCall[] } {
  let font = '11px sans-serif';
  const px = () => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '11');
  const calls: FillTextCall[] = [];
  const ctx = {
    get font() { return font; },
    set font(v: string) { font = v; },
    measureText: (s: string) => ({ width: [...s].length * px() }) as TextMetrics,
    fillText(text: string, x: number, y: number) { calls.push({ text, x, y }); },
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

// A very wide cell so nothing soft-wraps — only the explicit LF splits paragraphs.
const WIDE = 100000;
function geom(): RichCellGeom {
  return { alignH: 'left', alignV: 'top', cx: 0, cy: 0, cellW: WIDE, cellH: 1000, leftPad: 0, paddingX: 0, paddingY: 0 };
}

/** The texts painted on each distinct display line (a line per distinct y), in
 *  paint (= visual, left-to-right) order. */
function linesByY(calls: FillTextCall[]): { y: number; texts: string[] }[] {
  const ys = [...new Set(calls.map((c) => c.y))].sort((a, b) => a - b);
  return ys.map((y) => ({ y, texts: calls.filter((c) => c.y === y).map((c) => c.text) }));
}

/** The segments `drawWrappedRichText` lays out for a single-paragraph value,
 *  using the same tokenizer/widths so segment identity matches the draw path. */
function segmentsOf(text: string): { text: string }[] {
  const { ctx } = makeRecordingCtx();
  return layoutRichTextLines(ctx, [{ text }] as Run[], BASE, 1, WIDE).flatMap((l) => l.segments);
}

describe('drawWrappedRichText — Context base direction is per LF paragraph (§18.8.1 / UAX#9 P1–P3)', () => {
  it('paragraph 2 (RTL-first) is ordered under its OWN base, not paragraph 1’s (LTR)', () => {
    const P1 = 'ab';           // LTR first-strong
    const P2 = 'שלום abc';      // RTL first-strong, mixed RTL + LTR
    const runs: Run[] = [{ text: `${P1}\n${P2}` }];

    const { ctx, calls } = makeRecordingCtx();
    // readingOrder undefined = Context (first-strong per paragraph).
    drawWrappedRichText(ctx, runs, BASE, geom(), 1, 1, {});

    const lines = linesByY(calls);
    expect(lines).toHaveLength(2);
    expect(lines[0].texts).toEqual([P1]); // paragraph 1 unchanged

    const p2segs = segmentsOf(P2);
    const ownRtl = computeLineVisualOrder(p2segs, true);   // P2’s own base (Hebrew first-strong → RTL)
    const cellLtr = computeLineVisualOrder(p2segs, false); // the bug: cell-wide base from P1 → LTR
    // Precondition: this input actually exercises the bug (the two orders differ).
    expect(ownRtl.order).not.toEqual(cellLtr.order);

    const expected = ownRtl.order.map((i) => p2segs[i].text);
    expect(lines[1].texts).toEqual(expected);
  });

  it('a single paragraph (no hard break) still uses its first-strong base — unchanged', () => {
    const P = 'שלום abc'; // one paragraph: cell-wide == per-paragraph
    const { ctx, calls } = makeRecordingCtx();
    drawWrappedRichText(ctx, [{ text: P }], BASE, geom(), 1, 1, {});

    const lines = linesByY(calls);
    expect(lines).toHaveLength(1);
    const segs = segmentsOf(P);
    const expected = computeLineVisualOrder(segs, true).order.map((i) => segs[i].text);
    expect(lines[0].texts).toEqual(expected);
  });

  it('preserves a blank line between paragraphs in the drawn vertical advance', () => {
    // Control: "A\nB" advances one line height between A and B.
    const ctrl = makeRecordingCtx();
    drawWrappedRichText(ctrl.ctx, [{ text: 'A\nB' }], BASE, geom(), 1, 1, {});
    const cl = linesByY(ctrl.calls);
    expect(cl.map((l) => l.texts)).toEqual([['A'], ['B']]);
    const oneGap = cl[1].y - cl[0].y;

    // Subject: "A\n\nB" — the blank middle line paints nothing but still reserves
    // one line height, so B sits two line heights below A.
    const subj = makeRecordingCtx();
    drawWrappedRichText(subj.ctx, [{ text: 'A\n\nB' }], BASE, geom(), 1, 1, {});
    const sl = linesByY(subj.calls);
    expect(sl.map((l) => l.texts)).toEqual([['A'], ['B']]); // only 2 painted lines
    expect(sl[1].y - sl[0].y).toBeCloseTo(2 * oneGap, 5);
  });
});
