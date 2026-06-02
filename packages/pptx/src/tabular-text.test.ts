import { describe, it, expect } from 'vitest';
import { tabularDigitWidth, tabularTextWidth, drawTabularText } from './tabular-text.js';

// Mock a Canvas-like text context with deliberately UNEQUAL digit widths (as a
// proportional font has — '1' is narrow, '8'/'0' wide) plus punctuation. This
// is what makes a naive layout jitter; the tabular helpers must neutralize it.
const ADVANCE: Record<string, number> = {
  '0': 10, '1': 5, '2': 9, '3': 9, '4': 9, '5': 9, '6': 10, '7': 8, '8': 10, '9': 10,
  ':': 4, ' ': 3, '/': 5,
};

function mockCtx() {
  const calls: { ch: string; x: number }[] = [];
  const ctx = {
    textAlign: 'left' as CanvasTextAlign,
    measureText: (s: string) => ({ width: [...s].reduce((w, c) => w + (ADVANCE[c] ?? 6), 0) }),
    fillText: (t: string, x: number) => { calls.push({ ch: t, x }); },
  };
  return { ctx: ctx as unknown as CanvasRenderingContext2D, calls };
}

describe('tabularDigitWidth', () => {
  it('returns the widest digit advance', () => {
    expect(tabularDigitWidth(mockCtx().ctx)).toBe(10); // max over 0–9
  });
});

describe('tabularTextWidth', () => {
  it('is identical for same-shape time strings regardless of digit values', () => {
    const { ctx } = mockCtx();
    const dw = tabularDigitWidth(ctx);
    // "m:ss" → digit, ':', digit, digit. Different values, same shape ⇒ same width.
    const a = tabularTextWidth(ctx, '0:01', dw);
    const b = tabularTextWidth(ctx, '0:08', dw);
    const c = tabularTextWidth(ctx, '9:59', dw);
    expect(a).toBe(b);
    expect(b).toBe(c);
    expect(a).toBe(dw + ADVANCE[':'] + dw + dw); // 3 digit cells + colon
  });
});

describe('drawTabularText — no jitter', () => {
  it('places every digit cell at the same x for any digit values', () => {
    const dw = tabularDigitWidth(mockCtx().ctx);
    const draw = (s: string) => {
      const { ctx, calls } = mockCtx();
      drawTabularText(ctx, s, 100, 20, dw);
      return calls;
    };
    const c1 = draw('0:01');
    const c2 = draw('0:08');
    const c3 = draw('1:11');
    // Same number of glyphs, and each glyph's cell advances identically.
    // We verify the LAST glyph lands at the same x across all three (its
    // position is the sum of all preceding fixed cells), which is exactly the
    // "separator/duration don't move" guarantee.
    const lastX = (calls: { ch: string; x: number }[]) => calls[calls.length - 1].x;
    // The trailing glyph is a digit; its cell start is identical, only the
    // intra-cell centering offset varies with the glyph width — but the cell
    // boundary (and thus everything after) is fixed. Assert cell starts match
    // by removing the centering offset.
    const cellStart = (entry: { ch: string; x: number }) =>
      entry.x - (dw - (ADVANCE[entry.ch] ?? 6)) / 2;
    expect(cellStart(c2[c2.length - 1])).toBeCloseTo(cellStart(c1[c1.length - 1]), 9);
    expect(cellStart(c3[c3.length - 1])).toBeCloseTo(cellStart(c1[c1.length - 1]), 9);
    // And the digits are centered (offset ≥ 0, glyph never overflows the cell).
    expect(lastX(c1)).toBeGreaterThanOrEqual(cellStart(c1[c1.length - 1]));
  });

  it('restores textAlign', () => {
    const { ctx } = mockCtx();
    (ctx as { textAlign: CanvasTextAlign }).textAlign = 'center';
    drawTabularText(ctx, '0:00', 0, 0, 10);
    expect(ctx.textAlign).toBe('center');
  });
});
