import { describe, it, expect } from 'vitest';
import { sliceHorizontalExtent, overlayPercent } from './highlight-rect.js';

/**
 * IX2 highlight geometry. `sliceHorizontalExtent` measures the two prefixes of a
 * run to place the highlight box flush with the drawn glyphs. A monospace-style
 * fake measurer (n chars → n×W px) makes the arithmetic checkable.
 */
const W = 7;
const mono = (s: string) => s.length * W;

describe('sliceHorizontalExtent', () => {
  it('measures a mid-run slice from the two prefix widths', () => {
    // "the quick" — slice [4,9) = "quick".
    const { x, width } = sliceHorizontalExtent('the quick', 4, 9, mono);
    expect(x).toBe(4 * W);
    expect(width).toBe(5 * W);
  });

  it('places a leading slice at x=0', () => {
    const { x, width } = sliceHorizontalExtent('Hello', 0, 3, mono);
    expect(x).toBe(0);
    expect(width).toBe(3 * W);
  });

  it('measures the whole run text for an end-anchored slice', () => {
    const { x, width } = sliceHorizontalExtent('Hello', 2, 5, mono);
    expect(x).toBe(2 * W);
    expect(width).toBe(3 * W);
  });

  it('covers the entire run for a [0, len) slice', () => {
    const { x, width } = sliceHorizontalExtent('Hello', 0, 5, mono);
    expect(x).toBe(0);
    expect(width).toBe(5 * W);
  });

  it('never returns a negative width', () => {
    // Degenerate slice (start === end).
    const { width } = sliceHorizontalExtent('Hello', 3, 3, mono);
    expect(width).toBe(0);
  });
});

describe('overlayPercent', () => {
  it('expresses a coordinate as a percentage of the basis', () => {
    expect(overlayPercent(480, 960)).toBe('50%');
    expect(overlayPercent(0, 960)).toBe('0%');
    expect(overlayPercent(960, 960)).toBe('100%');
  });

  it('is scale-invariant — a fraction reads the same regardless of the basis', () => {
    // The same fractional position yields the same % whatever the intended box
    // size is, which is why a %-placed overlay tracks a scaled canvas.
    expect(overlayPercent(240, 960)).toBe(overlayPercent(120, 480));
  });

  it('yields 0% for a non-positive basis instead of NaN%/Infinity%', () => {
    // Nothing laid out yet (basis 0) must not emit `NaN%`.
    expect(overlayPercent(10, 0)).toBe('0%');
    expect(overlayPercent(10, -5)).toBe('0%');
  });
});
