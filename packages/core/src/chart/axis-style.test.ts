import { describe, it, expect } from 'vitest';
import { axisLineWidthPx, resolveAxisLine, isCrossBetween } from './axis-style.js';

// EMU_PER_PT = 12700. A 1 pt line = 12700 EMU.
describe('axisLineWidthPx', () => {
  it('scales EMU width to px by ptToPx, floored at 0.5px', () => {
    expect(axisLineWidthPx(12700, 1)).toBe(1); // 1pt × ptToPx 1
    expect(axisLineWidthPx(12700, 2)).toBe(2); // 1pt × ptToPx 2 (dpr 2)
    expect(axisLineWidthPx(3175, 1)).toBe(0.5); // 0.25pt → floored to 0.5
  });
  it('falls back to a 1px hairline when width is absent', () => {
    expect(axisLineWidthPx(undefined, 2)).toBe(1);
    expect(axisLineWidthPx(null, 2)).toBe(1);
    expect(axisLineWidthPx(0, 2)).toBe(1);
  });
});

describe('resolveAxisLine', () => {
  it('prefixes the colour with # and defaults to #aaa', () => {
    expect(resolveAxisLine('888888', 12700, 1)).toEqual({ color: '#888888', width: 1 });
    expect(resolveAxisLine(undefined, undefined, 2)).toEqual({ color: '#aaa', width: 1 });
    expect(resolveAxisLine(null, 25400, 1)).toEqual({ color: '#aaa', width: 2 });
  });
});

describe('isCrossBetween', () => {
  it('is true unless explicitly midCat', () => {
    expect(isCrossBetween({ catAxisCrossBetween: 'between' })).toBe(true);
    expect(isCrossBetween({ catAxisCrossBetween: '' })).toBe(true);
    expect(isCrossBetween({ catAxisCrossBetween: 'midCat' })).toBe(false);
  });
});
