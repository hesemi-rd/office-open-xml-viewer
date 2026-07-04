import { describe, it, expect } from 'vitest';
import { axisLineWidthPx, resolveAxisLine, resolveGridline, isCrossBetween } from './axis-style.js';

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

describe('resolveGridline', () => {
  it('defaults to the faint #e0e0e0 / 0.5px hairline when unstyled (byte-stable)', () => {
    expect(resolveGridline(undefined, undefined, 1)).toEqual({ color: '#e0e0e0', width: 0.5 });
    expect(resolveGridline(null, null, 2)).toEqual({ color: '#e0e0e0', width: 0.5 });
  });
  it('prefixes an explicit gridline colour with # (accent3 → #8FA878)', () => {
    // axisLineWidthPx floors the PT value at 0.5 before scaling: 0.25pt → 0.5,
    // then × ptToPx. So ptToPx=1 → 0.5px, ptToPx=2 → 1px.
    expect(resolveGridline('8FA878', 3175, 1)).toEqual({ color: '#8FA878', width: 0.5 });
    expect(resolveGridline('8FA878', 3175, 2)).toEqual({ color: '#8FA878', width: 1 });
  });
  it('scales an explicit gridline width through axisLineWidthPx', () => {
    expect(resolveGridline('808080', 12700, 1)).toEqual({ color: '#808080', width: 1 }); // 1pt
    expect(resolveGridline('808080', 25400, 2)).toEqual({ color: '#808080', width: 4 }); // 2pt × dpr2
  });
  it('keeps the 0.5px hairline when a colour is set but no width', () => {
    expect(resolveGridline('8FA878', undefined, 1)).toEqual({ color: '#8FA878', width: 0.5 });
  });
});

describe('isCrossBetween', () => {
  it('is true unless explicitly midCat', () => {
    expect(isCrossBetween({ catAxisCrossBetween: 'between' })).toBe(true);
    expect(isCrossBetween({ catAxisCrossBetween: '' })).toBe(true);
    expect(isCrossBetween({ catAxisCrossBetween: 'midCat' })).toBe(false);
  });
});
