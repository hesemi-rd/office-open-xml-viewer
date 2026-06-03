import { describe, it, expect } from 'vitest';
import { svgExtents, recolorSvg } from './mathjax';

// mathMLToSvg loads MathJax from a CDN and needs a DOM, so it is exercised in the
// browser (Storybook / VRT), not here. These cover the pure helpers.

describe('svgExtents', () => {
  it('parses the viewBox into baseline-relative em extents', () => {
    const e = svgExtents('<svg viewBox="0 -1642.5 9178 2338.5"></svg>');
    expect(e.widthEm).toBeCloseTo(9.178, 3);
    expect(e.ascentEm).toBeCloseTo(1.6425, 3);
    expect(e.descentEm).toBeCloseTo(0.696, 3);
  });

  it('returns zeros when no viewBox is present', () => {
    expect(svgExtents('<svg></svg>')).toEqual({ widthEm: 0, ascentEm: 0, descentEm: 0 });
  });
});

describe('recolorSvg', () => {
  it('replaces currentColor', () => {
    expect(recolorSvg('fill="currentColor" stroke="currentColor"', '#f00')).toBe(
      'fill="#f00" stroke="#f00"',
    );
  });
});
