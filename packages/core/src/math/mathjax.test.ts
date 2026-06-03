import { describe, it, expect } from 'vitest';
import { mathMLToSvg, svgExtents, recolorSvg } from './mathjax';
import { mathToMathML } from './mathml';
import type { MathNode } from '../types/math';

describe('mathMLToSvg (MathJax)', () => {
  it('renders a fraction with a radical to SVG with positive extents', async () => {
    const nodes: MathNode[] = [
      { kind: 'run', text: 'x', style: 'italic' },
      { kind: 'run', text: '=', style: 'roman' },
      {
        kind: 'fraction',
        num: [
          { kind: 'run', text: '-b', style: 'italic' },
          { kind: 'radical', radicand: [{ kind: 'run', text: 'c', style: 'italic' }] },
        ],
        den: [{ kind: 'run', text: '2a', style: 'italic' }],
      },
    ];
    const out = await mathMLToSvg(mathToMathML(nodes, true));
    expect(out.svg).toContain('<svg');
    expect(out.svg).toContain('<path'); // inlined glyph outlines
    expect(out.widthEm).toBeGreaterThan(0);
    expect(out.ascentEm).toBeGreaterThan(0);
    expect(out.descentEm).toBeGreaterThan(0);
  }, 20000);

  it('svgExtents parses the viewBox', () => {
    const svg = '<svg viewBox="0 -1642.5 9178 2338.5"></svg>';
    const e = svgExtents(svg);
    expect(e.widthEm).toBeCloseTo(9.178, 3);
    expect(e.ascentEm).toBeCloseTo(1.6425, 3);
    expect(e.descentEm).toBeCloseTo(0.696, 3);
  });

  it('recolorSvg replaces currentColor', () => {
    expect(recolorSvg('fill="currentColor"', '#f00')).toBe('fill="#f00"');
  });
});
