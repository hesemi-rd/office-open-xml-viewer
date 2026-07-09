import { describe, expect, it } from 'vitest';
import { buildSegments, layoutLines, type LayoutMathSeg } from './line-layout.js';
import type { RenderState } from './renderer.js';
import type { DocRun } from './types.js';

function makeLinearCtx(): CanvasRenderingContext2D {
  let font = '10px serif';
  const pxOf = (): number => parseFloat(/(\d+(?:\.\d+)?)px/.exec(font)?.[1] ?? '10');
  return {
    get font() { return font; },
    set font(v: string) { font = v; },
    letterSpacing: '0px',
    measureText: (s: string) => {
      const p = pxOf();
      return {
        width: [...s].length * p,
        fontBoundingBoxAscent: p * 0.8,
        fontBoundingBoxDescent: p * 0.2,
        actualBoundingBoxAscent: p * 0.8,
        actualBoundingBoxDescent: p * 0.2,
      } as TextMetrics;
    },
  } as unknown as CanvasRenderingContext2D;
}

describe('layoutLines math fallback', () => {
  it('measures inline OMML as plain text when the math renderer has not prepared an image', () => {
    const runs: DocRun[] = [
      {
        type: 'math',
        display: false,
        fontSize: 11,
        nodes: [
          { kind: 'radical', radicand: [{ kind: 'run', text: '9', style: 'italic' }] },
          { kind: 'run', text: '、', style: 'italic' },
          { kind: 'radical', radicand: [{ kind: 'run', text: '16', style: 'italic' }] },
        ],
      },
    ];
    const segs = buildSegments(runs, {} as RenderState);
    const lines = layoutLines(makeLinearCtx(), segs, 300, 0, 1);
    const math = lines[0].segments[0] as LayoutMathSeg;

    expect(math.fallbackText).toBe('√9、√16');
    expect(math.measuredWidth).toBeGreaterThan(0);
  });
});
