import { describe, it, expect } from 'vitest';
import { paraShadingRect } from './renderer.js';
import type { ParagraphBorders } from './types';

// ECMA-376 §17.3.1.31 — paragraph shading fills the border BOX: it extends to
// each present border, which §17.3.1.7 offsets outward by `w:space`. The fill
// must reach the border, not stop `space` short of it (sample-11: a right border
// with space=4 pt left the gray box detached from its border).

const edge = (space: number, style = 'single') =>
  ({ style, width: 4, color: '000000', space } as unknown as NonNullable<ParagraphBorders['right']>);

const noEdges: ParagraphBorders = { top: null, bottom: null, left: null, right: null, between: null };

describe('paraShadingRect (§17.3.1.31 shading fills the border box)', () => {
  it('extends the fill by a present border edge’s space (matching drawParaBorders)', () => {
    // Right border, space=4 pt, scale=2 → the fill widens by 4*2=8 px on the right.
    const b: ParagraphBorders = { ...noEdges, right: edge(4) };
    const r = paraShadingRect(100, 50, 300, 20, b, undefined, 2);
    expect(r.x).toBe(100);           // no left border → left unchanged
    expect(r.y).toBe(50);
    expect(r.w).toBe(300 + 8);       // right border space 4pt × scale 2 = 8 px
    expect(r.h).toBe(20);
  });

  it('extends every present edge; a `none`-style edge and absent borders do not', () => {
    const b: ParagraphBorders = {
      ...noEdges,
      left: edge(2),
      right: edge(3),
      top: edge(1),
      bottom: edge(5, 'none'), // none → no ink → no extension
    };
    const r = paraShadingRect(100, 50, 300, 20, b, undefined, 1);
    expect(r.x).toBe(100 - 2);       // left space 2
    expect(r.w).toBe(300 + 2 + 3);   // left 2 + right 3
    expect(r.y).toBe(50 - 1);        // top space 1
    expect(r.h).toBe(20 + 1 + 0);    // top 1 + bottom 0 (none)
  });

  it('returns the content box unchanged when there are no borders', () => {
    expect(paraShadingRect(10, 20, 30, 40, null, undefined, 2)).toEqual({ x: 10, y: 20, w: 30, h: 40 });
  });

  it('honors a merged run: suppressed top uses `between`; suppressed bottom does not extend', () => {
    const b: ParagraphBorders = { ...noEdges, top: edge(9), between: edge(2), bottom: edge(7) };
    // suppressTop → top edge gives way to `between` (space 2); suppressBottom → no bottom extension.
    const r = paraShadingRect(0, 0, 100, 10, b, { suppressTop: true, suppressBottom: true }, 1);
    expect(r.y).toBe(-2);            // between space 2 (not top's 9)
    expect(r.h).toBe(10 + 2 + 0);    // top(between) 2 + bottom 0 (suppressed)
  });
});
