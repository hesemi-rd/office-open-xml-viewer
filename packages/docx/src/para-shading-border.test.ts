import { describe, it, expect } from 'vitest';
import {
  paragraphBorderContentBox,
  paraBorderSegments,
  paraShadingRect,
  paintedParagraphHeight,
} from './renderer.js';
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

describe('paragraph border joins (§17.3.1.17)', () => {
  it('includes a hanging first-line/list marker on the paragraph start side', () => {
    expect(paragraphBorderContentBox(100, 500, 36, 0, -36, false)).toEqual({ x: 100, w: 500 });
    expect(paragraphBorderContentBox(100, 500, 0, 36, -36, true)).toEqual({ x: 100, w: 500 });
  });

  it('does not expand toward a positive first-line indent', () => {
    expect(paragraphBorderContentBox(100, 500, 36, 18, 12, false)).toEqual({ x: 136, w: 446 });
  });

  it('spans horizontal edges to the vertical border positions and vertical edges between them', () => {
    const borders: ParagraphBorders = {
      ...noEdges,
      top: edge(1),
      bottom: edge(2),
      left: edge(4),
      right: edge(5),
    };

    const segments = paraBorderSegments(10, 20, 100, 30, borders, undefined, 1);
    expect(segments.map(({ side, x1, y1, x2, y2 }) => ({ side, x1, y1, x2, y2 }))).toEqual([
      { side: 'top', x1: 6, y1: 19, x2: 115, y2: 19 },
      { side: 'bottom', x1: 6, y1: 52, x2: 115, y2: 52 },
      { side: 'left', x1: 6, y1: 19, x2: 6, y2: 52 },
      { side: 'right', x1: 115, y1: 19, x2: 115, y2: 52 },
    ]);
  });

  it('keeps a merged run open at an inner join when no between edge exists', () => {
    const borders: ParagraphBorders = {
      ...noEdges,
      top: edge(1),
      bottom: edge(2),
      left: edge(4),
      right: edge(5),
    };

    const segments = paraBorderSegments(10, 20, 100, 30, borders, { suppressBottom: true }, 1);
    expect(segments.find((segment) => segment.side === 'bottom')).toBeUndefined();
    expect(segments.find((segment) => segment.side === 'left')).toMatchObject({ y1: 19, y2: 50 });
    expect(segments.find((segment) => segment.side === 'right')).toMatchObject({ y1: 19, y2: 50 });
  });
});

// The shading fill HEIGHT must equal the paragraph border height, which
// renderParagraph measures AFTER the draw loop as `state.y − textAreaTopY`. Since
// shading is the background (painted BEFORE the loop), the height is pre-computed
// by paintedParagraphHeight, which replays the loop's exact per-line advancement
// (drawParagraphLine's ONLY two state.y mutations: the `topY` float-clearance
// max-jump, then `state.y += lineHForLine(line)`). These tests pin that the
// pre-computed height equals a hand-replayed advance for the normal, float and
// page-slice cases — the by-construction equality that makes the fill meet the
// bottom border (PR #641 fixed the other three edges; this fixes the bottom).
describe('paintedParagraphHeight (§17.3.1.7 shading height == border height)', () => {
  // Each test line carries its own box height `h`; the resolver returns it, so
  // the sum is explicit and independent of the real lineSpacing/docGrid math.
  type L = { topY?: number; h: number };
  const H = (l: L) => l.h;

  it('normal (no float, no slice): sums every line — same as the old naive total', () => {
    const lines: L[] = [{ h: 10 }, { h: 12 }, { h: 8 }];
    // textAreaTopY offset must NOT leak into the returned height.
    expect(paintedParagraphHeight(lines, 0, lines.length, 100, H)).toBe(30);
    expect(paintedParagraphHeight(lines, 0, lines.length, 0, H)).toBe(30);
  });

  it('float clearance: a line whose topY jumps past the natural flow grows the height', () => {
    // textAreaTopY=0. Replay: line0 y→10; line1 topY 100 > 10 → y=100, +12 → 112;
    // line2 → 120. The naive Σh would be 30, stopping short of the bottom border.
    const lines: L[] = [{ h: 10 }, { topY: 100, h: 12 }, { h: 8 }];
    expect(paintedParagraphHeight(lines, 0, lines.length, 0, H)).toBe(120);
    // A topY that does NOT exceed the running y is ignored (no backward jump).
    const noJump: L[] = [{ h: 10 }, { topY: 5, h: 12 }];
    expect(paintedParagraphHeight(noJump, 0, noJump.length, 0, H)).toBe(22);
  });

  it('page slice: only [sliceStart, paintEnd) contributes — no overfill past the slice', () => {
    const lines: L[] = [{ h: 10 }, { h: 12 }, { h: 8 }, { h: 6 }, { h: 4 }];
    // Slice [1,3): lines 12 + 8 = 20, measured from textAreaTopY (not the whole 40).
    expect(paintedParagraphHeight(lines, 1, 3, 50, H)).toBe(20);
    // paintEnd capped below lines.length still only sums the painted span.
    expect(paintedParagraphHeight(lines, 0, 2, 0, H)).toBe(22);
  });
});
