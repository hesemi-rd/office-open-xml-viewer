import { describe, it, expect } from 'vitest';
import { resolveTabWidths, type TabItem, type TabStopPx } from './tab-layout.js';

// ECMA-376 §21.1.2.1.x (a:tabLst / a:tab @pos @algn) — resolve each inline TAB
// segment's GAP WIDTH against the stop grid, in the READING frame (issue #916,
// generalising the #913 single-cell RTL semantics to N cells).
//
// All coordinates are reading-frame px measured from the LEADING text-inset edge
// (LTR: left inset; RTL: right inset). Content starts at `startPen` (the leading
// indent). A tab's gap makes the FOLLOWING cell end (`r`/`dec`), centre (`ctr`),
// or start (`l`) on the selected stop. The same math serves LTR and RTL: the
// physical mirroring falls out of the visual L2 reorder + cumulative draw
// (mirrors docx `layoutBidiTabStops`). Non-tab items keep their input width.

const S = (pos: number, algn: string): TabStopPx => ({ pos, algn });
const text = (width: number): TabItem => ({ isTab: false, width });
const tab = (): TabItem => ({ isTab: true, width: 0 });

describe('resolveTabWidths — inline tab gap resolution (§21.1.2.1.x)', () => {
  it('LTR single right stop: gap makes the following cell END on the stop', () => {
    // [TAB][text(40)], stop r@400 ⇒ cell right edge on 400 ⇒ gap 360.
    const out = resolveTabWidths([tab(), text(40)], [S(400, 'r')], 0, 600, 0);
    expect(out).toEqual([360, 40]);
  });

  it('supports N tab cells per line (the #916 item-1 fix)', () => {
    // [text(50)][TAB][text(40)][TAB][text(40)], stops r@200, r@400.
    const items = [text(50), tab(), text(40), tab(), text(40)];
    const out = resolveTabWidths(items, [S(200, 'r'), S(400, 'r')], 0, 600, 0);
    // gaps 110 and 160 ⇒ cells end exactly on 200 and 400.
    expect(out).toEqual([50, 110, 40, 160, 40]);
    // cumulative pen reaches each stop:
    expect(50 + out[1] + 40).toBeCloseTo(200, 6); // cell 1 right edge
    expect(50 + out[1] + 40 + out[3] + 40).toBeCloseTo(400, 6); // cell 2 right edge
  });

  it('centre stop centres the following cell on the stop', () => {
    // [TAB][text(40)], ctr@400 ⇒ gap = 400 − 40/2 = 380.
    const out = resolveTabWidths([tab(), text(40)], [S(400, 'ctr')], 0, 600, 0);
    expect(out).toEqual([380, 40]);
  });

  it('start (l) tab MATERIALISES the gap (the #916 item-3 fix)', () => {
    // [TAB][text(30)], l@100, pen 0 ⇒ gap 100 (previously never rendered).
    const out = resolveTabWidths([tab(), text(30)], [S(100, 'l')], 0, 600, 0);
    expect(out).toEqual([100, 30]);
  });

  it('dec tab aligns like right (frac 1; no decimal-point split) and renders the gap', () => {
    const out = resolveTabWidths([tab(), text(40)], [S(400, 'dec')], 0, 600, 0);
    expect(out).toEqual([360, 40]);
  });

  it('no reachable stop → the tab collapses to the fallback gap (degrade to a space)', () => {
    const out = resolveTabWidths([tab(), text(30)], [], 0, 600, 20);
    expect(out).toEqual([20, 30]);
  });

  it('clamps a cell that would overflow the trailing text edge (#835)', () => {
    // stop r@700 in a 600px area, cell 40 ⇒ target 660 → clamp so cell ends at 600.
    const out = resolveTabWidths([tab(), text(40)], [S(700, 'r')], 0, 600, 0);
    expect(out).toEqual([560, 40]); // cell spans [560, 600]
  });

  it('never moves the pen backwards: a right stop behind the pen collapses the tab', () => {
    // start@100 then A(20) leaves pen at 120; end@140 with a 40-wide cell wants
    // target 100 < pen ⇒ gap 0 (the cell continues from the pen). Matches docx
    // layoutBidiTabStops and the #913 case (5) reachability contract.
    const items = [tab(), text(20), tab(), text(40)];
    const out = resolveTabWidths(items, [S(100, 'l'), S(140, 'r')], 50, 600, 0);
    expect(out).toEqual([50, 20, 0, 40]);
  });

  it('selects the nearest stop strictly past the pen, custom order-agnostic', () => {
    // Stops given out of order; pen 0 picks 200, then pen past 200 picks 400.
    const items = [tab(), text(10), tab(), text(10)];
    const out = resolveTabWidths(items, [S(400, 'l'), S(200, 'l')], 0, 600, 0);
    expect(out).toEqual([200, 10, 190, 10]);
  });
});
