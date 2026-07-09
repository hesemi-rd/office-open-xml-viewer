import { describe, expect, it } from 'vitest';
import { resolveAnchorX, resolveAnchorY } from './anchor-geometry.js';

// resolveAnchorY reads only { scale, marginTop, marginBottom, pageH } of
// RenderState (via yContainer); cast a minimal stand-in like the other geometry
// tests. scale=1 so px == pt. pageH=800, top/bottom margins=72 ⇒ the "margin"
// container band is [72, 728], the "page" band is [0, 800].
interface MinState {
  scale: number;
  pageWidth: number;
  marginLeft: number;
  marginRight: number;
  contentX: number;
  contentW: number;
  marginTop: number;
  marginBottom: number;
  pageH: number;
}
const state: MinState = {
  scale: 1,
  pageWidth: 600,
  marginLeft: 60,
  marginRight: 40,
  contentX: 60,
  contentW: 500,
  marginTop: 72,
  marginBottom: 72,
  pageH: 800,
};

// resolveAnchorY(align, fromPara, offsetPt, heightPx, paragraphTopPx, state, relativeFrom)
const y = (align: string, relativeFrom: string, h = 100): number =>
  resolveAnchorY(align, false, 0, h, 0, state as never, relativeFrom);

describe('resolveAnchorY — ST_AlignV inside/outside (ECMA-376 §20.4.3.1)', () => {
  // The page-binding-relative inside/outside degrade to the top/bottom edge of
  // the container (odd-page approximation), mirroring resolveAnchorX's
  // inside→left / outside→right.
  //
  // NOTE (S-11): these assertions PIN the odd-page approximation, not the true
  // §20.4.3.1 page-parity behavior. Word flips the binding edge on EVEN pages
  // (inside→bottom, outside→top), which is not implemented — resolveAnchorY has
  // no page-index input. When that parity is added, update inside/outside here
  // to the page-relative expectations (the odd-page values below stay correct
  // for odd pages only).
  it('inside aligns to the container top edge (= top)', () => {
    expect(y('inside', 'page')).toBe(0); // page band top
    expect(y('inside', 'page')).toBe(y('top', 'page'));
    expect(y('inside', 'margin')).toBe(72); // margin band top
    expect(y('inside', 'margin')).toBe(y('top', 'margin'));
  });

  it('outside aligns to the container bottom edge (= bottom)', () => {
    expect(y('outside', 'page')).toBe(800 - 100); // page band bottom − height
    expect(y('outside', 'page')).toBe(y('bottom', 'page'));
    expect(y('outside', 'margin')).toBe(728 - 100); // margin band bottom − height
    expect(y('outside', 'margin')).toBe(y('bottom', 'margin'));
  });

  it('still resolves the explicit top/bottom/center cases', () => {
    expect(y('top', 'page')).toBe(0);
    expect(y('bottom', 'page')).toBe(700);
    expect(y('center', 'page')).toBe((800 - 100) / 2);
  });
});

describe('resolveAnchorX — wp:align choice ignores standalone posOffset', () => {
  it('right-aligns a standalone shape without adding anchorXPt/simplePos fallback', () => {
    const x = resolveAnchorX('right', true, 200, 100, state as never, 'margin', null, null);

    expect(x).toBe(460);
  });

  it('still adds the child offset when aligning a grouped shape by group width', () => {
    const x = resolveAnchorX('right', true, 20, 50, state as never, 'margin', null, 200);

    expect(x).toBe(380);
  });
});
