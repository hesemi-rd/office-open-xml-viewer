import { describe, expect, it } from 'vitest';
import { __test_resolveAnchorBox, type RenderState } from './renderer.js';
import type { ImageRun } from './types.js';

// Pin the wiring of `<wp:positionH/V>@relativeFrom` (ECMA-376 §20.4.3.2 /
// §20.4.3.5) through ImageRun into resolveAnchorBox. The placement math itself
// is covered by anchor-align.test.ts via resolveAnchorY; this file asserts that
// `anchorYRelativeFrom` / `anchorXRelativeFrom` reach the geometry helper so
// e.g. `relativeFrom="margin"` + `align="top"` pins the image to the body's
// top content margin instead of the page top (Y=0, inside the top margin).
//
// resolveAnchorBox reads only { scale, marginLeft/Right/Top/Bottom, pageH,
// pageWidth } from RenderState; cast a minimal stand-in like anchor-align.test.
const state = {
  scale: 1,
  marginLeft: 96,
  marginRight: 96,
  marginTop: 72,
  marginBottom: 72,
  pageH: 800,
  pageWidth: 612,
} as unknown as RenderState;

const baseImg: ImageRun = {
  imagePath: 'word/media/image1.png',
  mimeType: 'image/png',
  widthPt: 100,
  heightPt: 60,
  anchor: true,
  anchorXPt: 0,
  anchorYPt: 0,
  // No "fromMargin" hint — the bug repro path: the relativeFrom string must
  // do the work on its own.
  anchorXFromMargin: false,
  anchorYFromPara: false,
};

describe('resolveAnchorBox — positionV/@relativeFrom wiring (ECMA-376 §20.4.3.5)', () => {
  it('relativeFrom="margin" + align="top" pins Y to marginTop (sample-11 image arrow)', () => {
    // Bug 4a: without anchorYRelativeFrom plumbed through, this image lands at
    // Y=0 (inside the top page margin). The fix routes the raw "margin"
    // string into yContainer so resolveAnchorY returns the content-margin top.
    const img: ImageRun = {
      ...baseImg,
      anchorYAlign: 'top',
      anchorYRelativeFrom: 'margin',
    };
    const box = __test_resolveAnchorBox(img, state, 0);
    expect(box.y).toBe(state.marginTop); // 72, NOT 0
  });

  it('relativeFrom="page" + align="top" still lands at Y=0 (page band top)', () => {
    // Sanity: when relativeFrom is page, top alignment hits the page top edge.
    // Pins that the wire-up doesn't smuggle a margin default in.
    const img: ImageRun = {
      ...baseImg,
      anchorYAlign: 'top',
      anchorYRelativeFrom: 'page',
    };
    expect(__test_resolveAnchorBox(img, state, 0).y).toBe(0);
  });

  it('absent anchorYRelativeFrom + anchorYFromPara=false ⇒ legacy page-relative offset', () => {
    // Back-compat: an anchor that omitted positionV (no relativeFrom string,
    // no fromPara hint) keeps the previous "Y = anchorYPt as page-absolute"
    // behavior, even with align="top".
    const img: ImageRun = {
      ...baseImg,
      anchorYAlign: 'top',
      // anchorYRelativeFrom omitted
    };
    expect(__test_resolveAnchorBox(img, state, 0).y).toBe(0);
  });
});

describe('resolveAnchorBox — positionH/@relativeFrom wiring (ECMA-376 §20.4.3.2)', () => {
  it('relativeFrom="margin" + align="left" pins X to marginLeft', () => {
    const img: ImageRun = {
      ...baseImg,
      anchorXAlign: 'left',
      anchorXRelativeFrom: 'margin',
    };
    expect(__test_resolveAnchorBox(img, state, 0).x).toBe(state.marginLeft); // 96
  });

  it('relativeFrom="page" + align="left" lands at X=0', () => {
    const img: ImageRun = {
      ...baseImg,
      anchorXAlign: 'left',
      anchorXRelativeFrom: 'page',
    };
    expect(__test_resolveAnchorBox(img, state, 0).x).toBe(0);
  });
});
