import { describe, expect, it } from 'vitest';
import { resolveAnchorX, xContainer } from './anchor-geometry.js';

// ECMA-376 §20.4.3.4 ST_RelFromH: `column` = "relative to the extents of the
// COLUMN which contains its anchor" (NOT the page margins). The renderer tracks
// the current text column as `state.contentX` / `state.contentW` (already in px):
//   - body level, single column  → the section margin band
//   - a multi-column section      → the specific column band the anchor sits in
//   - inside a table cell         → the CELL's inner text box (renderCell sets
//     contentX = cellX + cellMarginLeft, contentW = cellW − margins)
// So `column` must resolve against contentX/contentW, letting a header-logo
// anchor authored `relativeFrom="column"` inside an RTL bidi cell land in that
// cell's column instead of being degraded to the page margin band.
//
// `character` shares this behavior (§20.4.3.4: relative to the anchor's run
// content position; with no run-relative data we degrade it to the containing
// column, the closest available base — see anchor-geometry.ts).
//
// xContainer reads { scale, pageWidth, marginLeft, marginRight, contentX,
// contentW } of RenderState; a minimal stand-in is cast like the other geometry
// tests. contentX/contentW are ALREADY scaled (px), unlike marginLeft etc.
interface MinState {
  scale: number;
  pageWidth: number;
  marginLeft: number;
  marginRight: number;
  contentX: number;
  contentW: number;
}

// Page 600 wide, margins 60/40 ⇒ margin band [60, 560]. But the current text
// column (contentX/contentW) is a NARROW cell box [400, 550] — deliberately
// different from the margin band so a wrong degrade-to-margin is observable.
const cellState: MinState = {
  scale: 1,
  pageWidth: 600,
  marginLeft: 60,
  marginRight: 40,
  contentX: 400,
  contentW: 150,
};

describe('xContainer — column / character resolve to the text column (ECMA-376 §20.4.3.4)', () => {
  it('column uses contentX/contentW, not the page margin band', () => {
    expect(xContainer('column', false, cellState as never)).toEqual({ start: 400, end: 550 });
  });

  it('character also degrades to the containing column', () => {
    expect(xContainer('character', false, cellState as never)).toEqual({ start: 400, end: 550 });
  });

  it('margin still uses the page margin band (unchanged)', () => {
    expect(xContainer('margin', false, cellState as never)).toEqual({ start: 60, end: 560 });
  });

  it('page still uses the full page (unchanged)', () => {
    expect(xContainer('page', false, cellState as never)).toEqual({ start: 0, end: 600 });
  });

  it('resolveAnchorX places a column-relative offset from the column left edge', () => {
    // Offset 100 pt from the cell column's left extent (400) ⇒ 500.
    const x = resolveAnchorX(
      undefined, // no align
      false,
      100, // anchorXPt
      30, // widthPx (unused without align)
      cellState as never,
      'column',
      null,
      null,
    );
    expect(x).toBe(500);
  });

  it('body-level column equals the margin band when contentX == marginLeft', () => {
    // At body level contentX/contentW are the section margin band, so column and
    // margin coincide — no regression for single-column body anchors.
    const bodyState: MinState = {
      scale: 1,
      pageWidth: 600,
      marginLeft: 60,
      marginRight: 40,
      contentX: 60,
      contentW: 500,
    };
    expect(xContainer('column', false, bodyState as never)).toEqual(
      xContainer('margin', false, bodyState as never),
    );
  });
});
