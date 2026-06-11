import { describe, it, expect } from 'vitest';
import { computeCommentPopupPosition } from './comment-popup.js';

/**
 * Unit tests for the pure popup-position calculator. The popup follows Excel's
 * note placement: by default it sits just past the top-right corner of the
 * commented cell. When it would overflow the viewport it flips to the opposite
 * side (right→left) and/or clamps to stay fully visible.
 *
 * Geometry is in canvasArea CSS-pixel space (the same space getCellRect and the
 * selection overlay use). For RTL sheets the *cell rect* is already mirrored by
 * the caller (via screenX), so the default direction is reversed: the popup
 * opens toward the cell's left so it grows into the sheet body, not off-screen.
 */
describe('computeCommentPopupPosition', () => {
  const GAP = 8; // must match the implementation's cell↔popup gap

  const cell = { x: 100, y: 60, w: 64, h: 20 };
  const popup = { w: 280, h: 90 };
  const viewport = { w: 800, h: 600 };

  it('places the popup to the right of the cell, anchored at the cell top (LTR)', () => {
    const pos = computeCommentPopupPosition({
      cell,
      popup,
      viewport,
      rtl: false,
    });
    // Right of the cell: left = cell.x + cell.w + GAP.
    expect(pos.left).toBe(cell.x + cell.w + GAP);
    // Anchored at the cell's top edge.
    expect(pos.top).toBe(cell.y);
  });

  it('flips to the left of the cell when the right side overflows (LTR)', () => {
    // A cell near the right edge: right placement (724 + 8 + 280) overflows 800.
    const rightCell = { x: 724, y: 60, w: 64, h: 20 };
    const pos = computeCommentPopupPosition({
      cell: rightCell,
      popup,
      viewport,
      rtl: false,
    });
    // Flipped to the left: left = cell.x - GAP - popup.w.
    expect(pos.left).toBe(rightCell.x - GAP - popup.w);
    expect(pos.left).toBeGreaterThanOrEqual(0);
  });

  it('clamps the top so the popup stays inside the bottom edge', () => {
    // A cell near the bottom: anchoring at cell.y (560) would push the
    // 90px-tall popup to 650, past the 600px viewport.
    const bottomCell = { x: 100, y: 560, w: 64, h: 20 };
    const pos = computeCommentPopupPosition({
      cell: bottomCell,
      popup,
      viewport,
      rtl: false,
    });
    expect(pos.top).toBe(viewport.h - popup.h);
    expect(pos.top + popup.h).toBeLessThanOrEqual(viewport.h);
  });

  it('clamps the top so the popup never goes above the viewport', () => {
    const topCell = { x: 100, y: 2, w: 64, h: 20 };
    const tallPopup = { w: 280, h: 90 };
    const pos = computeCommentPopupPosition({
      cell: topCell,
      popup: tallPopup,
      viewport,
      rtl: false,
    });
    expect(pos.top).toBeGreaterThanOrEqual(0);
  });

  it('opens toward the cell left by default for RTL (cell rect already mirrored)', () => {
    // For RTL the cell rect is mirrored, so the sheet body lies to the cell's
    // LEFT. The popup must open left so it grows into the body, not off-screen
    // past the right edge. Use a cell with room on the left (x large enough).
    const rtlCell = { x: 500, y: 60, w: 64, h: 20 };
    const pos = computeCommentPopupPosition({
      cell: rtlCell,
      popup,
      viewport,
      rtl: true,
    });
    expect(pos.left).toBe(rtlCell.x - GAP - popup.w);
    expect(pos.top).toBe(rtlCell.y);
  });

  it('flips RTL to the right of the cell when the left side overflows', () => {
    // A cell near the left edge: left placement (100 - 8 - 280 = -188) overflows.
    const leftCell = { x: 100, y: 60, w: 64, h: 20 };
    const pos = computeCommentPopupPosition({
      cell: leftCell,
      popup,
      viewport,
      rtl: true,
    });
    expect(pos.left).toBe(leftCell.x + leftCell.w + GAP);
    expect(pos.left + popup.w).toBeLessThanOrEqual(viewport.w);
  });

  it('clamps left into the viewport when neither side fits (very wide popup)', () => {
    const widePopup = { w: 780, h: 90 };
    const pos = computeCommentPopupPosition({
      cell,
      popup: widePopup,
      viewport,
      rtl: false,
    });
    expect(pos.left).toBeGreaterThanOrEqual(0);
    expect(pos.left + widePopup.w).toBeLessThanOrEqual(viewport.w);
  });
});
