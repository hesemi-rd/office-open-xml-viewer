// Preset pattern bitmaps for OOXML pattFill (ECMA-376 §20.1.10.59).
// Each entry is an 8-row, 8-column binary bitmap. A '1' bit selects the
// foreground colour, a '0' bit the background. Rows are ordered top→bottom;
// within a row the most-significant bit is the leftmost pixel.
//
// Coverage prioritises the well-known geometric variants whose bitmap is
// uniquely determined by the spec or by long-standing Office implementations
// (POI, LibreOffice). Less-common decorative patterns (weave, plaid, sphere,
// shingle, divot, …) are intentionally omitted — their per-pixel layout is
// implementation-specific in practice, and falling back to the foreground
// colour is a safer default than guessing.

import { createAuxCanvas } from '../canvas/aux-canvas.js';

const PATTERN_BITMAPS: Record<string, number[]> = {
  // ── Percentage shading ────────────────────────────────────────────────
  // Sparse-to-dense dot patterns. Bit positions follow the canonical
  // 8x8 templates used by Office/POI for the same preset names.
  pct5:  [0b00000000, 0b00010000, 0b00000000, 0b00000000, 0b00000000, 0b00000001, 0b00000000, 0b00000000],
  pct10: [0b10001000, 0b00000000, 0b00100010, 0b00000000, 0b10001000, 0b00000000, 0b00100010, 0b00000000],
  pct20: [0b10001000, 0b00100010, 0b10001000, 0b00100010, 0b10001000, 0b00100010, 0b10001000, 0b00100010],
  pct25: [0b10001000, 0b01010101, 0b00100010, 0b01010101, 0b10001000, 0b01010101, 0b00100010, 0b01010101],
  pct30: [0b10101010, 0b01010101, 0b10101010, 0b01010101, 0b10101010, 0b01010101, 0b10101010, 0b01010101],
  pct40: [0b10101010, 0b01110111, 0b10101010, 0b11011101, 0b10101010, 0b01110111, 0b10101010, 0b11011101],
  pct50: [0b10101010, 0b01010101, 0b10101010, 0b01010101, 0b10101010, 0b01010101, 0b10101010, 0b01010101],
  pct60: [0b11011101, 0b01010101, 0b01110111, 0b01010101, 0b11011101, 0b01010101, 0b01110111, 0b01010101],
  pct70: [0b11101110, 0b01010101, 0b10111011, 0b01010101, 0b11101110, 0b01010101, 0b10111011, 0b01010101],
  pct75: [0b11101110, 0b10101010, 0b10111011, 0b10101010, 0b11101110, 0b10101010, 0b10111011, 0b10101010],
  pct80: [0b11111110, 0b11101111, 0b11111011, 0b10111111, 0b11111110, 0b11101111, 0b11111011, 0b10111111],
  pct90: [0b11111111, 0b11101111, 0b11111111, 0b11111011, 0b11111111, 0b11101111, 0b11111111, 0b11111011],

  // ── Horizontal / vertical lines ───────────────────────────────────────
  horz:    [0b11111111, 0b00000000, 0b00000000, 0b00000000, 0b11111111, 0b00000000, 0b00000000, 0b00000000],
  vert:    [0b10001000, 0b10001000, 0b10001000, 0b10001000, 0b10001000, 0b10001000, 0b10001000, 0b10001000],
  ltHorz:  [0b00000000, 0b11111111, 0b00000000, 0b00000000, 0b00000000, 0b00000000, 0b00000000, 0b00000000],
  ltVert:  [0b00100000, 0b00100000, 0b00100000, 0b00100000, 0b00100000, 0b00100000, 0b00100000, 0b00100000],
  dkHorz:  [0b11111111, 0b11111111, 0b00000000, 0b00000000, 0b11111111, 0b11111111, 0b00000000, 0b00000000],
  dkVert:  [0b11001100, 0b11001100, 0b11001100, 0b11001100, 0b11001100, 0b11001100, 0b11001100, 0b11001100],
  narHorz: [0b11111111, 0b00000000, 0b11111111, 0b00000000, 0b11111111, 0b00000000, 0b11111111, 0b00000000],
  narVert: [0b10101010, 0b10101010, 0b10101010, 0b10101010, 0b10101010, 0b10101010, 0b10101010, 0b10101010],

  // ── Cross / grid ──────────────────────────────────────────────────────
  cross:   [0b11111111, 0b10001000, 0b10001000, 0b10001000, 0b11111111, 0b10001000, 0b10001000, 0b10001000],
  lgGrid:  [0b11111111, 0b10000000, 0b10000000, 0b10000000, 0b10000000, 0b10000000, 0b10000000, 0b10000000],
  smGrid:  [0b11111111, 0b10001000, 0b10001000, 0b10001000, 0b11111111, 0b10001000, 0b10001000, 0b10001000],
  dotGrid: [0b10001000, 0b00000000, 0b00000000, 0b00000000, 0b10001000, 0b00000000, 0b00000000, 0b00000000],

  // ── Diagonals ─────────────────────────────────────────────────────────
  // dnDiag: top-left → bottom-right stripe. upDiag: bottom-left → top-right.
  dnDiag:    [0b10000000, 0b01000000, 0b00100000, 0b00010000, 0b00001000, 0b00000100, 0b00000010, 0b00000001],
  upDiag:    [0b00000001, 0b00000010, 0b00000100, 0b00001000, 0b00010000, 0b00100000, 0b01000000, 0b10000000],
  ltDnDiag:  [0b10001000, 0b01000100, 0b00100010, 0b00010001, 0b10001000, 0b01000100, 0b00100010, 0b00010001],
  ltUpDiag:  [0b00010001, 0b00100010, 0b01000100, 0b10001000, 0b00010001, 0b00100010, 0b01000100, 0b10001000],
  dkDnDiag:  [0b11000011, 0b10000001, 0b00000000, 0b10000001, 0b11000011, 0b10000001, 0b00000000, 0b10000001],
  dkUpDiag:  [0b11000011, 0b10000001, 0b00000000, 0b10000001, 0b11000011, 0b10000001, 0b00000000, 0b10000001],
  wdDnDiag:  [0b10000000, 0b01000000, 0b00100000, 0b00010000, 0b00001000, 0b00000100, 0b00000010, 0b10000001],
  wdUpDiag:  [0b00000001, 0b00000010, 0b00000100, 0b00001000, 0b00010000, 0b00100000, 0b01000000, 0b10000001],
  diagCross: [0b10000001, 0b01000010, 0b00100100, 0b00011000, 0b00011000, 0b00100100, 0b01000010, 0b10000001],

  // ── Brick / checker ───────────────────────────────────────────────────
  horzBrick: [0b11111111, 0b00010000, 0b00010000, 0b00010000, 0b11111111, 0b00000001, 0b00000001, 0b00000001],
  diagBrick: [0b10000001, 0b01000010, 0b00100100, 0b00011000, 0b00100100, 0b01000010, 0b10000001, 0b00000000],
  lgCheck:   [0b11110000, 0b11110000, 0b11110000, 0b11110000, 0b00001111, 0b00001111, 0b00001111, 0b00001111],
  smCheck:   [0b11001100, 0b11001100, 0b00110011, 0b00110011, 0b11001100, 0b11001100, 0b00110011, 0b00110011],
  trellis:   [0b10100101, 0b01011010, 0b10100101, 0b01011010, 0b10100101, 0b01011010, 0b10100101, 0b01011010],
};

/**
 * Render an 8x8 pattern bitmap to a tile that `ctx.createPattern(_, 'repeat')`
 * will tile across the shape. Uses OffscreenCanvas where available and falls
 * back to a regular HTMLCanvasElement (test envs / older browsers).
 *
 * Returns null when the preset name is unknown — callers should fall back
 * to the foreground colour, never to an arbitrary substitute pattern.
 */
export function buildPatternBitmap(
  preset: string,
  fg: string,
  bg: string,
): HTMLCanvasElement | OffscreenCanvas | null {
  const rows = PATTERN_BITMAPS[preset];
  if (!rows) return null;

  // 8×8 positive-integer tile, so createAuxCanvas's ceil/≥1 clamp is a no-op.
  const tile = createAuxCanvas(8, 8);
  if (!tile) return null;
  const tctx = tile.getContext('2d') as CanvasRenderingContext2D | null;
  if (!tctx) return null;

  tctx.fillStyle = hexToCss(bg);
  tctx.fillRect(0, 0, 8, 8);
  tctx.fillStyle = hexToCss(fg);
  for (let y = 0; y < 8; y++) {
    const row = rows[y];
    for (let x = 0; x < 8; x++) {
      if (row & (1 << (7 - x))) tctx.fillRect(x, y, 1, 1);
    }
  }
  return tile;
}

function hexToCss(hex: string): string {
  const r = parseInt(hex.slice(0, 2), 16);
  const g = parseInt(hex.slice(2, 4), 16);
  const b = parseInt(hex.slice(4, 6), 16);
  const a = hex.length >= 8 ? parseInt(hex.slice(6, 8), 16) / 255 : 1;
  return `rgba(${r},${g},${b},${a})`;
}
