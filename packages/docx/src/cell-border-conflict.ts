import type { BorderSpec, CellBorders, TableBorders } from './types';

/**
 * ECMA-376 §17.4.66 (tcBorders) — adjacent table cell border conflict resolution.
 *
 * When cell spacing is zero, two cells that share an interior gridline each
 * contribute a border for that edge. Word displays exactly ONE of them, chosen
 * by the following rules (applied in order); this module is the pure kernel that
 * decides the winner. The renderer ({@link drawTableRows}) supplies the two
 * candidates for each shared edge and draws only the returned winner, so a
 * gridline is drawn once with the correct spec (no more "last cell painted wins").
 *
 * Rules (§17.4.66):
 *   0. If either border is `nil`/`none` (no border), the OTHER is displayed. If
 *      both are nil/none ⇒ nothing (`null`).
 *   1. A CELL border always beats a TABLE(-level or table-style) border.
 *   2. Weight = (# of lines in the border) × (the style's "border number"); the
 *      larger weight wins.
 *   3. Equal weight ⇒ the style higher on the precedence list wins.
 *   4. Identical style ⇒ the darker colour wins, by three successive brightness
 *      formulas: R+B+2G, then B+2G, then G (smaller value wins each).
 *   5. Still identical ⇒ the border FIRST in reading order (`a`) is displayed.
 */

/** A conflict candidate: the resolved {@link BorderSpec} for one cell's edge plus
 *  whether it came from the cell's own formatting or from a table/table-style
 *  border (rule #1). `null` ⇒ this side contributes no candidate at all. */
export interface BorderCandidate {
  spec: BorderSpec;
  source: 'cell' | 'table';
}

/** Structural location of a cell in the table grid. The same edge cascade is
 * consumed by both border paint and row-footprint measurement. */
export interface CellEdgeFlags {
  topRow: boolean;
  bottomRow: boolean;
  leftCol: boolean;
  rightCol: boolean;
}

/** Cell/table cascade result before an adjacent-cell conflict is resolved. */
export interface ResolvedCellEdges {
  top: BorderCandidate | null;
  bottom: BorderCandidate | null;
  left: BorderCandidate | null;
  right: BorderCandidate | null;
}

/** ECMA-376 §17.4.38/§17.4.39/§17.4.66 — resolve one cell's own, inside,
 * and outer table border cascade. `mirror` maps logical left/right to physical
 * sides for `bidiVisual`; horizontal edges are unchanged. */
export function resolveCellEdges(
  cell: CellBorders,
  table: TableBorders,
  edges: CellEdgeFlags,
  mirror: boolean,
): ResolvedCellEdges {
  const horizontal = (
    own: BorderSpec | null,
    outer: boolean,
    tableOuter: BorderSpec | null,
  ): BorderCandidate | null => {
    if (own) return { spec: own, source: 'cell' };
    const inherited = outer ? tableOuter : (cell.insideH ?? table.insideH);
    return inherited ? { spec: inherited, source: 'table' } : null;
  };
  const vertical = (
    own: BorderSpec | null,
    outer: boolean,
    tableOuter: BorderSpec | null,
  ): BorderCandidate | null => {
    if (own) return { spec: own, source: 'cell' };
    const inherited = outer ? tableOuter : (cell.insideV ?? table.insideV);
    return inherited ? { spec: inherited, source: 'table' } : null;
  };

  const top = horizontal(cell.top, edges.topRow, table.top);
  const bottom = horizontal(cell.bottom, edges.bottomRow, table.bottom);
  const left = mirror
    ? vertical(cell.right, edges.rightCol, table.right)
    : vertical(cell.left, edges.leftCol, table.left);
  const right = mirror
    ? vertical(cell.left, edges.leftCol, table.left)
    : vertical(cell.right, edges.rightCol, table.right);
  return { top, bottom, left, right };
}

/** ECMA-376 §17.4.66 — the "border number" rank of each ST_Border style. The
 *  larger the rank, the heavier the style (before the line-count multiplier).
 *  Unknown / art styles are treated as rank 0 (they never out-weigh a real line
 *  style; art borders are unsupported anyway). */
const BORDER_NUMBER: Record<string, number> = {
  single: 1,
  thick: 2,
  double: 3,
  dotted: 4,
  dashed: 5,
  dotDash: 6,
  dotDotDash: 7,
  triple: 8,
  thinThickSmallGap: 9,
  thickThinSmallGap: 10,
  thinThickThinSmallGap: 11,
  thinThickMediumGap: 12,
  thickThinMediumGap: 13,
  thinThickThinMediumGap: 14,
  thinThickLargeGap: 15,
  thickThinLargeGap: 16,
  thinThickThinLargeGap: 17,
  wave: 18,
  doubleWave: 19,
  dashSmallGap: 20,
  dashDotStroked: 21,
  threeDEmboss: 22,
  threeDEngrave: 23,
  outset: 24,
  inset: 25,
};

/** ECMA-376 §17.18.2 — the number of parallel lines each ST_Border style draws,
 *  the first factor of the §17.4.66 weight. A single rule is 1; a `double` is 2;
 *  a `triple` and the "thinThickThin" families are 3; the two-band "thinThick" /
 *  "thickThin" families and `doubleWave` are 2. All dash / dot / wave / 3D /
 *  outset / inset single-stroke styles are 1. Unknown styles default to 1. */
const BORDER_LINES: Record<string, number> = {
  double: 2,
  triple: 3,
  thinThickSmallGap: 2,
  thickThinSmallGap: 2,
  thinThickThinSmallGap: 3,
  thinThickMediumGap: 2,
  thickThinMediumGap: 2,
  thinThickThinMediumGap: 3,
  thinThickLargeGap: 2,
  thickThinLargeGap: 2,
  thinThickThinLargeGap: 3,
  doubleWave: 2,
};

/** §17.4.66 rule #3 precedence list — index 0 is the highest priority. Identical
 *  to the BORDER_NUMBER ordering (single first … inset last); a smaller index
 *  wins a weight tie. */
const PRECEDENCE: string[] = [
  'single', 'thick', 'double', 'dotted', 'dashed', 'dotDash', 'dotDotDash', 'triple',
  'thinThickSmallGap', 'thickThinSmallGap', 'thinThickThinSmallGap', 'thinThickMediumGap',
  'thickThinMediumGap', 'thinThickThinMediumGap', 'thinThickLargeGap', 'thickThinLargeGap',
  'thinThickThinLargeGap', 'wave', 'doubleWave', 'dashSmallGap', 'dashDotStroked',
  'threeDEmboss', 'threeDEngrave', 'outset', 'inset',
];

function borderNumber(style: string): number {
  return BORDER_NUMBER[style] ?? 0;
}
function borderLines(style: string): number {
  return BORDER_LINES[style] ?? 1;
}
function borderWeight(style: string): number {
  return borderLines(style) * borderNumber(style);
}
function precedenceIndex(style: string): number {
  const i = PRECEDENCE.indexOf(style);
  return i === -1 ? PRECEDENCE.length : i; // unknown ⇒ lowest priority
}

/** True for a border that draws no ink (`nil`/`none`), which §17.4.66 rule #0
 *  treats as "no border". */
function isNil(spec: BorderSpec): boolean {
  return spec.style === 'nil' || spec.style === 'none';
}

/** Parse a 6-hex colour to (r,g,b). `null`/auto ⇒ black (0,0,0): §17.4.66's
 *  darkness comparison uses the RENDERED colour, and auto resolves to black. A
 *  malformed value also falls back to black. */
function rgb(color: string | null): { r: number; g: number; b: number } {
  if (!color) return { r: 0, g: 0, b: 0 };
  const hex = color.replace(/^#/, '');
  if (hex.length !== 6 || /[^0-9a-fA-F]/.test(hex)) return { r: 0, g: 0, b: 0 };
  return {
    r: parseInt(hex.slice(0, 2), 16),
    g: parseInt(hex.slice(2, 4), 16),
    b: parseInt(hex.slice(4, 6), 16),
  };
}

/** §17.4.66 rules #4a–#4c — compare two colours by brightness, DARKER wins.
 *  Returns <0 when `a` is darker (wins), >0 when `b` is darker, 0 when identical
 *  under all three formulas. */
function compareBrightness(a: string | null, b: string | null): number {
  const ca = rgb(a);
  const cb = rgb(b);
  const f1 = (c: { r: number; g: number; b: number }) => c.r + c.b + 2 * c.g;
  const f2 = (c: { r: number; g: number; b: number }) => c.b + 2 * c.g;
  const f3 = (c: { r: number; g: number; b: number }) => c.g;
  for (const f of [f1, f2, f3]) {
    const d = f(ca) - f(cb);
    if (d !== 0) return d;
  }
  return 0;
}

/**
 * Resolve the winning border for a shared cell edge per ECMA-376 §17.4.66. `a` is
 * the border FIRST in reading order (it wins a total tie, rule #5). Either side
 * may be `null` (that cell contributes no border to this edge). Returns the
 * winning candidate, or `null` when neither side paints (both absent or nil/none).
 */
export function resolveBorderConflict(
  a: BorderCandidate | null,
  b: BorderCandidate | null,
): BorderCandidate | null {
  // Rule #0 — nil/none (or absent) contributes nothing.
  const av = a && !isNil(a.spec) ? a : null;
  const bv = b && !isNil(b.spec) ? b : null;
  if (!av && !bv) return null;
  if (!av) return bv;
  if (!bv) return av;

  // Rule #1 — a cell border beats a table border.
  if (av.source === 'cell' && bv.source === 'table') return av;
  if (bv.source === 'cell' && av.source === 'table') return bv;

  // Rule #2 — heavier weight wins.
  const wa = borderWeight(av.spec.style);
  const wb = borderWeight(bv.spec.style);
  if (wa !== wb) return wa > wb ? av : bv;

  // Rule #3 — equal weight ⇒ higher on the precedence list (smaller index).
  const pa = precedenceIndex(av.spec.style);
  const pb = precedenceIndex(bv.spec.style);
  if (pa !== pb) return pa < pb ? av : bv;

  // Rule #4 — identical style ⇒ darker colour wins.
  const cmp = compareBrightness(av.spec.color, bv.spec.color);
  if (cmp !== 0) return cmp < 0 ? av : bv;

  // Rule #5 — fully identical ⇒ the first in reading order.
  return av;
}
