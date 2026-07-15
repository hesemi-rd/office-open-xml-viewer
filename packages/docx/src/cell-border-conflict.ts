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
 * Rules (ECMA-376 §17.4.66 with Word deviations in [MS-OI29500] 2.1.169):
 *   0. `none` loses to the opposing border; `nil` suppresses the shared edge.
 *   1. A CELL border always beats a TABLE(-level or table-style) border.
 *   2. Weight = border width in eighth-points × Word border number; dotted and
 *      dashed have weight 1 regardless of width.
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

/** [MS-OI29500] 2.1.169 — Word's border number. Dotted and dashed are
 * handled separately because Word fixes their complete weight at 1. */
const BORDER_NUMBER: Record<string, number> = {
  single: 1,
  thick: 2,
  double: 3,
  dotDash: 8,
  dotDotDash: 9,
  triple: 10,
  thinThickSmallGap: 11,
  thickThinSmallGap: 12,
  thinThickThinSmallGap: 13,
  thinThickMediumGap: 14,
  thickThinMediumGap: 15,
  thinThickThinMediumGap: 16,
  thinThickLargeGap: 17,
  thickThinLargeGap: 18,
  thinThickThinLargeGap: 19,
  wave: 20,
  doubleWave: 21,
  dashSmallGap: 22,
  dashDotStroked: 23,
  threeDEmboss: 24,
  threeDEngrave: 25,
  outset: 26,
  inset: 27,
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
function borderWeight(spec: BorderSpec): number {
  if (spec.style === 'dotted' || spec.style === 'dashed') return 1;
  return Math.max(0, spec.width) * 8 * borderNumber(spec.style);
}
function precedenceIndex(style: string): number {
  const i = PRECEDENCE.indexOf(style);
  return i === -1 ? PRECEDENCE.length : i; // unknown ⇒ lowest priority
}

function isNil(spec: BorderSpec): boolean {
  return spec.style === 'nil';
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
  // Word differs from the base ECMA rule here: nil suppresses the shared edge,
  // while none merely contributes no competing border ([MS-OI29500] 2.1.169).
  if ((a && isNil(a.spec)) || (b && isNil(b.spec))) return null;
  const av = a && a.spec.style !== 'none' ? a : null;
  const bv = b && b.spec.style !== 'none' ? b : null;
  if (!av && !bv) return null;
  if (!av) return bv;
  if (!bv) return av;

  // Rule #1 — a cell border beats a table border.
  if (av.source === 'cell' && bv.source === 'table') return av;
  if (bv.source === 'cell' && av.source === 'table') return bv;

  // Rule #2 — heavier weight wins.
  const wa = borderWeight(av.spec);
  const wb = borderWeight(bv.spec);
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
