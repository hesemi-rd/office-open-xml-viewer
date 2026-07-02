/**
 * Parity harness: legacy hand-written `buildShapePath` switch (preset.ts) vs
 * the spec-driven preset-geometry engine (preset-geometry/, driven by
 * presets.json from presetShapeDefinitions.xml).
 *
 * ── Why this exists ──────────────────────────────────────────────────────────
 * core carries two preset-shape engines. The goal (Phase 4 A2) is to fold the
 * legacy switch into the spec engine WITHOUT changing a single rendered pixel.
 * A legacy `case` may only be removed when the two engines emit the same
 * geometry; this harness decides that numerically, and the always-on test at
 * the bottom keeps every migrated preset pinned to parity forever.
 *
 * ── Call-surface inventory of the legacy `buildShapePath` ────────────────────
 * (as of the audit; body rendering already prefers the spec engine everywhere)
 *
 *  1. packages/pptx/src/renderer.ts  `paintShapeBody`
 *     - silhouette repaint for effect masks (softEdge §20.1.8.53 mask,
 *       innerShdw §20.1.8.40 mask): reached for ANY preset, including ones the
 *       spec engine handles (`usePresetEngine && silhouette`).
 *     - full fallback when `!hasPreset(geom)` (e.g. `rect`, which
 *       presetShapeDefinitions.xml does not define — its geometry is implicit).
 *  2. packages/docx/src/renderer.ts  `renderAnchorShape`
 *     - fallback when `presetGeometry` is set but `!hasPreset(geom)`
 *       (in practice only `rect`: the docx parser emits raw <a:prstGeom prst>
 *       names, and its VML path always emits "rect").
 *  3. packages/core/src/index.ts — public re-export.
 *  xlsx does not use the legacy engine at all (renderPresetShape → rect
 *  fallback), and pptx picture clips already use buildPresetGeometryPath.
 *
 *  The switch also carries labels that are NOT ECMA-376 ST_ShapeType names
 *  (`oval`, `star`, `document`, …) plus camelCase labels that can never match
 *  the lower-cased input (`pieWedge`, `snipRoundRect`, `irregularSeal1/2`).
 *  Aliases are audited against their canonical spec target below.
 *
 * ── Comparison method ────────────────────────────────────────────────────────
 * Both engines emit Canvas path commands. A PathRecorder replays them per the
 * WHATWG canvas path semantics (implicit line-to before arcs, closePath
 * starting a new subpath at the closed subpath's origin, rect/roundRect
 * trailing pen position, arc sweep clamping) and flattens every segment into
 * dense point samples (≤0.35 px spacing). Two shapes match when:
 *   - subpath counts are equal, paired subpaths agree on the closed flag,
 *   - paired subpaths agree on winding orientation when the shape has ≥2
 *     subpaths and is filled nonzero (donut/smileyFace/frame are exempt: every
 *     live caller fills those with 'evenodd', where winding is irrelevant),
 *   - the symmetric max point-to-polyline deviation is ≤ 5e-3 px on a
 *     100–200 px box (well below any rasterisation threshold; identical-math
 *     cases sit at ~1e-12, absorbing only float noise + sampling error).
 * The flattening deliberately absorbs representation differences (ellipse()
 * vs chained arcs vs Béziers of the same curve); the sanity tests at the
 * bottom prove the normalisation neither hides real differences nor invents
 * spurious ones.
 *
 * Run with PRESET_PARITY_REPORT=<path> to write the full audit table there:
 *   PRESET_PARITY_REPORT=/tmp/parity.md npx vitest run packages/core/src/shape/preset-parity.test.ts
 */

import { writeFileSync } from 'node:fs';
import { describe, it, expect } from 'vitest';
import { buildShapePath, SPEC_MIGRATED_PRESETS } from './preset';
import { buildPresetGeometryPath } from './preset-geometry';
import presetsJson from './preset-geometry/presets.json';

// ── Path recording / flattening ──────────────────────────────────────────────

interface Pt { x: number; y: number }
interface RawSubpath { pts: Pt[]; closed: boolean }

/** Target sample spacing along any segment, px. */
const CHORD = 0.35;
/** Max angular step when flattening an arc, radians. */
const ARC_STEP = 0.02;
/** Geometric match tolerance, px (see header). */
const TOL = 5e-3;

class PathRecorder {
  private done: RawSubpath[] = [];
  private cur: Pt[] | null = null;
  private curStart: Pt | null = null;

  beginPath(): void {
    this.done = [];
    this.cur = null;
    this.curStart = null;
  }

  /** All subpaths, including the still-open trailing one. */
  subpaths(): RawSubpath[] {
    const all = this.done.slice();
    if (this.cur && this.cur.length > 0) all.push({ pts: this.cur, closed: false });
    return all;
  }

  private open(p: Pt): void {
    this.finish(false);
    this.cur = [{ x: p.x, y: p.y }];
    this.curStart = { x: p.x, y: p.y };
  }

  private finish(closed: boolean): void {
    if (this.cur && this.cur.length > 0) this.done.push({ pts: this.cur, closed });
    this.cur = null;
    this.curStart = null;
  }

  private pen(): Pt {
    const c = this.cur!;
    return c[c.length - 1];
  }

  /** Append `to`, interpolating so adjacent samples are ≤ CHORD apart. */
  private sampleLineTo(to: Pt): void {
    const from = this.pen();
    const d = Math.hypot(to.x - from.x, to.y - from.y);
    const steps = Math.max(1, Math.ceil(d / CHORD));
    for (let i = 1; i <= steps; i++) {
      const t = i / steps;
      this.cur!.push({ x: from.x + (to.x - from.x) * t, y: from.y + (to.y - from.y) * t });
    }
  }

  moveTo(x: number, y: number): void {
    this.open({ x, y });
  }

  lineTo(x: number, y: number): void {
    if (!this.cur) { this.open({ x, y }); return; } // canvas: ensure subpath for (x, y)
    this.sampleLineTo({ x, y });
  }

  closePath(): void {
    if (!this.cur || !this.curStart) return;
    const start = { x: this.curStart.x, y: this.curStart.y };
    this.finish(true);
    // canvas: closePath adds a new subpath whose first point is the closed
    // subpath's start — subsequent segments continue from there.
    this.open(start);
  }

  bezierCurveTo(c1x: number, c1y: number, c2x: number, c2y: number, x: number, y: number): void {
    if (!this.cur) this.open({ x: c1x, y: c1y }); // canvas: ensure subpath for cp1
    const p0 = this.pen();
    const len =
      Math.hypot(c1x - p0.x, c1y - p0.y) +
      Math.hypot(c2x - c1x, c2y - c1y) +
      Math.hypot(x - c2x, y - c2y);
    const steps = Math.min(4096, Math.max(8, Math.ceil(len / CHORD)));
    const x0 = p0.x, y0 = p0.y;
    for (let i = 1; i <= steps; i++) {
      const t = i / steps, u = 1 - t;
      this.cur!.push({
        x: u * u * u * x0 + 3 * u * u * t * c1x + 3 * u * t * t * c2x + t * t * t * x,
        y: u * u * u * y0 + 3 * u * u * t * c1y + 3 * u * t * t * c2y + t * t * t * y,
      });
    }
  }

  quadraticCurveTo(cx: number, cy: number, x: number, y: number): void {
    if (!this.cur) this.open({ x: cx, y: cy });
    const p0 = this.pen();
    const len = Math.hypot(cx - p0.x, cy - p0.y) + Math.hypot(x - cx, y - cy);
    const steps = Math.min(4096, Math.max(8, Math.ceil(len / CHORD)));
    const x0 = p0.x, y0 = p0.y;
    for (let i = 1; i <= steps; i++) {
      const t = i / steps, u = 1 - t;
      this.cur!.push({
        x: u * u * x0 + 2 * u * t * cx + t * t * x,
        y: u * u * y0 + 2 * u * t * cy + t * t * y,
      });
    }
  }

  ellipse(
    cx: number, cy: number, rx: number, ry: number,
    rot: number, a0: number, a1: number, ccw = false,
  ): void {
    const TAU = Math.PI * 2;
    // WHATWG sweep clamping: ≥ full turn collapses to exactly one revolution,
    // otherwise reduce mod 2π in the traversal direction.
    let sweep: number;
    if (!ccw) {
      if (a1 - a0 >= TAU) sweep = TAU;
      else {
        sweep = (a1 - a0) % TAU;
        if (sweep < 0) sweep += TAU;
      }
    } else {
      if (a0 - a1 >= TAU) sweep = -TAU;
      else {
        sweep = (a1 - a0) % TAU;
        if (sweep > 0) sweep -= TAU;
      }
    }
    const cosR = Math.cos(rot), sinR = Math.sin(rot);
    const at = (th: number): Pt => ({
      x: cx + rx * Math.cos(th) * cosR - ry * Math.sin(th) * sinR,
      y: cy + rx * Math.cos(th) * sinR + ry * Math.sin(th) * cosR,
    });
    const start = at(a0);
    if (this.cur) this.sampleLineTo(start); // canvas: implicit line to arc start
    else this.open(start);
    const rMax = Math.max(Math.abs(rx), Math.abs(ry));
    const steps = Math.min(
      8192,
      Math.max(
        8,
        Math.ceil(Math.abs(sweep) / ARC_STEP),
        Math.ceil((Math.abs(sweep) * rMax) / CHORD),
      ),
    );
    for (let i = 1; i <= steps; i++) this.cur!.push(at(a0 + (sweep * i) / steps));
  }

  arc(cx: number, cy: number, r: number, a0: number, a1: number, ccw = false): void {
    this.ellipse(cx, cy, r, r, 0, a0, a1, ccw);
  }

  rect(x: number, y: number, w: number, h: number): void {
    this.moveTo(x, y);
    this.lineTo(x + w, y);
    this.lineTo(x + w, y + h);
    this.lineTo(x, y + h);
    this.closePath(); // leaves the pen on a fresh subpath at (x, y) — canvas parity
  }

  roundRect(
    x: number, y: number, w: number, h: number,
    radii: number | Array<number | { x: number; y: number }>,
  ): void {
    // WHATWG: each radius is a number (circular) or DOMPointInit (elliptical).
    const norm = (r: number | { x: number; y: number }): { x: number; y: number } =>
      typeof r === 'number' ? { x: r, y: r } : { x: r.x, y: r.y };
    const list = (Array.isArray(radii) ? radii : [radii]).map(norm);
    // Corner expansion: 1 → all; 2 → [tl+br, tr+bl]; 3 → [tl, tr+bl, br]; 4 → each.
    let tl: { x: number; y: number }, tr: typeof tl, br: typeof tl, bl: typeof tl;
    if (list.length === 1) { tl = tr = br = bl = list[0]; }
    else if (list.length === 2) { tl = br = list[0]; tr = bl = list[1]; }
    else if (list.length === 3) { tl = list[0]; tr = bl = list[1]; br = list[2]; }
    else { [tl, tr, br, bl] = list; }
    // Scale down when corners overlap (WHATWG step 9).
    const scale = Math.min(
      1,
      w / Math.max(1e-12, tl.x + tr.x), w / Math.max(1e-12, bl.x + br.x),
      h / Math.max(1e-12, tl.y + bl.y), h / Math.max(1e-12, tr.y + br.y),
    );
    if (scale < 1) {
      tl = { x: tl.x * scale, y: tl.y * scale };
      tr = { x: tr.x * scale, y: tr.y * scale };
      br = { x: br.x * scale, y: br.y * scale };
      bl = { x: bl.x * scale, y: bl.y * scale };
    }
    const HALF = Math.PI / 2;
    this.moveTo(x + tl.x, y);
    this.lineTo(x + w - tr.x, y);
    if (tr.x > 0 && tr.y > 0) this.ellipse(x + w - tr.x, y + tr.y, tr.x, tr.y, 0, -HALF, 0);
    this.lineTo(x + w, y + h - br.y);
    if (br.x > 0 && br.y > 0) this.ellipse(x + w - br.x, y + h - br.y, br.x, br.y, 0, 0, HALF);
    this.lineTo(x + bl.x, y + h);
    if (bl.x > 0 && bl.y > 0) this.ellipse(x + bl.x, y + h - bl.y, bl.x, bl.y, 0, HALF, Math.PI);
    this.lineTo(x, y + tl.y);
    if (tl.x > 0 && tl.y > 0) this.ellipse(x + tl.x, y + tl.y, tl.x, tl.y, 0, Math.PI, Math.PI + HALF);
    this.closePath();
    this.moveTo(x, y); // canvas: roundRect leaves a fresh one-point subpath at (x, y)
  }
}

/** A recorder duck-typed as the ctx both engines expect. */
function recorder(): PathRecorder & CanvasRenderingContext2D {
  return new PathRecorder() as unknown as PathRecorder & CanvasRenderingContext2D;
}

// ── Normalised shapes + geometric comparison ────────────────────────────────

interface NormSub {
  pts: Pt[];
  closed: boolean;
  /** Shoelace signed area (implicit closure) — orientation carrier. */
  area: number;
  cx: number;
  cy: number;
}

function normalize(rec: PathRecorder): NormSub[] {
  const out: NormSub[] = [];
  for (const sp of rec.subpaths()) {
    // Drop consecutive duplicates.
    const pts: Pt[] = [];
    for (const p of sp.pts) {
      const last = pts[pts.length - 1];
      if (!last || Math.hypot(p.x - last.x, p.y - last.y) > 1e-9) pts.push(p);
    }
    if (sp.closed) {
      // The closing segment is part of the outline; sample it like any other.
      const first = pts[0], last = pts[pts.length - 1];
      if (pts.length > 1 && Math.hypot(first.x - last.x, first.y - last.y) > 1e-9) {
        const d = Math.hypot(first.x - last.x, first.y - last.y);
        const steps = Math.max(1, Math.ceil(d / CHORD));
        for (let i = 1; i < steps; i++) {
          const t = i / steps;
          pts.push({ x: last.x + (first.x - last.x) * t, y: last.y + (first.y - last.y) * t });
        }
      }
    }
    // Degenerate (empty / single point / zero length) subpaths draw nothing.
    let len = 0;
    for (let i = 1; i < pts.length; i++) len += Math.hypot(pts[i].x - pts[i - 1].x, pts[i].y - pts[i - 1].y);
    if (pts.length < 2 || len < 1e-9) continue;
    let area = 0, sx = 0, sy = 0;
    for (let i = 0; i < pts.length; i++) {
      const a = pts[i], b = pts[(i + 1) % pts.length];
      area += a.x * b.y - b.x * a.y;
      sx += a.x; sy += a.y;
    }
    out.push({ pts, closed: sp.closed, area: area / 2, cx: sx / pts.length, cy: sy / pts.length });
  }
  return out;
}

interface Seg { ax: number; ay: number; bx: number; by: number }

function collectSegs(subs: NormSub[]): Seg[] {
  const segs: Seg[] = [];
  for (const s of subs) {
    const n = s.pts.length;
    const wrap = s.closed ? n : n - 1;
    for (let i = 0; i < wrap; i++) {
      const a = s.pts[i], b = s.pts[(i + 1) % n];
      segs.push({ ax: a.x, ay: a.y, bx: b.x, by: b.y });
    }
  }
  return segs;
}

function segDist(px: number, py: number, s: Seg): number {
  const vx = s.bx - s.ax, vy = s.by - s.ay;
  const wx = px - s.ax, wy = py - s.ay;
  const vv = vx * vx + vy * vy;
  const t = vv > 0 ? Math.max(0, Math.min(1, (wx * vx + wy * vy) / vv)) : 0;
  return Math.hypot(px - (s.ax + vx * t), py - (s.ay + vy * t));
}

/** Spatial grid over segments for near-linear nearest-segment queries. */
class SegGrid {
  private cell: number;
  private map = new Map<string, number[]>();
  private minX = Infinity; private minY = Infinity;
  private maxX = -Infinity; private maxY = -Infinity;
  constructor(private segs: Seg[], cell = 4) {
    this.cell = cell;
    segs.forEach((s, idx) => {
      const x0 = Math.floor(Math.min(s.ax, s.bx) / cell);
      const x1 = Math.floor(Math.max(s.ax, s.bx) / cell);
      const y0 = Math.floor(Math.min(s.ay, s.by) / cell);
      const y1 = Math.floor(Math.max(s.ay, s.by) / cell);
      for (let gx = x0; gx <= x1; gx++) {
        for (let gy = y0; gy <= y1; gy++) {
          const k = gx + ',' + gy;
          let arr = this.map.get(k);
          if (!arr) { arr = []; this.map.set(k, arr); }
          arr.push(idx);
        }
      }
      this.minX = Math.min(this.minX, Math.min(s.ax, s.bx));
      this.maxX = Math.max(this.maxX, Math.max(s.ax, s.bx));
      this.minY = Math.min(this.minY, Math.min(s.ay, s.by));
      this.maxY = Math.max(this.maxY, Math.max(s.ay, s.by));
    });
  }

  /** Exact distance from (px, py) to the nearest segment. */
  nearest(px: number, py: number): number {
    const cgx = Math.floor(px / this.cell), cgy = Math.floor(py / this.cell);
    let best = Infinity;
    const maxRing = Math.ceil(
      (Math.max(this.maxX - this.minX, this.maxY - this.minY) + Math.hypot(px, py)) / this.cell,
    ) + 2;
    for (let ring = 0; ring <= maxRing; ring++) {
      // Every segment closer than (ring-1)*cell lives in a visited cell, so
      // once best undercuts that bound the scan is complete.
      if (best <= (ring - 1) * this.cell) break;
      for (let gx = cgx - ring; gx <= cgx + ring; gx++) {
        for (let gy = cgy - ring; gy <= cgy + ring; gy++) {
          if (Math.max(Math.abs(gx - cgx), Math.abs(gy - cgy)) !== ring) continue; // ring shell only
          const arr = this.map.get(gx + ',' + gy);
          if (!arr) continue;
          for (const i of arr) {
            const d = segDist(px, py, this.segs[i]);
            if (d < best) best = d;
          }
        }
      }
    }
    return best;
  }
}

/** Max over a's samples of the distance to b's outline (directed deviation). */
function directedDeviation(a: NormSub[], b: NormSub[]): number {
  const segs = collectSegs(b);
  if (segs.length === 0) return a.length === 0 ? 0 : Infinity;
  const grid = new SegGrid(segs);
  let worst = 0;
  for (const s of a) {
    for (const p of s.pts) {
      const d = grid.nearest(p.x, p.y);
      if (d > worst) worst = d;
    }
  }
  return worst;
}

export interface CompareOutcome {
  ok: boolean;
  /** Symmetric max deviation in px (NaN when structure already differs). */
  maxDev: number;
  reason: string | null;
}

/** Direction from pts[i0] to the first sample ≥ 0.05 px away (unit vector). */
function tangentFrom(pts: Pt[], i0: number, step: number): Pt | null {
  const a = pts[i0];
  for (let i = i0 + step; i >= 0 && i < pts.length; i += step) {
    const dx = pts[i].x - a.x;
    const dy = pts[i].y - a.y;
    const d = Math.hypot(dx, dy);
    if (d >= 0.05) return { x: dx / d, y: dy / d };
  }
  return null;
}

/** Open subpath whose ends coincide and meet with tangent continuity. */
function isSmoothLoop(s: NormSub): boolean {
  if (s.closed) return true;
  const first = s.pts[0];
  const last = s.pts[s.pts.length - 1];
  if (Math.hypot(first.x - last.x, first.y - last.y) > 1e-6) return false;
  const outgoing = tangentFrom(s.pts, 0, +1); // seam → forward
  const incomingRev = tangentFrom(s.pts, s.pts.length - 1, -1); // seam → backward
  if (!outgoing || !incomingRev) return false;
  // Incoming direction at the seam is the reverse of the backward probe.
  const dot = -(outgoing.x * incomingRev.x + outgoing.y * incomingRev.y);
  return dot >= Math.cos((2 * Math.PI) / 180); // within 2°
}

/**
 * Compare two recorded shapes for fill/stroke-equivalent geometry.
 * `orientationExempt`: skip the winding check (only sound when every live
 * caller fills the shape with the 'evenodd' rule).
 */
function compareShapes(
  a: PathRecorder,
  b: PathRecorder,
  orientationExempt = false,
): CompareOutcome {
  const A = normalize(a);
  const B = normalize(b);
  if (A.length !== B.length) {
    return { ok: false, maxDev: NaN, reason: `subpaths ${A.length}≠${B.length}` };
  }
  // Pair subpaths greedily by centroid proximity (emission order may differ).
  const used = new Set<number>();
  const pairs: Array<[NormSub, NormSub]> = [];
  for (const sa of A) {
    let bestI = -1;
    let bestD = Infinity;
    B.forEach((sb, i) => {
      if (used.has(i)) return;
      const d = Math.hypot(sa.cx - sb.cx, sa.cy - sb.cy);
      if (d < bestD) { bestD = d; bestI = i; }
    });
    used.add(bestI);
    pairs.push([sa, B[bestI]]);
  }
  for (const [sa, sb] of pairs) {
    if (sa.closed !== sb.closed) {
      // An OPEN subpath that loops back onto its own start with a smooth
      // tangent renders identically to the closed version: the fill is the
      // same region, and the stroke's two abutting butt caps at a smooth seam
      // cover exactly what a join would. Only a corner seam (tangent break)
      // or a genuinely open outline is a real difference.
      const open = sa.closed ? sb : sa;
      if (!isSmoothLoop(open)) {
        return { ok: false, maxDev: NaN, reason: 'closed-flag differs' };
      }
    }
  }
  if (!orientationExempt && A.length >= 2) {
    for (const [sa, sb] of pairs) {
      // Near-zero areas (open sliver paths) carry no reliable orientation.
      if (Math.abs(sa.area) > 1e-6 && Math.abs(sb.area) > 1e-6 &&
          Math.sign(sa.area) !== Math.sign(sb.area)) {
        return { ok: false, maxDev: NaN, reason: 'winding orientation differs' };
      }
    }
  }
  const dev = Math.max(directedDeviation(A, B), directedDeviation(B, A));
  return {
    ok: dev <= TOL,
    maxDev: dev,
    reason: dev <= TOL ? null : `geometry Δ=${dev.toPrecision(3)}px`,
  };
}

// ── Fill-region equivalence (report metadata only) ──────────────────────────
// Canvas `fill()` implicitly closes every open subpath, so fill semantics are
// evaluated with forced closure. Two shapes are fill-equivalent when a dense
// winding-rule scan classifies every sample point identically, ignoring
// samples that sit on either outline (where rasterisation is ambiguous
// anyway). A DIFF preset that is still fill-equivalent renders identical
// *silhouettes* (the only live legacy call path for engine-known presets) —
// these are the prime convergence candidates.

function fillSegs(subs: NormSub[]): Seg[] {
  const segs: Seg[] = [];
  for (const s of subs) {
    const n = s.pts.length;
    for (let i = 0; i < n; i++) {
      const a = s.pts[i], b = s.pts[(i + 1) % n]; // implicit closure
      segs.push({ ax: a.x, ay: a.y, bx: b.x, by: b.y });
    }
  }
  return segs;
}

function fillEquivalent(
  A: NormSub[],
  B: NormSub[],
  rule: 'nonzero' | 'evenodd',
): boolean {
  const segsA = fillSegs(A);
  const segsB = fillSegs(B);
  const all = segsA.concat(segsB);
  if (all.length === 0) return true;
  let minX = Infinity, minY = Infinity, maxX = -Infinity, maxY = -Infinity;
  for (const s of all) {
    minX = Math.min(minX, s.ax, s.bx); maxX = Math.max(maxX, s.ax, s.bx);
    minY = Math.min(minY, s.ay, s.by); maxY = Math.max(maxY, s.ay, s.by);
  }
  const N = 160;
  // Per scanline: sorted crossings + suffix sums, so classifying a sample is a
  // binary search instead of an O(segs) sweep.
  const rowClassifier = (segs: Seg[], py: number) => {
    const xs: number[] = [];
    const dirs: number[] = [];
    const cross: Array<{ x: number; dir: number }> = [];
    for (const s of segs) {
      // Half-open vertical span so shared vertices count once.
      const down = s.ay <= py && py < s.by;
      const up = s.by <= py && py < s.ay;
      if (!down && !up) continue;
      const t = (py - s.ay) / (s.by - s.ay);
      cross.push({ x: s.ax + (s.bx - s.ax) * t, dir: down ? 1 : -1 });
    }
    cross.sort((a, b) => a.x - b.x);
    for (const c of cross) { xs.push(c.x); dirs.push(c.dir); }
    // suffixWinding[i] = Σ dirs[i..]; suffixCount[i] = crossings at index ≥ i.
    const suffixWinding = new Array<number>(xs.length + 1).fill(0);
    for (let i = xs.length - 1; i >= 0; i--) suffixWinding[i] = suffixWinding[i + 1] + dirs[i];
    return (px: number): boolean => {
      // first index with xs[idx] > px
      let lo = 0, hi = xs.length;
      while (lo < hi) {
        const mid = (lo + hi) >> 1;
        if (xs[mid] > px) hi = mid;
        else lo = mid + 1;
      }
      return rule === 'nonzero' ? suffixWinding[lo] !== 0 : (xs.length - lo) % 2 === 1;
    };
  };
  let gridA: SegGrid | null = null;
  let gridB: SegGrid | null = null;
  for (let iy = 0; iy < N; iy++) {
    const py = minY + ((iy + 0.5) / N) * (maxY - minY);
    const inA = rowClassifier(segsA, py);
    const inB = rowClassifier(segsB, py);
    for (let ix = 0; ix < N; ix++) {
      const px = minX + ((ix + 0.5) / N) * (maxX - minX);
      if (inA(px) === inB(px)) continue;
      // Disagreement — genuine only when clear of both outlines.
      gridA ??= new SegGrid(segsA);
      gridB ??= new SegGrid(segsB);
      if (Math.min(gridA.nearest(px, py), gridB.nearest(px, py)) > 0.05) return false;
    }
  }
  return true;
}

// ── The legacy switch inventory ──────────────────────────────────────────────
// Every label the legacy switch carried LIVE at audit time (step 0 of the
// unification; camelCase labels `pieWedge`, `snipRoundRect`, `irregularSeal1`,
// `irregularSeal2` were unreachable — the switch lower-cases its input — and
// are not audited). `spec` is the presets.json geometry each label must be
// equivalent to before its case can be deleted; for non-ECMA alias labels
// that is the canonical preset whose body the label shared (or duplicated).
// Labels whose spec target is in SPEC_MIGRATED_PRESETS have had their case
// deleted — buildShapePath now routes them through the engine, so their audit
// rows double as delegation/alias-mapping checks.
const LEGACY_SWITCH: ReadonlyArray<{ label: string; spec: string }> = [
  { label: 'ellipse', spec: 'ellipse' },
  { label: 'oval', spec: 'ellipse' }, // alias
  { label: 'rtriangle', spec: 'rttriangle' }, // alias (duplicate body)
  { label: 'rttriangle', spec: 'rttriangle' },
  { label: 'triangle', spec: 'triangle' },
  { label: 'diamond', spec: 'diamond' },
  { label: 'parallelogram', spec: 'parallelogram' },
  { label: 'trapezoid', spec: 'trapezoid' },
  { label: 'roundrect', spec: 'roundrect' },
  { label: 'roundrectangle', spec: 'roundrect' }, // alias
  { label: 'pentagon', spec: 'pentagon' },
  { label: 'hexagon', spec: 'hexagon' },
  { label: 'heptagon', spec: 'heptagon' },
  { label: 'octagon', spec: 'octagon' },
  { label: 'decagon', spec: 'decagon' },
  { label: 'dodecagon', spec: 'dodecagon' },
  { label: 'star4', spec: 'star4' },
  { label: 'star5', spec: 'star5' },
  { label: 'star', spec: 'star5' }, // alias
  { label: 'star6', spec: 'star6' },
  { label: 'star7', spec: 'star7' },
  { label: 'star8', spec: 'star8' },
  { label: 'star10', spec: 'star10' },
  { label: 'star12', spec: 'star12' },
  { label: 'star16', spec: 'star16' },
  { label: 'star24', spec: 'star24' },
  { label: 'star32', spec: 'star32' },
  { label: 'rightarrow', spec: 'rightarrow' },
  { label: 'leftarrow', spec: 'leftarrow' },
  { label: 'uparrow', spec: 'uparrow' },
  { label: 'downarrow', spec: 'downarrow' },
  { label: 'leftrightarrow', spec: 'leftrightarrow' },
  { label: 'updownarrow', spec: 'updownarrow' },
  { label: 'notchedrightarrow', spec: 'notchedrightarrow' },
  { label: 'chevron', spec: 'chevron' },
  { label: 'homeplate', spec: 'homeplate' },
  { label: 'leftbracket', spec: 'leftbracket' },
  { label: 'rightbracket', spec: 'rightbracket' },
  { label: 'leftbrace', spec: 'leftbrace' },
  { label: 'rightbrace', spec: 'rightbrace' },
  { label: 'callout1', spec: 'callout1' },
  { label: 'bordercallout1', spec: 'bordercallout1' },
  { label: 'accentcallout1', spec: 'accentcallout1' },
  { label: 'accentbordercallout1', spec: 'accentbordercallout1' },
  { label: 'wedgerectcallout', spec: 'wedgerectcallout' },
  { label: 'wedgeellipsecallout', spec: 'wedgeellipsecallout' },
  { label: 'cloudcallout', spec: 'cloudcallout' },
  { label: 'line', spec: 'line' },
  { label: 'straightconnector1', spec: 'straightconnector1' },
  { label: 'bentconnector2', spec: 'bentconnector2' },
  { label: 'bentconnector3', spec: 'bentconnector3' },
  { label: 'bentconnector4', spec: 'bentconnector4' },
  { label: 'bentconnector5', spec: 'bentconnector5' },
  { label: 'curvedconnector2', spec: 'curvedconnector2' },
  { label: 'curvedconnector3', spec: 'curvedconnector3' },
  { label: 'curvedconnector4', spec: 'curvedconnector4' },
  { label: 'curvedconnector5', spec: 'curvedconnector5' },
  { label: 'heart', spec: 'heart' },
  { label: 'donut', spec: 'donut' },
  { label: 'nosmoking', spec: 'nosmoking' },
  { label: 'nosmokingsign', spec: 'nosmoking' }, // alias
  { label: 'pie', spec: 'pie' },
  { label: 'cloud', spec: 'cloud' },
  { label: 'funnel', spec: 'funnel' },
  { label: 'smileyface', spec: 'smileyface' },
  { label: 'document', spec: 'foldedcorner' }, // alias (shares foldedCorner body)
  { label: 'foldedcorner', spec: 'foldedcorner' },
  { label: 'snip1rect', spec: 'snip1rect' },
  { label: 'snip2samerect', spec: 'snip2samerect' },
  { label: 'snip2diagrect', spec: 'snip2diagrect' },
  { label: 'sniproundrect', spec: 'sniproundrect' },
  { label: 'round1rect', spec: 'round1rect' },
  { label: 'round2samerect', spec: 'round2samerect' },
  { label: 'round2diagrect', spec: 'round2diagrect' },
  { label: 'plaque', spec: 'plaque' },
  { label: 'can', spec: 'can' },
  { label: 'cube', spec: 'cube' },
  { label: 'bevel', spec: 'bevel' },
  { label: 'halfframe', spec: 'halfframe' },
  { label: 'corner', spec: 'corner' },
  { label: 'irregularseal1', spec: 'irregularseal1' },
  { label: 'irregularseal2', spec: 'irregularseal2' },
  { label: 'flowchartalternateprocess', spec: 'flowchartalternateprocess' },
  { label: 'flowchartprocess', spec: 'flowchartprocess' },
  { label: 'flowchartdecision', spec: 'flowchartdecision' },
  { label: 'flowchartterminator', spec: 'flowchartterminator' },
  { label: 'flowchartdocument', spec: 'flowchartdocument' },
  { label: 'flowchartpredefinedprocess', spec: 'flowchartpredefinedprocess' },
  { label: 'flowchartsort', spec: 'flowchartsort' },
  { label: 'flowchartmanualinput', spec: 'flowchartmanualinput' },
  { label: 'moon', spec: 'moon' },
  { label: 'arc', spec: 'arc' },
  { label: 'mathequal', spec: 'mathequal' },
  { label: 'mathmultiply', spec: 'mathmultiply' },
  { label: 'mathplus', spec: 'mathplus' },
  { label: 'mathminus', spec: 'mathminus' },
  { label: 'mathdivide', spec: 'mathdivide' },
  { label: 'quadarrow', spec: 'quadarrow' },
  { label: 'quadarrowcallout', spec: 'quadarrowcallout' },
  { label: 'wave', spec: 'wave' },
  { label: 'doublewave', spec: 'doublewave' },
  { label: 'sun', spec: 'sun' },
  { label: 'lightningbolt', spec: 'lightningbolt' },
  { label: 'frame', spec: 'frame' },
  { label: 'bracketpair', spec: 'bracketpair' },
  { label: 'bracepair', spec: 'bracepair' },
  { label: 'chord', spec: 'chord' },
  { label: 'blockarc', spec: 'blockarc' },
  { label: 'teardrop', spec: 'teardrop' },
  { label: 'diagstripe', spec: 'diagstripe' },
  { label: 'wedgeroundrectcallout', spec: 'wedgeroundrectcallout' },
  { label: 'rightarrowcallout', spec: 'rightarrowcallout' },
  { label: 'leftarrowcallout', spec: 'leftarrowcallout' },
  { label: 'uparrowcallout', spec: 'uparrowcallout' },
  { label: 'downarrowcallout', spec: 'downarrowcallout' },
  { label: 'leftrightarrowcallout', spec: 'leftrightarrowcallout' },
  { label: 'leftrightuparrow', spec: 'leftrightuparrow' },
  { label: 'leftuparrow', spec: 'leftuparrow' },
  { label: 'uturnarrow', spec: 'uturnarrow' },
  { label: 'bentarrow', spec: 'bentarrow' },
  { label: 'bentuparrow', spec: 'bentuparrow' },
  { label: 'plus', spec: 'plus' },
  { label: 'mathnotequal', spec: 'mathnotequal' },
  { label: 'flowchartconnector', spec: 'flowchartconnector' },
  { label: 'flowchartdelay', spec: 'flowchartdelay' },
  { label: 'flowchartdisplay', spec: 'flowchartdisplay' },
  { label: 'flowchartinputoutput', spec: 'flowchartinputoutput' },
  { label: 'flowchartpunchedcard', spec: 'flowchartpunchedcard' },
  { label: 'flowchartmerge', spec: 'flowchartmerge' },
  { label: 'flowchartextract', spec: 'flowchartextract' },
  { label: 'flowchartoffpageconnector', spec: 'flowchartoffpageconnector' },
  { label: 'flowchartonlinestorage', spec: 'flowchartonlinestorage' },
  { label: 'flowchartmanuallabel', spec: 'flowchartonlinestorage' }, // alias (shares body)
  { label: 'flowchartpuncheddisk', spec: 'flowchartonlinestorage' }, // alias (shares body)
  { label: 'horizontalscroll', spec: 'horizontalscroll' },
  { label: 'verticalscroll', spec: 'verticalscroll' },
  { label: 'ribbon', spec: 'ribbon' },
  { label: 'ribbon2', spec: 'ribbon2' },
  { label: 'ellipseribbon', spec: 'ellipseribbon' },
  { label: 'ellipseribbon2', spec: 'ellipseribbon2' },
  { label: 'circulararrow', spec: 'circulararrow' },
  { label: 'curvedrightarrow', spec: 'curvedrightarrow' },
  { label: 'curvedleftarrow', spec: 'curvedleftarrow' },
  { label: 'curveduparrow', spec: 'curveduparrow' },
  { label: 'curveddownarrow', spec: 'curveddownarrow' },
  { label: 'stripedrightarrow', spec: 'stripedrightarrow' },
  { label: 'flowchartpreparation', spec: 'flowchartpreparation' },
  { label: 'flowchartcollate', spec: 'flowchartcollate' },
  { label: 'flowchartmagneticdisk', spec: 'flowchartmagneticdisk' },
  { label: 'flowchartinternalstorage', spec: 'flowchartinternalstorage' },
  { label: 'flowchartmagneticdrum', spec: 'flowchartmagneticdrum' },
  { label: 'flowchartsumingjunction', spec: 'flowchartsummingjunction' }, // alias (typo)
  { label: 'flowchartsummingjunction', spec: 'flowchartsummingjunction' },
  { label: 'flowchartmagnetictape', spec: 'flowchartmagnetictape' },
  { label: 'flowchartpunchedtape', spec: 'flowchartpunchedtape' },
  { label: 'flowchartmanualoperation', spec: 'flowchartmanualoperation' },
  { label: 'flowchartmultidocument', spec: 'flowchartmultidocument' },
];

/** Every live caller fills these with 'evenodd' — winding may differ freely. */
const ORIENTATION_EXEMPT = new Set(['donut', 'smileyface', 'frame']);

// Representative boxes: square, wide, tall — offset origin to catch x/y bugs.
const BOXES = [
  { x: 13.25, y: 7.5, w: 100, h: 100 },
  { x: 13.25, y: 7.5, w: 200, h: 80 },
  { x: 13.25, y: 7.5, w: 80, h: 200 },
] as const;

interface PresetDefLite { adj: [string, string][] }
const PRESET_DEFS = presetsJson as unknown as Record<string, PresetDefLite>;

/** Record both engines for one box + adj set. */
function recordBoth(
  label: string,
  spec: string,
  box: (typeof BOXES)[number],
  adj: (number | null)[],
): { legacy: PathRecorder; engine: PathRecorder; engineOk: boolean } {
  const legacy = recorder();
  legacy.beginPath();
  buildShapePath(legacy, label, box.x, box.y, box.w, box.h, adj[0] ?? null, adj[1] ?? null, adj[2] ?? null, adj[3] ?? null);
  const engine = recorder();
  engine.beginPath();
  const engineOk = buildPresetGeometryPath(engine, spec, box.x, box.y, box.w, box.h, adj);
  return { legacy, engine, engineOk };
}

/** Compare a legacy label against its spec target for one box + adj set. */
function compareOnce(
  label: string,
  spec: string,
  box: (typeof BOXES)[number],
  adj: (number | null)[],
): CompareOutcome {
  const { legacy, engine, engineOk } = recordBoth(label, spec, box, adj);
  if (!engineOk) return { ok: false, maxDev: NaN, reason: 'preset missing from presets.json' };
  return compareShapes(legacy, engine, ORIENTATION_EXEMPT.has(spec));
}

export interface AuditRow {
  label: string;
  spec: string;
  /** Default-adjust parity across all three boxes. */
  matchDefault: boolean;
  /**
   * Perturbed-adjust parity. For presets whose spec declares no avLst this
   * probes whether the LEGACY body honours undeclared adjust values the spec
   * engine (and PowerPoint) would ignore. null → no probe ran.
   */
  matchAdj: boolean | null;
  /**
   * For DIFF rows only: identical filled region under the live fill rule
   * despite the coordinate/topology difference — i.e. the silhouette (the
   * only live legacy path for engine-known presets) would not change.
   */
  fillEquiv: boolean | null;
  detail: string;
}

/** Perturbed adjust values: default × 0.6 (a default of 0 becomes 5000).
 *  When the spec declares no adjusts, probe with a mid-range value so a
 *  legacy body that (wrongly) honours undeclared adj values is flagged. */
function perturbedAdj(spec: string): number[] {
  const def = PRESET_DEFS[spec];
  if (!def || def.adj.length === 0) return [30000, 30000, 30000, 30000];
  return def.adj.map(([, fmla]) => {
    const v = Number(fmla.replace(/^val\s+/, '')) || 0;
    return v === 0 ? 5000 : Math.round(v * 0.6);
  });
}

function auditRow(entry: { label: string; spec: string }): AuditRow {
  const details: string[] = [];
  let matchDefault = true;
  for (const box of BOXES) {
    const r = compareOnce(entry.label, entry.spec, box, []);
    if (!r.ok) {
      matchDefault = false;
      details.push(`${box.w}×${box.h}: ${r.reason}`);
    } else if (r.maxDev > TOL / 10) {
      details.push(`${box.w}×${box.h}: dev=${r.maxDev.toExponential(1)}`);
    }
  }
  const adj = perturbedAdj(entry.spec);
  let matchAdj = true;
  for (const box of BOXES) {
    const r = compareOnce(entry.label, entry.spec, box, adj);
    if (!r.ok) {
      matchAdj = false;
      if (matchDefault) details.push(`adj ${box.w}×${box.h}: ${r.reason}`);
      break;
    }
  }
  let fillEquiv: boolean | null = null;
  if (!matchDefault) {
    const rule = ORIENTATION_EXEMPT.has(entry.spec) ? 'evenodd' : 'nonzero';
    fillEquiv = BOXES.every((box) => {
      const { legacy, engine, engineOk } = recordBoth(entry.label, entry.spec, box, []);
      return engineOk && fillEquivalent(normalize(legacy), normalize(engine), rule);
    });
  }
  return { label: entry.label, spec: entry.spec, matchDefault, matchAdj, fillEquiv, detail: details.join('; ') };
}

export function runAudit(): AuditRow[] {
  return LEGACY_SWITCH.map(auditRow);
}

// ── Harness self-checks: the normalisation must not distort results ─────────

describe('preset parity harness', () => {
  it('treats different representations of the same geometry as equal', () => {
    // Full circle via ellipse() vs two chained half arcs.
    const a = recorder();
    a.beginPath();
    a.ellipse(50, 50, 40, 40, 0, 0, Math.PI * 2);
    a.closePath();
    const b = recorder();
    b.beginPath();
    b.moveTo(90, 50);
    b.ellipse(50, 50, 40, 40, 0, 0, Math.PI);
    b.ellipse(50, 50, 40, 40, 0, Math.PI, Math.PI * 2);
    b.closePath();
    expect(compareShapes(a, b).ok).toBe(true);

    // Square via rect() vs explicit moveTo/lineTo/closePath.
    const c = recorder();
    c.beginPath();
    c.rect(10, 10, 80, 40);
    const d = recorder();
    d.beginPath();
    d.moveTo(10, 10);
    d.lineTo(90, 10);
    d.lineTo(90, 50);
    d.lineTo(10, 50);
    d.closePath();
    expect(compareShapes(c, d).ok).toBe(true);
  });

  it('flags genuinely different geometry', () => {
    const a = recorder();
    a.beginPath();
    a.rect(10, 10, 80, 80);
    const b = recorder();
    b.beginPath();
    b.ellipse(50, 50, 40, 40, 0, 0, Math.PI * 2);
    b.closePath();
    expect(compareShapes(a, b).ok).toBe(false);
  });

  it('flags a sub-tolerance-scale but real offset', () => {
    const a = recorder();
    a.beginPath();
    a.rect(10, 10, 80, 80);
    const b = recorder();
    b.beginPath();
    b.rect(10.02, 10, 80, 80); // 0.02 px shift > 5e-3 tolerance
    const r = compareShapes(a, b);
    expect(r.ok).toBe(false);
    expect(r.maxDev).toBeGreaterThan(TOL);
  });

  it('distinguishes open from closed outlines', () => {
    const open = recorder();
    open.beginPath();
    open.moveTo(10, 10);
    open.lineTo(90, 10);
    open.lineTo(90, 90);
    open.lineTo(10, 90);
    const closed = recorder();
    closed.beginPath();
    closed.moveTo(10, 10);
    closed.lineTo(90, 10);
    closed.lineTo(90, 90);
    closed.lineTo(10, 90);
    closed.closePath();
    // Genuinely open outline (missing edge) ≠ closed square.
    expect(compareShapes(open, closed).ok).toBe(false);

    // Manually-looped square (coincident ends, CORNER seam) still differs:
    // abutting butt caps notch the corner where a join would be sharp.
    const looped = recorder();
    looped.beginPath();
    looped.moveTo(10, 10);
    looped.lineTo(90, 10);
    looped.lineTo(90, 90);
    looped.lineTo(10, 90);
    looped.lineTo(10, 10);
    expect(compareShapes(looped, closed).ok).toBe(false);

    // Full-circle loop with a SMOOTH seam is equivalent to its closed twin.
    const openCircle = recorder();
    openCircle.beginPath();
    openCircle.ellipse(50, 50, 30, 30, 0, 0, Math.PI * 2);
    const closedCircle = recorder();
    closedCircle.beginPath();
    closedCircle.ellipse(50, 50, 30, 30, 0, 0, Math.PI * 2);
    closedCircle.closePath();
    expect(compareShapes(openCircle, closedCircle).ok).toBe(true);
  });

  it('audits every live legacy label against a real presets.json entry', () => {
    for (const { spec } of LEGACY_SWITCH) {
      expect(PRESET_DEFS[spec], `spec target "${spec}"`).toBeDefined();
    }
    // No duplicates.
    const labels = LEGACY_SWITCH.map((e) => e.label);
    expect(new Set(labels).size).toBe(labels.length);
  });

  it('keeps every migrated preset pinned to spec-engine parity', () => {
    // For migrated presets buildShapePath must emit exactly the spec engine's
    // geometry — via the delegation route AND for every alias label that
    // historically resolved to the same shape. Checked at default and
    // perturbed adjusts so the adj plumbing (adj1..adj4 array) stays sound.
    const entries = LEGACY_SWITCH.filter((e) => SPEC_MIGRATED_PRESETS.has(e.spec));
    // Every migrated name must appear in the inventory (no orphan set entry).
    for (const name of SPEC_MIGRATED_PRESETS) {
      expect(
        entries.some((e) => e.spec === name),
        `SPEC_MIGRATED_PRESETS entry "${name}" missing from the audit inventory`,
      ).toBe(true);
      expect(PRESET_DEFS[name], `migrated "${name}" must exist in presets.json`).toBeDefined();
    }
    for (const e of entries) {
      for (const box of BOXES) {
        const dflt = compareOnce(e.label, e.spec, box, []);
        expect(dflt.ok, `${e.label} (default adj, ${box.w}×${box.h}): ${dflt.reason}`).toBe(true);
        const adj = perturbedAdj(e.spec);
        const pert = compareOnce(e.label, e.spec, box, adj);
        expect(pert.ok, `${e.label} (perturbed adj, ${box.w}×${box.h}): ${pert.reason}`).toBe(true);
      }
    }
  });

  it('writes the audit table when PRESET_PARITY_REPORT is a path', () => {
    const out = process.env.PRESET_PARITY_REPORT;
    if (!out) return;
    const rows = runAudit();
    const matched = rows.filter((r) => r.matchDefault);
    const lines: string[] = [];
    lines.push('| label | spec target | default adj | perturbed adj | fill-equiv | detail |');
    lines.push('|---|---|---|---|---|---|');
    for (const r of rows) {
      lines.push(
        `| ${r.label} | ${r.spec} | ${r.matchDefault ? 'MATCH' : 'DIFF'} | ${
          r.matchAdj == null ? '—' : r.matchAdj ? 'MATCH' : 'DIFF'
        } | ${r.fillEquiv == null ? '—' : r.fillEquiv ? 'YES' : 'no'} | ${r.detail} |`,
      );
    }
    lines.push('');
    lines.push(`matched(default): ${matched.length}/${rows.length}`);
    lines.push(
      `matched(default+adj): ${rows.filter((r) => r.matchDefault && r.matchAdj).length}/${rows.length}`,
    );
    lines.push(
      `diff but fill-equivalent: ${rows.filter((r) => !r.matchDefault && r.fillEquiv).length}`,
    );
    writeFileSync(out, lines.join('\n') + '\n');
  });
});
