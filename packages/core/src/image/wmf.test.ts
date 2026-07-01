import { describe, it, expect, vi, beforeEach, afterEach } from 'vitest';
import {
  isWmf,
  isEmf,
  isMetafileMime,
  playWmf,
  renderWmfToBitmap,
  wmfRasterTarget,
  decodeRasterOrMetafile,
} from './wmf.js';

// ── WMF (Windows Metafile) player unit tests ────────────────────────────────
// The renderer falls back to this player for `.wmf`/`.emf` blips the browser
// can't decode (createImageBitmap throws on metafiles). image1.emf in
// sample-10.docx is, despite the extension, a *standard* (non-placeable) WMF
// whose labels are vector POLYPOLYGON glyph outlines (no text-out records), so
// the player only needs window mapping, an object table (pens+brushes), and
// POLYLINE/POLYGON/POLYPOLYGON/RECTANGLE drawing.
//
// `playWmf(bytes, ctx, W, H, suppressBoundaryFrame?)` is the pure record-replay
// core: it issues moveTo/lineTo/stroke/fill calls onto an injected ctx, so a
// recording mock pins coordinate mapping + state without needing OffscreenCanvas
// (absent in the node test env).
//
// Shared in core (originally docx-only). The window/device-boundary cosmetic
// stroke suppression is now a HEURISTIC gated behind `suppressBoundaryFrame`
// (default OFF = spec-clean, every edge drawn); docx opts IN when it re-points.

// ── byte builders ───────────────────────────────────────────────────────────

/** Little-endian byte writer for crafting WMF records. */
class Writer {
  private bytes: number[] = [];
  u16(v: number) {
    this.bytes.push(v & 0xff, (v >>> 8) & 0xff);
    return this;
  }
  i16(v: number) {
    return this.u16(v & 0xffff);
  }
  u32(v: number) {
    this.bytes.push(v & 0xff, (v >>> 8) & 0xff, (v >>> 16) & 0xff, (v >>> 24) & 0xff);
    return this;
  }
  raw(...vals: number[]) {
    for (const v of vals) this.bytes.push(v & 0xff);
    return this;
  }
  build(): Uint8Array {
    return new Uint8Array(this.bytes);
  }
}

/** Standard (non-placeable) 18-byte WMF header. numObjects defaults small. */
function wmfHeader(numObjects = 8): Uint8Array {
  return new Writer()
    .u16(1) // mtType = 1 (in-memory? 1 or 2 both legal; we accept both)
    .u16(9) // mtHeaderSize (words)
    .u16(0x0300) // mtVersion
    .u32(0) // mtSize (words) — players ignore for our purposes
    .u16(numObjects) // mtNoObjects
    .u32(0) // mtMaxRecord (words)
    .u16(0) // mtNoParameters
    .build();
}

/** A WMF record: u32 sizeWords (incl. the 6-byte size+function header), u16
 *  function, then params. `paramWords` is the number of 16-bit param words. */
function record(fn: number, params: (w: Writer) => void): Uint8Array {
  const pw = new Writer();
  params(pw);
  const paramBytes = pw.build();
  if (paramBytes.length % 2 !== 0) throw new Error('param bytes must be even');
  const sizeWords = 3 + paramBytes.length / 2; // 3 words = u32 size + u16 fn
  const rec = new Writer().u32(sizeWords).u16(fn);
  const head = rec.build();
  const out = new Uint8Array(head.length + paramBytes.length);
  out.set(head, 0);
  out.set(paramBytes, head.length);
  return out;
}

function concat(...parts: Uint8Array[]): Uint8Array {
  const total = parts.reduce((n, p) => n + p.length, 0);
  const out = new Uint8Array(total);
  let off = 0;
  for (const p of parts) {
    out.set(p, off);
    off += p.length;
  }
  return out;
}

// WMF record function codes
const FN = {
  EOF: 0x0000,
  SETPOLYFILLMODE: 0x0106,
  SETWINDOWORG: 0x020b,
  SETWINDOWEXT: 0x020c,
  SELECTOBJECT: 0x012d,
  DELETEOBJECT: 0x01f0,
  POLYGON: 0x0324,
  POLYLINE: 0x0325,
  POLYPOLYGON: 0x0538,
  RECTANGLE: 0x041b,
  CREATEPENINDIRECT: 0x02fa,
  CREATEBRUSHINDIRECT: 0x02fc,
  STRETCHDIBITS: 0x0f43,
} as const;

// ── recording mock ctx (records the draw calls + style mutations) ───────────

interface Call {
  op: string;
  args: number[];
}
interface MockCtx {
  ctx: CanvasRenderingContext2D;
  calls: Call[];
  styles: { fill: string[]; stroke: string[]; lineWidth: number[]; fillRules: (string | undefined)[] };
}

function makeRecordingCtx(): MockCtx {
  const calls: Call[] = [];
  const styles = { fill: [] as string[], stroke: [] as string[], lineWidth: [] as number[], fillRules: [] as (string | undefined)[] };
  let _fill = '#000';
  let _stroke = '#000';
  let _lw = 1;
  const ctx = {
    get fillStyle() { return _fill; },
    set fillStyle(v: string) { _fill = v; },
    get strokeStyle() { return _stroke; },
    set strokeStyle(v: string) { _stroke = v; },
    get lineWidth() { return _lw; },
    set lineWidth(v: number) { _lw = v; },
    lineJoin: 'miter' as CanvasLineJoin,
    lineCap: 'butt' as CanvasLineCap,
    save() { calls.push({ op: 'save', args: [] }); },
    restore() { calls.push({ op: 'restore', args: [] }); },
    beginPath() { calls.push({ op: 'beginPath', args: [] }); },
    closePath() { calls.push({ op: 'closePath', args: [] }); },
    moveTo(x: number, y: number) { calls.push({ op: 'moveTo', args: [x, y] }); },
    lineTo(x: number, y: number) { calls.push({ op: 'lineTo', args: [x, y] }); },
    rect(x: number, y: number, w: number, h: number) { calls.push({ op: 'rect', args: [x, y, w, h] }); },
    stroke() {
      calls.push({ op: 'stroke', args: [] });
      styles.stroke.push(_stroke);
      styles.lineWidth.push(_lw);
    },
    fill(rule?: string) {
      calls.push({ op: 'fill', args: [] });
      styles.fill.push(_fill);
      styles.fillRules.push(rule);
    },
  };
  return { ctx: ctx as unknown as CanvasRenderingContext2D, calls, styles };
}

// ── isWmf / isEmf detection ─────────────────────────────────────────────────

describe('isWmf / isEmf detection', () => {
  it('detects a standard (non-placeable) WMF header', () => {
    const bytes = wmfHeader();
    expect(isWmf(bytes)).toBe(true);
    expect(isEmf(bytes)).toBe(false);
  });

  it('detects mtType=2 as WMF too', () => {
    const w = new Writer().u16(2).u16(9).u16(0x0300).u32(0).u16(8).u32(0).u16(0);
    expect(isWmf(w.build())).toBe(true);
  });

  it('detects a placeable WMF via the D7CDC69A magic', () => {
    // 22-byte placeable header: magic, handle, bbox(4×i16), inch, reserved, checksum
    const placeable = new Writer()
      .raw(0xd7, 0xcd, 0xc6, 0x9a) // magic
      .u16(0) // hWmf
      .i16(0).i16(0).i16(100).i16(100) // bbox
      .u16(96) // inch
      .u32(0) // reserved
      .u16(0); // checksum
    // followed by a standard header
    const full = concat(placeable.build(), wmfHeader());
    expect(isWmf(full)).toBe(true);
  });

  it('detects a true EMF (ENHMETAHEADER) and does NOT treat it as WMF', () => {
    // u32@0 = 1 (EMR_HEADER iType), u32@40 = 0x464D4520 (" EMF" signature).
    const buf = new Uint8Array(48);
    const dv = new DataView(buf.buffer);
    dv.setUint32(0, 1, true);
    dv.setUint32(40, 0x464d4520, true);
    expect(isEmf(buf)).toBe(true);
    expect(isWmf(buf)).toBe(false);
  });

  it('rejects random bytes', () => {
    const rnd = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0xde, 0xad, 0xbe, 0xef]);
    expect(isWmf(rnd)).toBe(false);
    expect(isEmf(rnd)).toBe(false);
  });

  it('rejects a too-short buffer', () => {
    expect(isWmf(new Uint8Array([1, 0, 9]))).toBe(false);
    expect(isEmf(new Uint8Array([1, 0, 0, 0]))).toBe(false);
  });
});

// ── playWmf: minimal polyline replay with window mapping + pen select ────────

describe('playWmf — window mapping, pen, polyline', () => {
  it('replays a minimal WMF and maps logical→device coords with the current pen', () => {
    // Window org (0,0), ext (100,100); target bitmap 200×200 → scale ×2.
    // Pen: solid, color 0x00FF0000 = blue (COLORREF 0x00BBGGRR → R=0,G=0,B=0xFF).
    const file = concat(
      wmfHeader(),
      // SETWINDOWORG: y first, x second
      record(FN.SETWINDOWORG, (w) => w.i16(0).i16(0)),
      // SETWINDOWEXT: y first, x second → yExt=100, xExt=100
      record(FN.SETWINDOWEXT, (w) => w.i16(100).i16(100)),
      // CREATEPENINDIRECT: style=0 (solid), widthX=1, widthY=0, color 0x00FF0000
      record(FN.CREATEPENINDIRECT, (w) => w.u16(0).i16(1).i16(0).u32(0x00ff0000)),
      // SELECTOBJECT idx 0
      record(FN.SELECTOBJECT, (w) => w.u16(0)),
      // POLYLINE: 3 pts (10,20)-(30,40)-(50,60)
      record(FN.POLYLINE, (w) => w.i16(3).i16(10).i16(20).i16(30).i16(40).i16(50).i16(60)),
      record(FN.EOF, () => {}),
    );

    const m = makeRecordingCtx();
    const drew = playWmf(file, m.ctx, 200, 200);
    expect(drew).toBe(true);

    // Expect a moveTo to the first point then lineTo for the rest, mapped ×2.
    const moves = m.calls.filter((c) => c.op === 'moveTo');
    const lines = m.calls.filter((c) => c.op === 'lineTo');
    expect(moves.length).toBe(1);
    expect(moves[0].args).toEqual([20, 40]); // (10,20) × 2
    expect(lines.length).toBe(2);
    expect(lines[0].args).toEqual([60, 80]); // (30,40) × 2
    expect(lines[1].args).toEqual([100, 120]); // (50,60) × 2

    // The polyline strokes (no fill) with the selected blue pen.
    const strokes = m.calls.filter((c) => c.op === 'stroke');
    expect(strokes.length).toBe(1);
    expect(m.styles.stroke.at(-1)?.toLowerCase()).toBe('#0000ff'); // R=0 G=0 B=255
    // Polyline never fills.
    expect(m.calls.some((c) => c.op === 'fill')).toBe(false);
  });

  it('honors window origin and a negative ext (axis flip)', () => {
    // org (10,10); ext x=100, y=-100 (Y flips). Target 100×100 → |scale| ×1.
    const file = concat(
      wmfHeader(),
      record(FN.SETWINDOWORG, (w) => w.i16(10).i16(10)), // yOrg=10, xOrg=10
      record(FN.SETWINDOWEXT, (w) => w.i16(-100).i16(100)), // yExt=-100, xExt=100
      record(FN.CREATEPENINDIRECT, (w) => w.u16(0).i16(1).i16(0).u32(0x00000000)),
      record(FN.SELECTOBJECT, (w) => w.u16(0)),
      record(FN.POLYLINE, (w) => w.i16(2).i16(10).i16(10).i16(60).i16(60)),
      record(FN.EOF, () => {}),
    );
    const m = makeRecordingCtx();
    playWmf(file, m.ctx, 100, 100);
    const moves = m.calls.filter((c) => c.op === 'moveTo');
    // logical (10,10) is the origin → device (0,0) even with the flip.
    expect(moves[0].args).toEqual([0, -0]);
    const lines = m.calls.filter((c) => c.op === 'lineTo');
    // (60,60): dx=50 → x=50; dy=50, yExt=-100 → device y = 50 * (100 / -100) = -50.
    expect(lines[0].args[0]).toBe(50);
    expect(lines[0].args[1]).toBe(-50);
  });

  it('NULL-style pen (style 5) does not stroke', () => {
    const file = concat(
      wmfHeader(),
      record(FN.SETWINDOWORG, (w) => w.i16(0).i16(0)),
      record(FN.SETWINDOWEXT, (w) => w.i16(100).i16(100)),
      record(FN.CREATEPENINDIRECT, (w) => w.u16(5).i16(1).i16(0).u32(0x00000000)), // PS_NULL
      record(FN.SELECTOBJECT, (w) => w.u16(0)),
      record(FN.POLYLINE, (w) => w.i16(2).i16(0).i16(0).i16(10).i16(10)),
      record(FN.EOF, () => {}),
    );
    const m = makeRecordingCtx();
    playWmf(file, m.ctx, 100, 100);
    expect(m.calls.some((c) => c.op === 'stroke')).toBe(false);
  });
});

// ── playWmf: polygon / polypolygon fill + fill rule ─────────────────────────

describe('playWmf — polygon / polypolygon fill', () => {
  it('fills + strokes a POLYGON with the current brush + pen', () => {
    const file = concat(
      wmfHeader(),
      record(FN.SETWINDOWORG, (w) => w.i16(0).i16(0)),
      record(FN.SETWINDOWEXT, (w) => w.i16(10).i16(10)),
      // brush: SOLID (0), color green 0x0000FF00 (R=0,G=0xFF,B=0)
      record(FN.CREATEBRUSHINDIRECT, (w) => w.u16(0).u32(0x0000ff00).u16(0)),
      record(FN.SELECTOBJECT, (w) => w.u16(0)),
      // pen: solid red 0x000000FF (R=0xFF)
      record(FN.CREATEPENINDIRECT, (w) => w.u16(0).i16(1).i16(0).u32(0x000000ff)),
      record(FN.SELECTOBJECT, (w) => w.u16(1)),
      record(FN.POLYGON, (w) => w.i16(3).i16(0).i16(0).i16(10).i16(0).i16(5).i16(10)),
      record(FN.EOF, () => {}),
    );
    const m = makeRecordingCtx();
    expect(playWmf(file, m.ctx, 10, 10)).toBe(true);
    expect(m.styles.fill.at(-1)?.toLowerCase()).toBe('#00ff00'); // green brush
    expect(m.styles.stroke.at(-1)?.toLowerCase()).toBe('#ff0000'); // red pen
  });

  it('NULL brush (style 1) does not fill', () => {
    const file = concat(
      wmfHeader(),
      record(FN.SETWINDOWORG, (w) => w.i16(0).i16(0)),
      record(FN.SETWINDOWEXT, (w) => w.i16(10).i16(10)),
      record(FN.CREATEBRUSHINDIRECT, (w) => w.u16(1).u32(0x00000000).u16(0)), // BS_NULL
      record(FN.SELECTOBJECT, (w) => w.u16(0)),
      record(FN.CREATEPENINDIRECT, (w) => w.u16(0).i16(1).i16(0).u32(0x00000000)),
      record(FN.SELECTOBJECT, (w) => w.u16(1)),
      record(FN.POLYGON, (w) => w.i16(3).i16(0).i16(0).i16(10).i16(0).i16(5).i16(10)),
      record(FN.EOF, () => {}),
    );
    const m = makeRecordingCtx();
    playWmf(file, m.ctx, 10, 10);
    expect(m.calls.some((c) => c.op === 'fill')).toBe(false);
    expect(m.calls.some((c) => c.op === 'stroke')).toBe(true);
  });

  it('POLYPOLYGON honors SETPOLYFILLMODE (ALTERNATE → evenodd) for glyph holes', () => {
    const file = concat(
      wmfHeader(),
      record(FN.SETWINDOWORG, (w) => w.i16(0).i16(0)),
      record(FN.SETWINDOWEXT, (w) => w.i16(20).i16(20)),
      record(FN.SETPOLYFILLMODE, (w) => w.u16(1)), // ALTERNATE
      record(FN.CREATEBRUSHINDIRECT, (w) => w.u16(0).u32(0x00000000).u16(0)),
      record(FN.SELECTOBJECT, (w) => w.u16(0)),
      // 2 sub-polys: outer 4-gon, inner 4-gon (a hole). u16 numPolys, u16 counts, then pts.
      record(FN.POLYPOLYGON, (w) =>
        w
          .u16(2)
          .u16(4)
          .u16(4)
          // outer
          .i16(0).i16(0).i16(20).i16(0).i16(20).i16(20).i16(0).i16(20)
          // inner hole
          .i16(5).i16(5).i16(15).i16(5).i16(15).i16(15).i16(5).i16(15),
      ),
      record(FN.EOF, () => {}),
    );
    const m = makeRecordingCtx();
    expect(playWmf(file, m.ctx, 20, 20)).toBe(true);
    // One fill spanning both sub-paths, with the evenodd rule.
    const fills = m.calls.filter((c) => c.op === 'fill');
    expect(fills.length).toBeGreaterThanOrEqual(1);
    expect(m.styles.fillRules.at(-1)).toBe('evenodd');
    // Both sub-polygons contributed moveTo's (8 vertices → 2 moveTo + 6 lineTo at least).
    expect(m.calls.filter((c) => c.op === 'moveTo').length).toBeGreaterThanOrEqual(2);
  });
});

// ── playWmf: spec-clean default draws boundary edges (heuristic OFF) ─────────

describe('playWmf — spec-clean default (suppressBoundaryFrame=false) draws all edges', () => {
  it('a RECTANGLE coincident with the device bounds STILL strokes its outline by default', () => {
    // window org (0,0) ext (100,100), device 100×100 ⇒ logical==device. The rect
    // spans the full window; with the heuristic OFF every edge is drawn.
    const file = concat(
      wmfHeader(),
      record(FN.SETWINDOWORG, (w) => w.i16(0).i16(0)),
      record(FN.SETWINDOWEXT, (w) => w.i16(100).i16(100)),
      record(FN.CREATEPENINDIRECT, (w) => w.u16(0).i16(0).i16(0).u32(0x00000000)),
      record(FN.SELECTOBJECT, (w) => w.u16(0)),
      // RECTANGLE params: bottom, right, top, left (full window).
      record(FN.RECTANGLE, (w) => w.i16(100).i16(100).i16(0).i16(0)),
      record(FN.EOF, () => {}),
    );
    const m = makeRecordingCtx();
    playWmf(file, m.ctx, 100, 100); // default: suppressBoundaryFrame=false
    // Spec-clean: the boundary-coincident outline IS stroked as one closed path.
    expect(m.calls.some((c) => c.op === 'stroke')).toBe(true);
    expect(m.calls.filter((c) => c.op === 'moveTo').length).toBe(1);
    expect(m.calls.filter((c) => c.op === 'lineTo').length).toBe(3);
    expect(m.calls.some((c) => c.op === 'closePath')).toBe(true);
  });
});

// ── playWmf: window/device-boundary cosmetic-stroke suppression (heuristic) ──

describe('playWmf — window/device-boundary stroke suppression (suppressBoundaryFrame=true)', () => {
  // HEURISTIC (see deviceInteriorEdges in wmf.ts): a cosmetic stroke whose edge
  // coincides with the metafile window/device boundary (x∈{0,W} or y∈{0,H}) is
  // suppressed, because Word renders no visible frame there. This is NOT GDI's
  // actual clip (which excludes only the right/bottom edges); we drop all four
  // boundary lines to remove the common full-window "frame rectangle" drawn with
  // a 1px cosmetic pen. Opt-in via the suppressBoundaryFrame flag (docx behavior).

  it('a RECTANGLE coincident with the device bounds paints NO outline (frame suppressed)', () => {
    // window org (0,0) ext (100,100), device 100×100 ⇒ logical==device. The rect
    // spans the full window, so all four edges land on the boundary.
    const file = concat(
      wmfHeader(),
      record(FN.SETWINDOWORG, (w) => w.i16(0).i16(0)),
      record(FN.SETWINDOWEXT, (w) => w.i16(100).i16(100)),
      // cosmetic (width 0) PS_SOLID black pen — a typical full-window frame pen.
      record(FN.CREATEPENINDIRECT, (w) => w.u16(0).i16(0).i16(0).u32(0x00000000)),
      record(FN.SELECTOBJECT, (w) => w.u16(0)),
      // RECTANGLE params: bottom, right, top, left (full window).
      record(FN.RECTANGLE, (w) => w.i16(100).i16(100).i16(0).i16(0)),
      record(FN.EOF, () => {}),
    );
    const m = makeRecordingCtx();
    playWmf(file, m.ctx, 100, 100, true);
    // No brush selected ⇒ no fill; and the boundary-coincident outline is not
    // stroked ⇒ the rectangle contributes nothing.
    expect(m.calls.some((c) => c.op === 'stroke')).toBe(false);
    expect(m.calls.some((c) => c.op === 'fill')).toBe(false);
  });

  it('a RECTANGLE one pixel INSIDE the bounds still strokes all four edges (not a size rule)', () => {
    // Same pen/window, but the rect is inset by 1 unit on every side, so no edge
    // lies on the surface boundary — the outline must paint normally.
    const file = concat(
      wmfHeader(),
      record(FN.SETWINDOWORG, (w) => w.i16(0).i16(0)),
      record(FN.SETWINDOWEXT, (w) => w.i16(100).i16(100)),
      record(FN.CREATEPENINDIRECT, (w) => w.u16(0).i16(0).i16(0).u32(0x00000000)),
      record(FN.SELECTOBJECT, (w) => w.u16(0)),
      record(FN.RECTANGLE, (w) => w.i16(99).i16(99).i16(1).i16(1)),
      record(FN.EOF, () => {}),
    );
    const m = makeRecordingCtx();
    playWmf(file, m.ctx, 100, 100, true);
    expect(m.calls.some((c) => c.op === 'stroke')).toBe(true);
    // All four interior edges drawn: a single continuous sub-path (1 moveTo) with
    // four lineTo's (closed rectangle).
    expect(m.calls.filter((c) => c.op === 'moveTo').length).toBe(1);
    expect(m.calls.filter((c) => c.op === 'lineTo').length).toBe(4);
  });

  it('a boundary RECTANGLE WITH a brush still FILLS (only the cosmetic outline is suppressed)', () => {
    const file = concat(
      wmfHeader(),
      record(FN.SETWINDOWORG, (w) => w.i16(0).i16(0)),
      record(FN.SETWINDOWEXT, (w) => w.i16(100).i16(100)),
      // green solid brush + cosmetic black pen.
      record(FN.CREATEBRUSHINDIRECT, (w) => w.u16(0).u32(0x0000ff00).u16(0)),
      record(FN.SELECTOBJECT, (w) => w.u16(0)),
      record(FN.CREATEPENINDIRECT, (w) => w.u16(0).i16(0).i16(0).u32(0x00000000)),
      record(FN.SELECTOBJECT, (w) => w.u16(1)),
      record(FN.RECTANGLE, (w) => w.i16(100).i16(100).i16(0).i16(0)),
      record(FN.EOF, () => {}),
    );
    const m = makeRecordingCtx();
    playWmf(file, m.ctx, 100, 100, true);
    expect(m.calls.some((c) => c.op === 'fill')).toBe(true);
    expect(m.styles.fill.at(-1)?.toLowerCase()).toBe('#00ff00');
    // Outline still suppressed (all edges on the boundary).
    expect(m.calls.some((c) => c.op === 'stroke')).toBe(false);
  });

  it('only the boundary edges of a partially-coincident polygon are dropped', () => {
    // A right triangle whose bottom edge runs along y=0 (on the boundary) but
    // whose other two edges are interior: the bottom edge is dropped, the other
    // two still stroke.
    const file = concat(
      wmfHeader(),
      record(FN.SETWINDOWORG, (w) => w.i16(0).i16(0)),
      record(FN.SETWINDOWEXT, (w) => w.i16(100).i16(100)),
      record(FN.CREATEPENINDIRECT, (w) => w.u16(0).i16(0).i16(0).u32(0x000000ff)),
      record(FN.SELECTOBJECT, (w) => w.u16(0)),
      // vertices (0,0)-(100,0)-(50,50): edge (0,0)->(100,0) is on y=0.
      record(FN.POLYGON, (w) => w.i16(3).i16(0).i16(0).i16(100).i16(0).i16(50).i16(50)),
      record(FN.EOF, () => {}),
    );
    const m = makeRecordingCtx();
    playWmf(file, m.ctx, 100, 100, true);
    expect(m.calls.some((c) => c.op === 'stroke')).toBe(true);
    // Two surviving edges, contiguous: (100,0)->(50,50)->(0,0). One sub-path.
    expect(m.calls.filter((c) => c.op === 'moveTo').length).toBe(1);
    expect(m.calls.filter((c) => c.op === 'lineTo').length).toBe(2);
  });
});

// ── playWmf: object table create/select/delete/reuse ────────────────────────

describe('playWmf — object table create / delete / slot reuse', () => {
  it('reuses the freed slot when a deleted object index is later recreated', () => {
    // Create pen#0 (blue), pen#1 (green); select#0; delete#0; create pen (red) →
    // must land in the freed slot 0; select#0 → current pen is red.
    const file = concat(
      wmfHeader(),
      record(FN.SETWINDOWORG, (w) => w.i16(0).i16(0)),
      record(FN.SETWINDOWEXT, (w) => w.i16(10).i16(10)),
      record(FN.CREATEPENINDIRECT, (w) => w.u16(0).i16(1).i16(0).u32(0x00ff0000)), // slot0 blue
      record(FN.CREATEPENINDIRECT, (w) => w.u16(0).i16(1).i16(0).u32(0x0000ff00)), // slot1 green
      record(FN.SELECTOBJECT, (w) => w.u16(0)),
      record(FN.DELETEOBJECT, (w) => w.u16(0)), // free slot0
      record(FN.CREATEPENINDIRECT, (w) => w.u16(0).i16(1).i16(0).u32(0x000000ff)), // → slot0 red
      record(FN.SELECTOBJECT, (w) => w.u16(0)), // select slot0 (now red)
      record(FN.POLYLINE, (w) => w.i16(2).i16(0).i16(0).i16(10).i16(10)),
      record(FN.EOF, () => {}),
    );
    const m = makeRecordingCtx();
    playWmf(file, m.ctx, 10, 10);
    // The stroke uses the recreated red pen in the reused slot 0.
    expect(m.styles.stroke.at(-1)?.toLowerCase()).toBe('#ff0000');
  });

  it('selecting a brush vs a pen routes by object kind', () => {
    // slot0 = brush(green), slot1 = pen(red). select#0 (brush) then #1 (pen).
    const file = concat(
      wmfHeader(),
      record(FN.SETWINDOWORG, (w) => w.i16(0).i16(0)),
      record(FN.SETWINDOWEXT, (w) => w.i16(10).i16(10)),
      record(FN.CREATEBRUSHINDIRECT, (w) => w.u16(0).u32(0x0000ff00).u16(0)), // slot0 brush green
      record(FN.CREATEPENINDIRECT, (w) => w.u16(0).i16(1).i16(0).u32(0x000000ff)), // slot1 pen red
      record(FN.SELECTOBJECT, (w) => w.u16(0)), // brush
      record(FN.SELECTOBJECT, (w) => w.u16(1)), // pen
      record(FN.POLYGON, (w) => w.i16(3).i16(0).i16(0).i16(10).i16(0).i16(5).i16(10)),
      record(FN.EOF, () => {}),
    );
    const m = makeRecordingCtx();
    playWmf(file, m.ctx, 10, 10);
    expect(m.styles.fill.at(-1)?.toLowerCase()).toBe('#00ff00');
    expect(m.styles.stroke.at(-1)?.toLowerCase()).toBe('#ff0000');
  });
});

// ── playWmf: graceful bail on malformed records ─────────────────────────────

describe('playWmf — robustness', () => {
  it('bails gracefully on a record claiming a bogus (too-small) size', () => {
    const bad = concat(
      wmfHeader(),
      record(FN.SETWINDOWORG, (w) => w.i16(0).i16(0)),
      record(FN.SETWINDOWEXT, (w) => w.i16(10).i16(10)),
      record(FN.CREATEPENINDIRECT, (w) => w.u16(0).i16(1).i16(0).u32(0)),
      record(FN.SELECTOBJECT, (w) => w.u16(0)),
      record(FN.POLYLINE, (w) => w.i16(2).i16(0).i16(0).i16(5).i16(5)),
      // a deliberately corrupt record: sizeWords = 1 (< 3) — must stop the loop.
      new Writer().u32(1).u16(FN.POLYLINE).build(),
      record(FN.EOF, () => {}),
    );
    const m = makeRecordingCtx();
    // Whatever was drawn before the corrupt record stands; no throw.
    expect(() => playWmf(bad, m.ctx, 10, 10)).not.toThrow();
    expect(m.calls.some((c) => c.op === 'stroke')).toBe(true);
  });

  it('returns false for non-WMF bytes', () => {
    const m = makeRecordingCtx();
    expect(playWmf(new Uint8Array([0xde, 0xad, 0xbe, 0xef, 0, 0, 0, 0]), m.ctx, 10, 10)).toBe(false);
  });
});

// ── playWmf: STRETCHDIBITS embedded raster DIB ──────────────────────────────

describe('playWmf — STRETCHDIBITS (embedded raster DIB)', () => {
  // A META_STRETCHDIBITS record wraps a packed DIB. Its params (after the 6-byte
  // header) are: u32 RasterOp, i16 SrcHeight/SrcWidth/YSrc/XSrc, u16 UsageSrc,
  // i16 DestHeight/DestWidth/YDest/XDest, then the packed DIB. OffscreenCanvas is
  // absent in the node test env, so blitDibToCtx returns false (no draw) — we
  // assert the record parses without throwing and that later records still run.

  /** Append a packed top-down 2×2 24-bit BI_RGB DIB (40-byte header + 16 pixel
   *  bytes: 2 rows × 8-byte stride). Returns the same Writer for chaining. */
  function packed2x2Dib24(w: Writer): Writer {
    // BITMAPINFOHEADER (top-down: height = -2).
    w.u32(40).u32(2).u32((-2) >>> 0).u16(1).u16(24).u32(0).u32(0).u32(0).u32(0).u32(0).u32(0);
    // 2 rows, BGR pixels, each row padded to an 8-byte stride.
    w.raw(0, 0, 255, 0, 255, 0, 0, 0); // row0: red (B0 G0 R255), green (B0 G255 R0), pad
    w.raw(255, 0, 0, 30, 20, 10, 0, 0); // row1: blue, (10,20,30), pad
    return w;
  }

  it('parses a STRETCHDIBITS record without throwing and continues to later records', () => {
    const m = makeRecordingCtx();
    const file = concat(
      wmfHeader(),
      record(FN.SETWINDOWORG, (w) => w.i16(0).i16(0)),
      record(FN.SETWINDOWEXT, (w) => w.i16(100).i16(100)),
      // STRETCHDIBITS: draw the DIB into dest rect (XDest=0,YDest=0, 50×50).
      record(FN.STRETCHDIBITS, (w) => {
        w.u32(0x00cc0020); // RasterOperation SRCCOPY (ignored)
        w.i16(2).i16(2); // SrcHeight, SrcWidth
        w.i16(0).i16(0); // YSrc, XSrc
        w.u16(0); // UsageSrc (DIB_RGB_COLORS)
        w.i16(50).i16(50); // DestHeight, DestWidth
        w.i16(0).i16(0); // YDest, XDest
        packed2x2Dib24(w); // packed DIB
      }),
      // A polyline AFTER the blt must still execute (proves the loop advanced past
      // the STRETCHDIBITS record by its size, not by mis-parsing the DIB bytes).
      record(FN.CREATEPENINDIRECT, (w) => w.u16(0).i16(1).i16(0).u32(0x000000ff)),
      record(FN.SELECTOBJECT, (w) => w.u16(0)),
      record(FN.POLYLINE, (w) => w.i16(2).i16(10).i16(10).i16(20).i16(20)),
      record(FN.EOF, () => {}),
    );

    // No OffscreenCanvas in the node env ⇒ blitDibToCtx returns false (no draw);
    // the parse path must still run cleanly and the trailing polyline strokes.
    let drew = false;
    expect(() => {
      drew = playWmf(file, m.ctx, 100, 100);
    }).not.toThrow();
    // The trailing polyline drew, so playWmf reports true and the loop advanced
    // past the STRETCHDIBITS record correctly.
    expect(drew).toBe(true);
    const strokes = m.calls.filter((c) => c.op === 'stroke');
    expect(strokes.length).toBe(1);
    expect(m.styles.stroke.at(-1)?.toLowerCase()).toBe('#ff0000');
    // The polyline endpoints mapped ×1 (window 100 → device 100).
    const moves = m.calls.filter((c) => c.op === 'moveTo');
    expect(moves[0].args).toEqual([10, 10]);
  });
});

// ── wmfRasterTarget: supersampled, capped sizing ────────────────────────────

describe('wmfRasterTarget', () => {
  it('supersamples the intended pt size by 2×', () => {
    expect(wmfRasterTarget(100, 50)).toEqual({ w: 200, h: 100 });
  });

  it('falls back to a 300pt square (×2 → 600px) when size is unknown (0)', () => {
    expect(wmfRasterTarget(0, 0)).toEqual({ w: 600, h: 600 });
  });

  it('caps each dimension at 2000px', () => {
    expect(wmfRasterTarget(5000, 5000)).toEqual({ w: 2000, h: 2000 });
  });
});

// ── renderWmfToBitmap: OffscreenCanvas wrapper (browser/worker only) ─────────

describe('renderWmfToBitmap', () => {
  beforeEach(() => {
    // OffscreenCanvas + createImageBitmap don't exist in the node test env.
    const recorded: { ctx: unknown } = { ctx: null };
    vi.stubGlobal(
      'OffscreenCanvas',
      class {
        width: number;
        height: number;
        constructor(w: number, h: number) {
          this.width = w;
          this.height = h;
        }
        getContext() {
          const ctx = {
            fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
            lineJoin: 'miter', lineCap: 'butt',
            save() {}, restore() {}, beginPath() {}, closePath() {},
            moveTo() {}, lineTo() {}, rect() {}, stroke() {}, fill() {},
          };
          recorded.ctx = ctx;
          return ctx;
        }
      },
    );
    vi.stubGlobal(
      'createImageBitmap',
      vi.fn(async (src: { width: number; height: number }) => ({ width: src.width, height: src.height, close() {} }) as unknown as ImageBitmap),
    );
  });
  afterEach(() => vi.unstubAllGlobals());

  it('rasterizes a minimal WMF to an ImageBitmap of the target size', async () => {
    const file = concat(
      wmfHeader(),
      record(FN.SETWINDOWORG, (w) => w.i16(0).i16(0)),
      record(FN.SETWINDOWEXT, (w) => w.i16(100).i16(100)),
      record(FN.CREATEPENINDIRECT, (w) => w.u16(0).i16(1).i16(0).u32(0)),
      record(FN.SELECTOBJECT, (w) => w.u16(0)),
      record(FN.POLYLINE, (w) => w.i16(2).i16(0).i16(0).i16(50).i16(50)),
      record(FN.EOF, () => {}),
    );
    const bmp = await renderWmfToBitmap(file, 64, 48);
    expect(bmp).not.toBeNull();
    expect(bmp?.width).toBe(64);
    expect(bmp?.height).toBe(48);
  });

  it('returns null for non-WMF bytes', async () => {
    const bmp = await renderWmfToBitmap(new Uint8Array([1, 2, 3, 4]), 10, 10);
    expect(bmp).toBeNull();
  });

  it('returns null when nothing draws (empty metafile)', async () => {
    const file = concat(wmfHeader(), record(FN.EOF, () => {}));
    const bmp = await renderWmfToBitmap(file, 10, 10);
    expect(bmp).toBeNull();
  });
});

// ── decodeRasterOrMetafile: the shared raster/metafile decoder ───────────────

describe('decodeRasterOrMetafile', () => {
  beforeEach(() => {
    // OffscreenCanvas + createImageBitmap don't exist in the node test env.
    vi.stubGlobal(
      'OffscreenCanvas',
      class {
        width: number;
        height: number;
        constructor(w: number, h: number) {
          this.width = w;
          this.height = h;
        }
        getContext() {
          return {
            fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
            lineJoin: 'miter', lineCap: 'butt',
            save() {}, restore() {}, beginPath() {}, closePath() {},
            moveTo() {}, lineTo() {}, rect() {}, stroke() {}, fill() {},
          };
        }
      },
    );
  });
  afterEach(() => vi.unstubAllGlobals());

  it('WMF bytes → rasterizes through the player (non-null bitmap), sized via wmfRasterTarget', async () => {
    // createImageBitmap here only ever sees the OffscreenCanvas (the WMF branch
    // never calls it on the blob), so report the canvas dims back as the bitmap.
    const cib = vi.fn(async (src: { width: number; height: number }) =>
      ({ width: src.width, height: src.height, close() {} }) as unknown as ImageBitmap);
    vi.stubGlobal('createImageBitmap', cib);
    const file = concat(
      wmfHeader(),
      record(FN.SETWINDOWORG, (w) => w.i16(0).i16(0)),
      record(FN.SETWINDOWEXT, (w) => w.i16(100).i16(100)),
      record(FN.CREATEPENINDIRECT, (w) => w.u16(0).i16(1).i16(0).u32(0)),
      record(FN.SELECTOBJECT, (w) => w.u16(0)),
      record(FN.POLYLINE, (w) => w.i16(2).i16(0).i16(0).i16(50).i16(50)),
      record(FN.EOF, () => {}),
    );
    const blob = new Blob([file as BlobPart], { type: 'image/wmf' });
    const bmp = await decodeRasterOrMetafile(blob, { widthPt: 100, heightPt: 50 });
    expect(bmp).not.toBeNull();
    // wmfRasterTarget(100,50) = 200×100 (×2 supersample).
    expect(bmp?.width).toBe(200);
    expect(bmp?.height).toBe(100);
    // The WMF branch rasterizes the OffscreenCanvas, never the blob directly.
    expect(cib).toHaveBeenCalledTimes(1);
    const arg = cib.mock.calls[0][0] as { width: number; height: number };
    expect(arg.width).toBe(200);
    expect(arg.height).toBe(100);
  });

  it('EMF magic → null (skipped gracefully, no throw, no createImageBitmap on the blob)', async () => {
    const cib = vi.fn(async () => ({ width: 1, height: 1, close() {} }) as unknown as ImageBitmap);
    vi.stubGlobal('createImageBitmap', cib);
    const buf = new Uint8Array(48);
    const dv = new DataView(buf.buffer);
    dv.setUint32(0, 1, true); // EMR_HEADER iType
    dv.setUint32(40, 0x464d4520, true); // " EMF" signature
    const blob = new Blob([buf as BlobPart], { type: 'image/emf' });
    const bmp = await decodeRasterOrMetafile(blob, { widthPt: 100, heightPt: 100 });
    expect(bmp).toBeNull();
    expect(cib).not.toHaveBeenCalled();
  });

  it('a PNG blob → the createImageBitmap path (returns the decoded bitmap)', async () => {
    const fake = { width: 7, height: 9, close() {} } as unknown as ImageBitmap;
    const cib = vi.fn(async (_blob: Blob) => fake);
    vi.stubGlobal('createImageBitmap', cib);
    // PNG magic 89 50 4E 47 0D 0A 1A 0A — not WMF, not EMF.
    const png = new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0, 0, 0, 0]);
    const blob = new Blob([png as BlobPart], { type: 'image/png' });
    const bmp = await decodeRasterOrMetafile(blob, { widthPt: 50, heightPt: 50 });
    expect(bmp).toBe(fake);
    expect(cib).toHaveBeenCalledTimes(1);
    expect(cib).toHaveBeenCalledWith(blob);
  });

  it('an empty WMF (no geometry) → null, not a throw', async () => {
    vi.stubGlobal('createImageBitmap', vi.fn(async (src: { width: number; height: number }) =>
      ({ width: src.width, height: src.height, close() {} }) as unknown as ImageBitmap));
    const file = concat(wmfHeader(), record(FN.EOF, () => {}));
    const blob = new Blob([file as BlobPart], { type: 'image/wmf' });
    const bmp = await decodeRasterOrMetafile(blob, { widthPt: 100, heightPt: 100 });
    expect(bmp).toBeNull();
  });

  it('suppressBoundaryFrame defaults OFF (full-window rect draws its frame) and can be opted IN', async () => {
    // Capture whether the player stroked by recording on the offscreen ctx.
    let strokeCount = 0;
    vi.stubGlobal(
      'OffscreenCanvas',
      class {
        width: number;
        height: number;
        constructor(w: number, h: number) { this.width = w; this.height = h; }
        getContext() {
          return {
            fillStyle: '#000', strokeStyle: '#000', lineWidth: 1,
            lineJoin: 'miter', lineCap: 'butt',
            save() {}, restore() {}, beginPath() {}, closePath() {},
            moveTo() {}, lineTo() {}, rect() {},
            stroke() { strokeCount++; }, fill() {},
          };
        }
      },
    );
    vi.stubGlobal('createImageBitmap', vi.fn(async (src: { width: number; height: number }) =>
      ({ width: src.width, height: src.height, close() {} }) as unknown as ImageBitmap));
    // Full-window rectangle with a cosmetic pen at device == window (200×200 for
    // widthPt/heightPt 100 → ×2). The window ext must match the device so the
    // rect edges land exactly on the boundary.
    const mkFile = () =>
      concat(
        wmfHeader(),
        record(FN.SETWINDOWORG, (w) => w.i16(0).i16(0)),
        record(FN.SETWINDOWEXT, (w) => w.i16(200).i16(200)),
        record(FN.CREATEPENINDIRECT, (w) => w.u16(0).i16(0).i16(0).u32(0x00000000)),
        record(FN.SELECTOBJECT, (w) => w.u16(0)),
        record(FN.RECTANGLE, (w) => w.i16(200).i16(200).i16(0).i16(0)),
        record(FN.EOF, () => {}),
      );

    strokeCount = 0;
    const def = await decodeRasterOrMetafile(new Blob([mkFile() as BlobPart], { type: 'image/wmf' }), { widthPt: 100, heightPt: 100 });
    // Default OFF: the frame strokes, so the bitmap is non-null.
    expect(strokeCount).toBe(1);
    expect(def).not.toBeNull();

    strokeCount = 0;
    const sup = await decodeRasterOrMetafile(new Blob([mkFile() as BlobPart], { type: 'image/wmf' }), {
      widthPt: 100, heightPt: 100, suppressBoundaryFrame: true,
    });
    // Opted IN: every edge is on the boundary → nothing strokes → no geometry → null.
    expect(strokeCount).toBe(0);
    expect(sup).toBeNull();
  });
});

describe('isMetafileMime', () => {
  it('is true only for the WMF/EMF metafile MIME types', () => {
    expect(isMetafileMime('image/wmf')).toBe(true);
    expect(isMetafileMime('image/emf')).toBe(true);
    expect(isMetafileMime('image/png')).toBe(false);
    expect(isMetafileMime('image/jpeg')).toBe(false);
    expect(isMetafileMime('image/svg+xml')).toBe(false);
    expect(isMetafileMime(undefined)).toBe(false);
  });
});
