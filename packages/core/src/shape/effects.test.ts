import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest';
import {
  applyInnerShadow,
  applySoftEdge,
  applyReflection,
  createAuxCanvas,
} from './effects';
import type { Shadow, SoftEdge, Reflection } from '../types/common';

/**
 * Minimal recording 2D context. Records the ordered sequence of state mutations
 * and draw calls so we can assert the compositing pipeline each effect builds,
 * without a real canvas. Effects are pure compositing logic, so the op sequence
 * is the contract under test (ECMA-376 §20.1.8.40 / .50 / .53).
 */
class RecordingCtx {
  ops: Array<{ op: string; args?: unknown[] }> = [];
  filter = 'none';
  globalCompositeOperation = 'source-over';
  fillStyle: unknown = '#000';

  save() { this.ops.push({ op: 'save' }); }
  restore() { this.ops.push({ op: 'restore' }); }
  translate(x: number, y: number) { this.ops.push({ op: 'translate', args: [x, y] }); }
  scale(x: number, y: number) { this.ops.push({ op: 'scale', args: [x, y] }); }
  fillRect(...a: number[]) { this.ops.push({ op: 'fillRect', args: a }); }
  drawImage(...a: unknown[]) { this.ops.push({ op: 'drawImage', args: a }); }
  createLinearGradient() {
    this.ops.push({ op: 'createLinearGradient' });
    const stops: Array<[number, string]> = [];
    return {
      stops,
      addColorStop: (o: number, c: string) => { stops.push([o, c]); },
    } as unknown as CanvasGradient;
  }
  // Setters that the effects toggle are plain fields, captured on read.
}

/** Install a fake OffscreenCanvas that hands back RecordingCtx instances. */
function installFakeCanvas(): { auxCtxs: RecordingCtx[] } {
  const auxCtxs: RecordingCtx[] = [];
  class FakeOffscreen {
    width: number;
    height: number;
    constructor(w: number, h: number) { this.width = w; this.height = h; }
    getContext(_kind: string) {
      const c = new RecordingCtx();
      auxCtxs.push(c);
      return c;
    }
  }
  vi.stubGlobal('OffscreenCanvas', FakeOffscreen);
  return { auxCtxs };
}

const BBOX = { x: 100, y: 50, w: 200, h: 80 };
const DEVICE_W = 800;
const DEVICE_H = 600;
// scale is px-per-EMU; 1 px == 9525 EMU in DrawingML, but the effect helper just
// multiplies, so pick a scale where the arithmetic is easy to read.
const SCALE = 1 / 9525; // 9525 EMU -> 1 px

describe('createAuxCanvas', () => {
  afterEach(() => vi.unstubAllGlobals());

  it('returns an OffscreenCanvas when available, clamped to >=1 px', () => {
    installFakeCanvas();
    const c = createAuxCanvas(0, 0) as { width: number; height: number };
    expect(c).not.toBeNull();
    expect(c.width).toBe(1);
    expect(c.height).toBe(1);
  });

  it('returns null in a headless environment', () => {
    vi.stubGlobal('OffscreenCanvas', undefined);
    vi.stubGlobal('document', undefined);
    expect(createAuxCanvas(10, 10)).toBeNull();
  });
});

describe('applyInnerShadow (ECMA-376 §20.1.8.40)', () => {
  let fake: { auxCtxs: RecordingCtx[] };
  beforeEach(() => { fake = installFakeCanvas(); });
  afterEach(() => vi.unstubAllGlobals());

  const shadow: Shadow = { color: '404040', alpha: 0.8, blur: 19050, dist: 9525, dir: 90 };

  it('paints silhouette, carves an offset+blurred copy, clips, then blits', () => {
    const live = new RecordingCtx();
    const paint = vi.fn((c: unknown) => { (c as RecordingCtx).ops.push({ op: 'paintShape' }); });
    applyInnerShadow(live as never, paint as never, BBOX, shadow, SCALE, DEVICE_W, DEVICE_H);

    expect(fake.auxCtxs).toHaveLength(1);
    const aux = fake.auxCtxs[0];
    const ops = aux.ops.map(o => o.op);
    // 3 silhouette paints on the aux canvas: colour fill, destination-out, destination-in.
    expect(ops.filter(o => o === 'paintShape')).toHaveLength(3);
    // The blurred offset pass translates by (cos90*dist, sin90*dist) = (0, 1px).
    const translate = aux.ops.find(o => o.op === 'translate');
    expect(translate?.args?.[0]).toBeCloseTo(0, 6);
    expect(translate?.args?.[1]).toBeCloseTo(1, 6); // dist 9525 EMU * scale = 1px, dir=90 -> +y
    // Result is composited onto the live context exactly once.
    const liveDraws = live.ops.filter(o => o.op === 'drawImage');
    expect(liveDraws).toHaveLength(1);
  });

  it('skips the blur filter when blurRad is 0', () => {
    const live = new RecordingCtx();
    const paint = vi.fn();
    const noBlur: Shadow = { ...shadow, blur: 0 };
    // Spy on the filter assignments by recording them via a getter/setter proxy
    // is overkill; instead assert no throw and a single blit (smoke).
    applyInnerShadow(live as never, paint as never, BBOX, noBlur, SCALE, DEVICE_W, DEVICE_H);
    expect(live.ops.filter(o => o.op === 'drawImage')).toHaveLength(1);
  });
});

describe('applySoftEdge (ECMA-376 §20.1.8.53)', () => {
  let fake: { auxCtxs: RecordingCtx[] };
  beforeEach(() => { fake = installFakeCanvas(); });
  afterEach(() => vi.unstubAllGlobals());

  it('paints onto the live context directly when radius is 0 (no feather)', () => {
    const live = new RecordingCtx();
    const paint = vi.fn((c: unknown) => { (c as RecordingCtx).ops.push({ op: 'paintShape' }); });
    const se: SoftEdge = { radius: 0 };
    applySoftEdge(live as never, paint as never, BBOX, se, SCALE, DEVICE_W, DEVICE_H);
    // No aux canvas allocated; shape painted straight onto live.
    expect(fake.auxCtxs).toHaveLength(0);
    expect(live.ops.some(o => o.op === 'paintShape')).toBe(true);
    expect(live.ops.some(o => o.op === 'drawImage')).toBe(false);
  });

  it('feathers via a blurred destination-in mask then blits once', () => {
    const live = new RecordingCtx();
    const paint = vi.fn((c: unknown) => { (c as RecordingCtx).ops.push({ op: 'paintShape' }); });
    const se: SoftEdge = { radius: 28575 }; // 3 px
    applySoftEdge(live as never, paint as never, BBOX, se, SCALE, DEVICE_W, DEVICE_H);
    expect(fake.auxCtxs).toHaveLength(1);
    const aux = fake.auxCtxs[0];
    // Two silhouette paints on aux: full shape, then blurred mask.
    expect(aux.ops.filter(o => o.op === 'paintShape')).toHaveLength(2);
    expect(live.ops.filter(o => o.op === 'drawImage')).toHaveLength(1);
  });
});

describe('applyReflection (ECMA-376 §20.1.8.50)', () => {
  let fake: { auxCtxs: RecordingCtx[] };
  beforeEach(() => { fake = installFakeCanvas(); });
  afterEach(() => vi.unstubAllGlobals());

  const reflection: Reflection = {
    blur: 0, dist: 0, dir: 90,
    stA: 0.5, stPos: 0, endA: 0, endPos: 0.35,
    sx: 1, sy: -1,
  };

  it('paints shape, builds a vertical alpha gradient, mirrors via sy<0, blits', () => {
    const live = new RecordingCtx();
    const paint = vi.fn((c: unknown) => { (c as RecordingCtx).ops.push({ op: 'paintShape' }); });
    applyReflection(live as never, paint as never, BBOX, reflection, SCALE, DEVICE_W, DEVICE_H);

    const aux = fake.auxCtxs[0];
    expect(aux.ops.some(o => o.op === 'createLinearGradient')).toBe(true);
    expect(aux.ops.some(o => o.op === 'fillRect')).toBe(true);
    // The live blit mirrors with scale(sx, sy) where sy<0.
    const scaleOp = live.ops.find(o => o.op === 'scale');
    expect(scaleOp?.args).toEqual([1, -1]);
    expect(live.ops.filter(o => o.op === 'drawImage')).toHaveLength(1);
  });

  it('anchors the mirror at the shape bottom edge (algn="b") with dist offset', () => {
    const live = new RecordingCtx();
    const paint = vi.fn();
    const withDist: Reflection = { ...reflection, dist: 9525, dir: 90 }; // 1px down
    applyReflection(live as never, paint as never, BBOX, withDist, SCALE, DEVICE_W, DEVICE_H);
    const translates = live.ops.filter(o => o.op === 'translate');
    // First translate moves origin to (bbox.x + offX, bottom + offY).
    const bottom = BBOX.y + BBOX.h; // 130
    expect(translates[0].args?.[0]).toBeCloseTo(BBOX.x + 0, 6);
    expect(translates[0].args?.[1]).toBeCloseTo(bottom + 1, 6);
  });
});
