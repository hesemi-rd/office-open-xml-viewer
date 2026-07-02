/**
 * Auxiliary (offscreen) canvas allocation shared across the renderers and the
 * metafile rasterizers.
 *
 * Lives under `canvas/` (a leaf utility layer) so both `shape/` (effects,
 * scene3d) and `image/` (wmf/emf/dib metafile players) can depend on it without
 * an unnatural `image → shape` edge. `shape/effects.ts` re-exports it for its
 * historical import path.
 */

/** A canvas usable off-DOM: an OffscreenCanvas in a worker, else a detached
 *  `<canvas>`. */
export type AuxCanvas = HTMLCanvasElement | OffscreenCanvas;
/** The 2D context type of an {@link AuxCanvas}. */
export type AuxContext = CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D;

/**
 * Allocate an auxiliary canvas. Prefers OffscreenCanvas (no DOM pollution, works
 * inside a worker); falls back to a detached <canvas>; returns null when neither
 * is available (e.g. a headless unit-test environment) so callers can no-op.
 * Dimensions are clamped to >=1 to avoid zero-size canvas errors.
 */
export function createAuxCanvas(w: number, h: number): AuxCanvas | null {
  const cw = Math.max(1, Math.ceil(w));
  const ch = Math.max(1, Math.ceil(h));
  if (typeof OffscreenCanvas !== 'undefined') {
    return new OffscreenCanvas(cw, ch);
  }
  if (typeof document !== 'undefined') {
    const c = document.createElement('canvas');
    c.width = cw;
    c.height = ch;
    return c;
  }
  return null;
}
