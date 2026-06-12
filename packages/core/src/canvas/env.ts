/**
 * Worker-safe environment guards shared by the renderers.
 *
 * `HTMLCanvasElement` and `window` are not defined inside a Web Worker, so a
 * bare `x instanceof HTMLCanvasElement` (or a bare `devicePixelRatio`
 * identifier) throws a ReferenceError there. Route every render-path check
 * through these helpers so the same code runs on the main thread and in the
 * render worker.
 */

/** True when `target` is a DOM canvas. Safe to call in a worker (returns false). */
export function isHTMLCanvas(target: unknown): target is HTMLCanvasElement {
  return typeof HTMLCanvasElement !== 'undefined' && target instanceof HTMLCanvasElement;
}

/** `window.devicePixelRatio` on the main thread; `fallback` in a worker. */
export function defaultDpr(fallback = 1): number {
  return typeof window !== 'undefined' ? (window.devicePixelRatio || fallback) : fallback;
}
