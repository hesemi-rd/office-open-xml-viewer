import { createAuxCanvasForContext } from '../canvas/aux-canvas.js';

type Ctx2D = CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D;

interface CacheState {
  cache: Map<string, number>;
  epoch: number;
}

interface FontSetLike {
  addEventListener?: (name: string, listener: () => void) => void;
}

interface DocumentLike {
  fonts?: FontSetLike;
}

const PROLONGED_SOUND_MARK = 0x30fc;
const RASTER_SIZE_PX = 512;
const RASTER_FONT_PX = 384;
const documentStates = new WeakMap<object, CacheState>();
const surfaceStates = new WeakMap<object, CacheState>();
const fallbackState: CacheState = { cache: new Map(), epoch: 0 };

function replaceFontSize(font: string, replacement: string): string {
  return font.replace(
    /(^|\s)\d*\.?\d+(?:px|pt|pc|in|cm|mm|q|em|rem|%)(?:\/[^\s]+)?(?=\s)/i,
    `$1${replacement}`,
  );
}

function sourceDocument(ctx: Ctx2D): DocumentLike | null {
  const canvas = ctx.canvas as { ownerDocument?: DocumentLike };
  if (canvas.ownerDocument) return canvas.ownerDocument;
  return typeof document !== 'undefined' ? document : null;
}

function cacheState(ctx: Ctx2D): CacheState {
  const doc = sourceDocument(ctx);
  if (doc !== null && typeof doc === 'object') {
    const existing = documentStates.get(doc);
    if (existing) return existing;
    const state: CacheState = { cache: new Map(), epoch: 0 };
    const invalidate = () => {
      state.epoch += 1;
      state.cache.clear();
    };
    doc.fonts?.addEventListener?.('loadingdone', invalidate);
    doc.fonts?.addEventListener?.('loadingerror', invalidate);
    documentStates.set(doc, state);
    return state;
  }

  const surface = ctx.canvas;
  if (surface !== null && (typeof surface === 'object' || typeof surface === 'function')) {
    const existing = surfaceStates.get(surface);
    if (existing) return existing;
    const state: CacheState = { cache: new Map(), epoch: 0 };
    surfaceStates.set(surface, state);
    return state;
  }
  return fallbackState;
}

function median(values: number[]): number {
  values.sort((a, b) => a - b);
  const middle = Math.floor(values.length / 2);
  return values.length % 2 === 0
    ? (values[middle - 1] + values[middle]) / 2
    : values[middle];
}

function measureSlope(ctx: Ctx2D): number {
  const canvas = createAuxCanvasForContext(ctx, RASTER_SIZE_PX, RASTER_SIZE_PX);
  if (canvas === null) return 0;
  const scratch = canvas.getContext('2d', { willReadFrequently: true }) as
    | CanvasRenderingContext2D
    | OffscreenCanvasRenderingContext2D
    | null;
  if (scratch === null) return 0;

  scratch.clearRect(0, 0, RASTER_SIZE_PX, RASTER_SIZE_PX);
  scratch.font = replaceFontSize(ctx.font, `${RASTER_FONT_PX}px`);
  scratch.fillStyle = '#000';
  scratch.textAlign = 'center';
  scratch.textBaseline = 'middle';
  scratch.fillText('ー', RASTER_SIZE_PX / 2, RASTER_SIZE_PX / 2);
  const data = scratch.getImageData(0, 0, RASTER_SIZE_PX, RASTER_SIZE_PX).data;
  const centroids: Array<{ x: number; y: number }> = [];
  for (let x = 0; x < RASTER_SIZE_PX; x += 1) {
    let alphaSum = 0;
    let weightedY = 0;
    for (let y = 0; y < RASTER_SIZE_PX; y += 1) {
      const alpha = data[(y * RASTER_SIZE_PX + x) * 4 + 3];
      alphaSum += alpha;
      weightedY += y * alpha;
    }
    if (alphaSum > 0) centroids.push({ x, y: weightedY / alphaSum });
  }
  if (centroids.length < 2) return 0;

  const slopes: number[] = [];
  for (let i = 0; i < centroids.length; i += 1) {
    for (let j = i + 1; j < centroids.length; j += 1) {
      slopes.push((centroids[j].y - centroids[i].y) / (centroids[j].x - centroids[i].x));
    }
  }
  return slopes.length > 0 ? median(slopes) : 0;
}

/** Only U+30FC is straightened; the two wave marks retain their designed drift. */
export function verticalFallbackShearEnabled(cp: number): boolean {
  return cp === PROLONGED_SOUND_MARK;
}

/**
 * Runtime-measured horizontal-glyph slope used by the non-`vert` mirror fallback.
 * DOM results follow the document FontFace epoch; worker/Node results are cached
 * only on the source canvas surface, so a new surface remeasures after host-side
 * font registration. Any allocation, paint, or readback failure returns and
 * caches zero, preserving the prior fallback.
 */
export function verticalFallbackShearCoefficient(ctx: Ctx2D, cp: number): number {
  if (!verticalFallbackShearEnabled(cp)) return 0;
  const state = cacheState(ctx);
  const key = `${state.epoch}:${replaceFontSize(ctx.font, '<size>')}:${cp}`;
  const cached = state.cache.get(key);
  if (cached !== undefined) return cached;
  let coefficient = 0;
  try {
    coefficient = measureSlope(ctx);
    if (!Number.isFinite(coefficient)) coefficient = 0;
  } catch {
    coefficient = 0;
  }
  state.cache.set(key, coefficient);
  return coefficient;
}
