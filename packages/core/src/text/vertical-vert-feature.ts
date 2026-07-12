type Ctx2D = CanvasRenderingContext2D | OffscreenCanvasRenderingContext2D;

interface ProbeState {
  canvas: HTMLCanvasElement;
  cache: Map<string, boolean>;
  epoch: number;
}

const PROBE_SIZE_PX = 256;
const PROBE_FONT_PX = 200;
const probeStates = new WeakMap<Document, ProbeState>();

interface RasterSnapshot {
  alpha: Uint8ClampedArray;
  geometry: number[];
  metrics: number[];
}

export interface VerticalGlyphCellMetrics {
  advancePx: number;
  inkBeforePx: number;
  inkAfterPx: number;
  cellAdvancePx: number;
  originInCellPx: number;
}

function canvasElementFor(ctx: Ctx2D): HTMLCanvasElement | null {
  if (typeof HTMLCanvasElement === 'undefined') return null;
  return ctx.canvas instanceof HTMLCanvasElement ? ctx.canvas : null;
}

function replaceFontSize(font: string, replacement: string): string {
  return font.replace(
    /(^|\s)\d*\.?\d+(?:px|pt|pc|in|cm|mm|q|em|rem|%)(?:\/[^\s]+)?(?=\s)/i,
    `$1${replacement}`,
  );
}

function composeVertFeature(featureSettings: string): string {
  const normalized = featureSettings.trim();
  return normalized === '' || normalized.toLowerCase() === 'normal'
    ? '"vert" 1'
    : `${featureSettings}, "vert" 1`;
}

function probeState(doc: Document): ProbeState {
  const existing = probeStates.get(doc);
  if (existing) return existing;

  const canvas = doc.createElement('canvas');
  canvas.width = PROBE_SIZE_PX;
  canvas.height = PROBE_SIZE_PX;
  canvas.setAttribute('aria-hidden', 'true');
  Object.assign(canvas.style, {
    position: 'fixed',
    left: '-99999px',
    top: '0',
    opacity: '0',
    pointerEvents: 'none',
  });
  const state: ProbeState = { canvas, cache: new Map(), epoch: 0 };
  const invalidate = () => {
    state.epoch += 1;
    state.cache.clear();
  };
  doc.fonts?.addEventListener?.('loadingdone', invalidate);
  doc.fonts?.addEventListener?.('loadingerror', invalidate);
  probeStates.set(doc, state);
  return state;
}

function rasterSnapshot(ctx: CanvasRenderingContext2D, text: string): RasterSnapshot | null {
  const { width, height } = ctx.canvas;
  const data = ctx.getImageData(0, 0, width, height).data;
  const alpha = new Uint8ClampedArray(width * height);
  let minX = width;
  let minY = height;
  let maxX = -1;
  let maxY = -1;
  let alphaSum = 0;
  let weightedX = 0;
  let weightedY = 0;
  for (let y = 0; y < height; y += 1) {
    for (let x = 0; x < width; x += 1) {
      const a = data[(y * width + x) * 4 + 3];
      alpha[y * width + x] = a;
      if (a === 0) continue;
      minX = Math.min(minX, x);
      minY = Math.min(minY, y);
      maxX = Math.max(maxX, x);
      maxY = Math.max(maxY, y);
      alphaSum += a;
      weightedX += x * a;
      weightedY += y * a;
    }
  }
  if (maxX < minX || maxY < minY || alphaSum === 0) return null;
  const m = ctx.measureText(text);
  const finite = (value: number | undefined): number =>
    typeof value === 'number' && Number.isFinite(value) ? value : 0;
  return {
    alpha,
    geometry: [minX, minY, maxX, maxY, weightedX / alphaSum, weightedY / alphaSum],
    metrics: [
      finite(m.width),
      finite(m.actualBoundingBoxLeft),
      finite(m.actualBoundingBoxRight),
      finite(m.actualBoundingBoxAscent),
      finite(m.actualBoundingBoxDescent),
    ],
  };
}

function rasterizeProbe(
  ctx: CanvasRenderingContext2D,
  cp: number,
  featureSettings: string,
): RasterSnapshot | null {
  const canvas = ctx.canvas;
  canvas.style.fontFeatureSettings = featureSettings;
  ctx.font = ctx.font;
  ctx.clearRect(0, 0, canvas.width, canvas.height);
  const text = String.fromCodePoint(cp);
  ctx.fillText(text, canvas.width / 2, canvas.height / 2);
  return rasterSnapshot(ctx, text);
}

function arrayDifference(a: ArrayLike<number>, b: ArrayLike<number>): number {
  let difference = 0;
  const length = Math.min(a.length, b.length);
  for (let i = 0; i < length; i += 1) difference += Math.abs(a[i] - b[i]);
  return difference;
}

function rasterChanged(
  plain: RasterSnapshot,
  plainRepeat: RasterSnapshot,
  featured: RasterSnapshot,
): boolean {
  const rasterNoise = arrayDifference(plain.alpha, plainRepeat.alpha);
  const rasterSignal = arrayDifference(plain.alpha, featured.alpha);
  const geometryNoise = arrayDifference(plain.geometry, plainRepeat.geometry);
  const geometrySignal = arrayDifference(plain.geometry, featured.geometry);
  const metricNoise = arrayDifference(plain.metrics, plainRepeat.metrics);
  const metricSignal = arrayDifference(plain.metrics, featured.metrics);
  return rasterSignal > rasterNoise || geometrySignal > geometryNoise || metricSignal > metricNoise;
}

/**
 * Whether this main-thread canvas/font demonstrably paints a different glyph or
 * placement for `cp` after composing the OpenType `vert` feature. Results are
 * cached by code point, feature settings, and family/weight/style shorthand, and
 * invalidated whenever the document FontFaceSet completes or fails a load.
 */
export function verticalVertGlyphReachable(ctx: Ctx2D, cp: number): boolean {
  const target = canvasElementFor(ctx);
  if (target === null || typeof document === 'undefined') return false;

  const doc = target.ownerDocument ?? document;
  const state = probeState(doc);
  const sourceFeature = target.style.fontFeatureSettings;
  const key = `${state.epoch}:${replaceFontSize(ctx.font, '<size>')}:${sourceFeature}:${cp}`;
  const cached = state.cache.get(key);
  if (cached !== undefined) return cached;

  let supported = false;
  const parent = doc.body ?? doc.documentElement;
  if (!parent) return false;
  const wasAttached = state.canvas.isConnected;
  if (!wasAttached) parent.appendChild(state.canvas);
  const probe = state.canvas.getContext('2d', { willReadFrequently: true });
  if (probe !== null) {
    const previousFeature = state.canvas.style.fontFeatureSettings;
    try {
      probe.font = replaceFontSize(ctx.font, `${PROBE_FONT_PX}px`);
      probe.fillStyle = '#000';
      probe.textAlign = 'center';
      probe.textBaseline = 'middle';
      const plain = rasterizeProbe(probe, cp, sourceFeature);
      const plainRepeat = rasterizeProbe(probe, cp, sourceFeature);
      const vert = rasterizeProbe(probe, cp, composeVertFeature(sourceFeature));
      supported =
        plain !== null &&
        plainRepeat !== null &&
        vert !== null &&
        rasterChanged(plain, plainRepeat, vert);
    } catch {
      supported = false;
    } finally {
      state.canvas.style.fontFeatureSettings = previousFeature;
      probe.font = probe.font;
      probe.clearRect(0, 0, state.canvas.width, state.canvas.height);
      if (!wasAttached) state.canvas.remove();
    }
  }
  state.cache.set(key, supported);
  return supported;
}

/** Compose `vert` onto the element features, refont, draw, then restore exactly. */
export function withVertFeature<T>(ctx: Ctx2D, draw: () => T): T {
  const canvas = ctx.canvas as HTMLCanvasElement;
  const style = canvas?.style;
  if (!style) return draw();

  const previous = style.fontFeatureSettings;
  style.fontFeatureSettings = composeVertFeature(previous);
  ctx.font = ctx.font;
  try {
    return draw();
  } finally {
    style.fontFeatureSettings = previous;
    ctx.font = ctx.font;
  }
}

/**
 * Measure one glyph under the exact composed `vert` feature state used by paint.
 * The feature advance defines the vertical cell and its origin stays at the
 * nominal half-advance. A/D are reported for diagnostics, but designed ink pokes
 * may cross a neighbour cell and must not displace the font's glyph origin.
 */
export function measureVerticalVertGlyph(ctx: Ctx2D, ch: string): VerticalGlyphCellMetrics {
  return withVertFeature(ctx, () => {
    const previousAlign = ctx.textAlign;
    const previousBaseline = ctx.textBaseline;
    ctx.textAlign = 'center';
    ctx.textBaseline = 'middle';
    try {
      const m = ctx.measureText(ch);
      const advancePx = Number.isFinite(m.width) ? Math.max(0, m.width) : 0;
      const hasInkMetrics =
        typeof m.actualBoundingBoxAscent === 'number' &&
        Number.isFinite(m.actualBoundingBoxAscent) &&
        typeof m.actualBoundingBoxDescent === 'number' &&
        Number.isFinite(m.actualBoundingBoxDescent);
      const inkBeforePx = hasInkMetrics ? m.actualBoundingBoxAscent : 0;
      const inkAfterPx = hasInkMetrics ? m.actualBoundingBoxDescent : 0;
      return {
        advancePx,
        inkBeforePx,
        inkAfterPx,
        cellAdvancePx: advancePx,
        originInCellPx: advancePx / 2,
      };
    } finally {
      ctx.textAlign = previousAlign;
      ctx.textBaseline = previousBaseline;
    }
  });
}
