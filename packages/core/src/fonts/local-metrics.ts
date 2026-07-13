import { retainFace, releaseFaces } from './font-registry.js';
import { activeFontSet, withFontCeiling } from './preload.js';

/** A family-level local-font measurement request. This API is shared by all
 * OOXML formats; each format decides which Office behavior requires a measured
 * multiplier and supplies the evidence-backed value. */
export interface LocalFontMetricRequest {
  /** Authored OOXML family name used as the lookup key. */
  family: string;
  /** Exact local() names, in preference order. No generic fallback is allowed. */
  localNames: readonly string[];
  /** Office single-line multiplier applied to the local face's design box. */
  lineHeightMultiplier: number;
}

export interface ResolvedLocalFontMetric {
  /** Isolated FontFace family registered for Canvas measure and paint. */
  family: string;
  /** Office single-line height divided by em. */
  lineHeightRatio: number;
}

export interface LoadedLocalFontMetrics {
  /** Refcounted faces retained by this caller; release on document destroy. */
  faces: FontFace[];
  /** Normalized authored family → exact local face and measured line ratio. */
  metrics: Record<string, ResolvedLocalFontMetric>;
}

export function normalizeLocalFontMetricFamily(family: string): string {
  return family.trim().toLowerCase();
}

function quoteLocalName(name: string): string {
  return `local("${name.replaceAll('\\', '\\\\').replaceAll('"', '\\"')}")`;
}

function stableHash(value: string): string {
  let hash = 0x811c9dc5;
  for (let i = 0; i < value.length; i++) {
    hash ^= value.charCodeAt(i);
    hash = Math.imul(hash, 0x01000193);
  }
  return (hash >>> 0).toString(16).padStart(8, '0');
}

function measureContext():
  | CanvasRenderingContext2D
  | OffscreenCanvasRenderingContext2D
  | null {
  if (typeof OffscreenCanvas !== 'undefined') {
    return new OffscreenCanvas(1, 1).getContext('2d');
  }
  if (typeof document !== 'undefined' && document?.createElement) {
    return document.createElement('canvas').getContext('2d');
  }
  return null;
}

/** Load exact local faces and measure their design line boxes without accepting
 * CSS fallback substitution. A missing local() source rejects FontFace.load(),
 * so no metric is returned and the renderer keeps its static OOXML/Office
 * profile. The isolated alias makes the measured and painted face identical.
 */
export async function loadLocalFontMetrics(
  requests: readonly LocalFontMetricRequest[],
): Promise<LoadedLocalFontMetrics> {
  const set = activeFontSet();
  const ctx = measureContext();
  if (!set || !ctx || typeof FontFace === 'undefined') return { faces: [], metrics: {} };

  const faces: FontFace[] = [];
  const metrics: Record<string, ResolvedLocalFontMetric> = {};
  for (const request of requests) {
    const family = request.family.trim();
    const localNames = request.localNames.map((name) => name.trim()).filter(Boolean);
    if (!family || localNames.length === 0 || !(request.lineHeightMultiplier > 0)) continue;
    const source = localNames.map(quoteLocalName).join(', ');
    const signature = `local-metric:${normalizeLocalFontMetricFamily(family)}:${source}`;
    const alias = `__ooxml_local_${stableHash(signature)}`;
    const { face } = retainFace(signature, set, () => {
      const created = new FontFace(alias, source);
      set.add(created);
      return created;
    });

    try {
      // A second document can retain the shared face while the first holder's
      // load is still in flight. FontFace.load() is idempotent and joins that
      // pending load, so await it for every holder before measuring; `isNew`
      // alone is not a sufficient loaded-state guarantee under concurrent opens.
      const loaded = await withFontCeiling(face.load());
      if (!loaded || face.status !== 'loaded') throw new Error('local font load timed out');
      ctx.font = `100px "${alias}"`;
      const measured = ctx.measureText('Hg国');
      const ascent = measured.fontBoundingBoxAscent;
      const descent = measured.fontBoundingBoxDescent;
      if (!(Number.isFinite(ascent) && Number.isFinite(descent) && ascent + descent > 0)) {
        throw new Error('font design metrics unavailable');
      }
      const lineHeightRatio = ((ascent + descent) / 100) * request.lineHeightMultiplier;
      faces.push(face);
      metrics[normalizeLocalFontMetricFamily(family)] = { family: alias, lineHeightRatio };
    } catch {
      releaseFaces([face]);
    }
  }
  return { faces, metrics };
}

export function unloadLocalFontMetrics(faces: Iterable<FontFace>): void {
  releaseFaces(faces);
}
