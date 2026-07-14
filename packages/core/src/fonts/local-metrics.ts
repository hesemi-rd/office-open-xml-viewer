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
  /** Evidence-backed Office single-line multiplier applied to the local face's
   * design box. Omit when the face is loaded only as an exact geometry alias. */
  lineHeightMultiplier?: number;
  /** Canvas request tuple mapped to the isolated regular local face. CSS Fonts
   * local() selects by full/PostScript face name, not these descriptors, so
   * non-normal tuples deliberately use Canvas synthesis from the loaded face. */
  weight?: number;
  style?: 'normal' | 'italic';
  /** Document-used strings measured transiently into the geometry hash. The
   * strings themselves are never retained in the resolved snapshot. */
  geometryProbeTexts?: readonly string[];
}

export interface ResolvedLocalFontMetric {
  /** Isolated FontFace family registered for Canvas measure and paint. */
  family: string;
  /** Office single-line height divided by em, only for documented profiles. */
  lineHeightRatio?: number;
  /** Authored Canvas tuple isolated behind this alias. */
  requestedFamily?: string;
  weight?: number;
  style?: 'normal' | 'italic';
  /** Measured tuple signature included in document layout fingerprints. */
  geometrySignature?: string;
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

function collisionFreeAlias(signature: string): string {
  // Fixed-width Unicode-scalar encoding is injective and CSS-family safe. A
  // short digest is not sufficient here: two sources sharing an alias would
  // make Canvas face selection ambiguous even if registry keys stayed distinct.
  return `__ooxml_local_${[...signature]
    .map((scalar) => (scalar.codePointAt(0) ?? 0).toString(16).padStart(6, '0'))
    .join('')}`;
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

/** Load positively identified local faces without accepting CSS fallback.
 * CSS Fonts local() identifies a single face by full/PostScript name; it cannot
 * enumerate an installed family's bold/italic faces. We therefore register one
 * isolated regular alias and intentionally let Canvas synthesize every authored
 * weight/style tuple from that exact loaded face. Each synthesized tuple is
 * measured into `geometrySignature`, so capability, geometry, diagnostics, and
 * fingerprints describe the same face used for paint. A missing local() source
 * rejects FontFace.load() and contributes no inventory entry.
 */
export async function loadLocalFontMetrics(
  requests: readonly LocalFontMetricRequest[],
): Promise<LoadedLocalFontMetrics> {
  const set = activeFontSet();
  const ctx = measureContext();
  if (!set || !ctx || typeof FontFace === 'undefined') return { faces: [], metrics: {} };

  const faces: FontFace[] = [];
  const metrics: Record<string, ResolvedLocalFontMetric> = {};
  type PreparedRequest = LocalFontMetricRequest & {
    family: string;
    normalizedFamily: string;
    source: string;
    weight: number;
    style: 'normal' | 'italic';
  };
  const grouped = new Map<string, { source: string; requests: PreparedRequest[] }>();
  for (const request of requests) {
    const family = request.family.trim();
    const localNames = request.localNames.map((name) => name.trim()).filter(Boolean);
    if (!family || localNames.length === 0) continue;
    if (request.lineHeightMultiplier != null && !(request.lineHeightMultiplier > 0)) continue;
    const weight = request.weight ?? 400;
    const style = request.style ?? 'normal';
    if (!(weight >= 100 && weight <= 900) || (style !== 'normal' && style !== 'italic')) continue;
    const source = localNames.map(quoteLocalName).join(', ');
    const normalizedFamily = normalizeLocalFontMetricFamily(family);
    const signature = `local-face:${source}`;
    const group = grouped.get(signature) ?? { source, requests: [] };
    group.requests.push({ ...request, family, normalizedFamily, source, weight, style });
    grouped.set(signature, group);
  }

  for (const [signature, group] of grouped) {
    const alias = collisionFreeAlias(signature);
    const { face } = retainFace(signature, set, () => {
      const created = new FontFace(alias, group.source);
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
      let hasMetrics = false;
      for (const request of group.requests) {
        ctx.font = `${request.style} ${request.weight} 100px "${alias}"`;
        const geometry: Array<number | null> = [];
        let geometryAvailable = true;
        for (const probe of ['Aa0', '国', 'ش', ...(request.geometryProbeTexts ?? [])]) {
          const measured = ctx.measureText(probe);
          if (!Number.isFinite(measured.width)) {
            geometryAvailable = false;
            break;
          }
          geometry.push(
            measured.width,
            Number.isFinite(measured.actualBoundingBoxAscent) ? measured.actualBoundingBoxAscent : null,
            Number.isFinite(measured.actualBoundingBoxDescent) ? measured.actualBoundingBoxDescent : null,
            Number.isFinite(measured.fontBoundingBoxAscent) ? measured.fontBoundingBoxAscent : null,
            Number.isFinite(measured.fontBoundingBoxDescent) ? measured.fontBoundingBoxDescent : null,
          );
        }
        if (!geometryAvailable) continue;
        let lineHeightRatio: number | undefined;
        if (request.lineHeightMultiplier != null) {
          const measured = ctx.measureText('Hg国');
          const ascent = measured.fontBoundingBoxAscent;
          const descent = measured.fontBoundingBoxDescent;
          if (!(Number.isFinite(ascent) && Number.isFinite(descent) && ascent + descent > 0)) {
            continue;
          }
          lineHeightRatio = ((ascent + descent) / 100) * request.lineHeightMultiplier;
        }
        const key = request.weight === 400 && request.style === 'normal'
          ? request.normalizedFamily
          : `${request.normalizedFamily}:${request.weight}:${request.style}`;
        metrics[key] = {
          family: alias,
          ...(lineHeightRatio == null ? {} : { lineHeightRatio }),
          requestedFamily: request.family,
          weight: request.weight,
          style: request.style,
          // Collision-safe canonical numeric serialization; document text is
          // absent, and service identity never depends on a short hash.
          geometrySignature: JSON.stringify(geometry),
        };
        hasMetrics = true;
      }
      if (!hasMetrics) throw new Error('font geometry unavailable');
      faces.push(face);
    } catch {
      releaseFaces([face]);
    }
  }
  return { faces, metrics };
}

export function unloadLocalFontMetrics(faces: Iterable<FontFace>): void {
  releaseFaces(faces);
}
