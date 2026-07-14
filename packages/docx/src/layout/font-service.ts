import type { LayoutDiagnostic } from './types.js';
import { stableFingerprint } from './fingerprint.js';

export type FontResolutionSource = 'embedded' | 'local' | 'google' | 'substitute' | 'generic';
export type FontStyle = 'normal' | 'italic';

export interface FontRequest {
  readonly requestedFamily?: string | null;
  readonly genericFamily?: 'serif' | 'sans-serif' | 'monospace';
  readonly weight?: number;
  readonly style?: FontStyle;
}

export interface FontResolution {
  readonly requestedFamily: string;
  readonly resolvedFamily: string;
  readonly source: FontResolutionSource;
  readonly weight: number;
  readonly style: FontStyle;
  readonly diagnostics: readonly LayoutDiagnostic[];
  readonly genericFamily: 'serif' | 'sans-serif' | 'monospace';
}

export interface FontResolver {
  readonly fingerprint: string;
  resolve(request: Readonly<FontRequest>): FontResolution;
}

export interface FontInventoryFace {
  readonly requestedFamily: string;
  readonly resolvedFamily: string;
  readonly source: Exclude<FontResolutionSource, 'generic'>;
  readonly weights?: readonly number[];
  readonly styles?: readonly FontStyle[];
}

function normalizeFamily(value: string): string {
  return value.trim().toLocaleLowerCase('en-US');
}

function normalizedWeight(value: number | undefined): number {
  if (value == null || !Number.isFinite(value)) return 400;
  return Math.min(900, Math.max(100, Math.round(value / 100) * 100));
}

function freezeResolution(value: FontResolution): FontResolution {
  return Object.freeze({ ...value, diagnostics: Object.freeze([...value.diagnostics]) });
}

/**
 * Snapshot the font inventory used by one document. ECMA-376 §17.8.2 leaves
 * the substitution algorithm implementation-defined, so a substituted or
 * generic result is carried as an explicit diagnostic instead of being hidden
 * in paragraph geometry.
 */
export function createFontResolver(inventory: readonly FontInventoryFace[]): FontResolver {
  const sourcePriority: Readonly<Record<Exclude<FontResolutionSource, 'generic'>, number>> = {
    embedded: 0,
    local: 1,
    google: 2,
    substitute: 3,
  };
  const faces = inventory
    .filter((face) => face.requestedFamily.trim() && face.resolvedFamily.trim())
    .map((face) => Object.freeze({
      ...face,
      weights: Object.freeze([...(face.weights ?? [400])]),
      styles: Object.freeze([...(face.styles ?? ['normal'])]),
    }))
    .sort((a, b) => {
      const family = normalizeFamily(a.requestedFamily).localeCompare(normalizeFamily(b.requestedFamily));
      return family || sourcePriority[a.source] - sourcePriority[b.source]
        || a.resolvedFamily.localeCompare(b.resolvedFamily);
    });
  const byFamily = new Map<string, (typeof faces)[number][]>();
  for (const face of faces) {
    const key = normalizeFamily(face.requestedFamily);
    byFamily.set(key, [...(byFamily.get(key) ?? []), face]);
  }
  const fingerprint = stableFingerprint('fonts', faces);

  return Object.freeze({
    fingerprint,
    resolve(request: Readonly<FontRequest>): FontResolution {
      const requestedFamily = request.requestedFamily?.trim() || request.genericFamily || 'sans-serif';
      const weight = normalizedWeight(request.weight);
      const style = request.style ?? 'normal';
      const candidates = byFamily.get(normalizeFamily(requestedFamily)) ?? [];
      const face = candidates.find((candidate) =>
        candidate.weights.includes(weight) && candidate.styles.includes(style),
      );
      if (face) {
        const diagnostics: LayoutDiagnostic[] = face.source === 'substitute'
          ? [{
              code: 'UNSUPPORTED_FEATURE',
              severity: 'warning',
              message: `ECMA-376 §17.8.2 implementation-dependent font substitution: ${requestedFamily} resolved to ${face.resolvedFamily}`,
            }]
          : [];
        return freezeResolution({
          requestedFamily,
          resolvedFamily: face.resolvedFamily,
          source: face.source,
          weight,
          style,
          diagnostics,
          genericFamily: request.genericFamily ?? 'sans-serif',
        });
      }

      const generic = request.genericFamily ?? 'sans-serif';
      return freezeResolution({
        requestedFamily,
        // Never leave an uninventoried authored family in Canvas's font list:
        // that would silently pick a host-local face outside this snapshot.
        resolvedFamily: generic,
        source: 'generic',
        weight,
        style,
        diagnostics: [{
          code: 'UNSUPPORTED_FEATURE',
          severity: 'warning',
          message: `Font ${requestedFamily} is unavailable; using generic ${generic}`,
        }],
        genericFamily: generic,
      });
    },
  });
}
