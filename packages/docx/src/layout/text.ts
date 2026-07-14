import type { LayoutDiagnostic } from './types.js';
import { graphemeClusterOffsets, type ResolvedLocalFontMetric } from '@silurus/ooxml-core';
import type {
  FontResolution,
  FontResolver,
  FontStyle,
} from './font-service.js';
import { stableFingerprint } from './fingerprint.js';

export type FontScriptSlot = 'ascii' | 'highAnsi' | 'eastAsia' | 'complexScript';

export interface TextFontSlots {
  readonly ascii?: string | null;
  readonly highAnsi?: string | null;
  readonly eastAsia?: string | null;
  readonly complexScript?: string | null;
}

export interface TextShapeRequest {
  readonly text: string;
  readonly fontSizePt: number;
  readonly fonts: TextFontSlots;
  readonly themeFonts?: TextFontSlots;
  readonly weight?: number;
  readonly style?: FontStyle;
  readonly complexScript?: boolean;
  readonly genericFamily?: 'serif' | 'sans-serif' | 'monospace';
  readonly letterSpacingPt?: number;
  /** Resolve script slots and faces without touching the measurement adapter. */
  readonly measure?: boolean;
}

export interface GlyphMeasureRequest {
  readonly text: string;
  readonly resolvedFamily: string;
  readonly fontSizePt: number;
  readonly weight: number;
  readonly style: FontStyle;
  readonly letterSpacingPt: number;
  readonly genericFamily: 'serif' | 'sans-serif' | 'monospace';
}

export interface GlyphMeasurement {
  readonly advancePt: number;
  readonly ascentPt: number;
  readonly descentPt: number;
}

export interface GlyphMeasurer {
  readonly fingerprint: string;
  measure(request: Readonly<GlyphMeasureRequest>): GlyphMeasurement;
}

export interface TextShapeSpan extends GlyphMeasurement {
  readonly text: string;
  readonly start: number;
  readonly end: number;
  readonly script: FontScriptSlot;
  readonly font: FontResolution;
}

export interface TextShapeResult extends GlyphMeasurement {
  readonly spans: readonly TextShapeSpan[];
  readonly diagnostics: readonly LayoutDiagnostic[];
}

export interface TextLayoutService {
  readonly fingerprint: string;
  readonly localMetrics: Readonly<Record<string, Readonly<ResolvedLocalFontMetric>>>;
  shape(request: Readonly<TextShapeRequest>): TextShapeResult;
}

export interface TextLayoutServiceInput {
  readonly fonts: FontResolver;
  readonly measurer: GlyphMeasurer;
  readonly localMetrics?: Readonly<Record<string, Readonly<ResolvedLocalFontMetric>>>;
}

function scriptSlot(codePoint: number, forceComplex: boolean): FontScriptSlot {
  if (forceComplex) return 'complexScript';
  if (
    (codePoint >= 0x2e80 && codePoint <= 0x9fff)
    || (codePoint >= 0xac00 && codePoint <= 0xd7af)
    || (codePoint >= 0xf900 && codePoint <= 0xfaff)
    || (codePoint >= 0x3040 && codePoint <= 0x30ff)
    || (codePoint >= 0xff00 && codePoint <= 0xffef)
    || (codePoint >= 0x20000 && codePoint <= 0x2fa1f)
  ) return 'eastAsia';
  return codePoint <= 0x7f ? 'ascii' : 'highAnsi';
}

function requestedFamily(request: Readonly<TextShapeRequest>, slot: FontScriptSlot): string | null | undefined {
  return request.fonts[slot]
    ?? request.themeFonts?.[slot]
    ?? request.fonts.ascii
    ?? request.themeFonts?.ascii;
}

/**
 * Shape per script span because ECMA-376 §17.3.2.26 selects rFonts slots per
 * Unicode character; choosing one family for an entire mixed-script run loses
 * authored East Asian and complex-script faces.
 */
export function createTextLayoutService(input: TextLayoutServiceInput): TextLayoutService {
  const localMetrics = Object.freeze(Object.fromEntries(
    Object.entries(input.localMetrics ?? {})
      .sort(([a], [b]) => a.localeCompare(b))
      .map(([family, metric]) => [family, Object.freeze({ ...metric })]),
  ));
  const fingerprint = stableFingerprint('text', {
    fonts: input.fonts.fingerprint,
    measurer: input.measurer.fingerprint,
    localMetrics,
  });
  return Object.freeze({
    fingerprint,
    localMetrics,
    shape(request: Readonly<TextShapeRequest>): TextShapeResult {
      if (!Number.isFinite(request.fontSizePt) || request.fontSizePt < 0) {
        throw new RangeError('fontSizePt must be a finite non-negative number');
      }
      const grouped: { text: string; start: number; end: number; script: FontScriptSlot }[] = [];
      const boundaries = [0, ...graphemeClusterOffsets(request.text), request.text.length]
        .filter((offset, index, values) => index === 0 || offset !== values[index - 1]);
      for (let index = 0; index < boundaries.length - 1; index += 1) {
        const start = boundaries[index];
        const end = boundaries[index + 1];
        const character = request.text.slice(start, end);
        const script = scriptSlot(character.codePointAt(0) ?? 0, request.complexScript ?? false);
        const previous = grouped.at(-1);
        if (previous?.script === script) {
          previous.text += character;
          previous.end = end;
        } else {
          grouped.push({ text: character, start, end, script });
        }
      }

      const spans = grouped.map((group): TextShapeSpan => {
        const font = input.fonts.resolve({
          requestedFamily: requestedFamily(request, group.script),
          genericFamily: request.genericFamily,
          weight: request.weight,
          style: request.style,
        });
        const measurement = request.measure === false ? {
          advancePt: 0,
          ascentPt: 0,
          descentPt: 0,
        } : input.measurer.measure({
          text: group.text,
          resolvedFamily: font.resolvedFamily,
          fontSizePt: request.fontSizePt,
          weight: font.weight,
          style: font.style,
          letterSpacingPt: request.letterSpacingPt ?? 0,
          genericFamily: font.genericFamily,
        });
        return Object.freeze({ ...group, ...measurement, font });
      });
      const diagnostics = spans.flatMap((span) => span.font.diagnostics);
      return Object.freeze({
        advancePt: spans.reduce((sum, span) => sum + span.advancePt, 0),
        ascentPt: Math.max(0, ...spans.map((span) => span.ascentPt)),
        descentPt: Math.max(0, ...spans.map((span) => span.descentPt)),
        spans: Object.freeze(spans),
        diagnostics: Object.freeze(diagnostics),
      });
    },
  });
}
