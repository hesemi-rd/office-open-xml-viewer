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
  /** ECMA-376 §17.3.2.26 rFonts@hint after style inheritance. */
  readonly fontHint?: 'default' | 'eastAsia' | 'cs';
  /** Resolved w:lang@eastAsia, normalized to lower case. */
  readonly eastAsiaLanguage?: string;
  /** fontTable w:charset for the selected eastAsia face (hex byte). */
  readonly eastAsiaFontCharset?: string;
  readonly genericFamily?: 'serif' | 'sans-serif' | 'monospace';
  readonly letterSpacingPt?: number;
  /** Resolved §17.3.2.19 w:kern state at this run size. Absent preserves the
   * measurement adapter's inherited kerning policy. */
  readonly kerning?: boolean;
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
  readonly kerning?: boolean;
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
  readonly eastAsiaFontCharsets?: Readonly<Record<string, string>>;
}

const LATIN1_EAST_ASIA = new Set([
  0x00a1, 0x00a4, 0x00a7, 0x00a8, 0x00aa, 0x00ad, 0x00af,
  0x00b0, 0x00b1, 0x00b2, 0x00b3, 0x00b4, 0x00b6, 0x00b7,
  0x00b8, 0x00b9, 0x00ba, 0x00bc, 0x00bd, 0x00be, 0x00bf, 0x00d7, 0x00f7,
]);
const LATIN1_CHINESE_EAST_ASIA = new Set([
  0x00e0, 0x00e1, 0x00e8, 0x00e9, 0x00ea, 0x00ec, 0x00ed,
  0x00f2, 0x00f3, 0x00f9, 0x00fa, 0x00fc,
]);

function scriptSlot(
  codePoint: number,
  forceComplex: boolean,
  hint: TextShapeRequest['fontHint'],
  eastAsiaLanguage: string | undefined,
  eastAsiaFontCharset: string | undefined,
): FontScriptSlot {
  const hintedEastAsia = hint === 'eastAsia';
  const chinese = eastAsiaLanguage?.split(/[-_]/, 1)[0]?.toLowerCase() === 'zh';
  const chineseCharset = /^(?:86|88)$/i.test(eastAsiaFontCharset?.trim() ?? '');
  let tableSlot: Exclude<FontScriptSlot, 'complexScript'> = 'highAnsi';
  // ECMA-376 §17.3.2.26 assigns the Hebrew/Arabic-family ranges to the ASCII
  // slot unless the run is explicitly complex-script (`w:cs` / `w:rtl`). They
  // must not fall through to highAnsi merely because their scalar is > 0x7f.
  if (codePoint <= 0x007f) tableSlot = 'ascii';
  else if (codePoint <= 0x00ff) {
    tableSlot = hintedEastAsia && (
      LATIN1_EAST_ASIA.has(codePoint)
      || (chinese && LATIN1_CHINESE_EAST_ASIA.has(codePoint))
    ) ? 'eastAsia' : 'highAnsi';
  } else if (codePoint >= 0x0100 && codePoint <= 0x02af) {
    tableSlot = hintedEastAsia && (chinese || chineseCharset) ? 'eastAsia' : 'highAnsi';
  } else if (codePoint >= 0x02b0 && codePoint <= 0x04ff) {
    tableSlot = hintedEastAsia ? 'eastAsia' : 'highAnsi';
  } else if (
    (codePoint >= 0x0590 && codePoint <= 0x07bf)
    || (codePoint >= 0xfb1d && codePoint <= 0xfdff)
    || (codePoint >= 0xfe70 && codePoint <= 0xfeff)
  ) tableSlot = 'ascii';
  else if (
    (codePoint >= 0x1100 && codePoint <= 0x11ff)
    || (codePoint >= 0x2e80 && codePoint <= 0x9fff)
    || (codePoint >= 0xa000 && codePoint <= 0xa4cf)
    || (codePoint >= 0xac00 && codePoint <= 0xd7af)
    || (codePoint >= 0xf900 && codePoint <= 0xfaff)
    || (codePoint >= 0xfe30 && codePoint <= 0xfe6f)
    || (codePoint >= 0xff00 && codePoint <= 0xffef)
    || (codePoint >= 0x20000 && codePoint <= 0x2fa1f)
  ) tableSlot = 'eastAsia';
  else if (codePoint >= 0x1e00 && codePoint <= 0x1eff) {
    tableSlot = hintedEastAsia && chinese ? 'eastAsia' : 'highAnsi';
  } else if (
    (codePoint >= 0x2000 && codePoint <= 0x27bf)
    || (codePoint >= 0xe000 && codePoint <= 0xf8ff)
    || (codePoint >= 0xfb00 && codePoint <= 0xfb1c)
  ) tableSlot = hintedEastAsia ? 'eastAsia' : 'highAnsi';

  // §17.3.2.26 step 2: an eastAsia table result is protected from w:cs/w:rtl
  // only when rFonts@hint explicitly selects eastAsia. Otherwise cs wins.
  if (tableSlot === 'eastAsia' && hintedEastAsia) return tableSlot;
  if (forceComplex) return 'complexScript';
  return tableSlot;
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
    eastAsiaFontCharsets: input.eastAsiaFontCharsets ?? {},
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
        const eastAsiaFamily = requestedFamily(request, 'eastAsia');
        const eastAsiaCharset = request.eastAsiaFontCharset
          ?? (eastAsiaFamily
            ? input.eastAsiaFontCharsets?.[eastAsiaFamily.trim().toLowerCase()]
            : undefined);
        const script = scriptSlot(
          character.codePointAt(0) ?? 0,
          request.complexScript ?? false,
          request.fontHint,
          request.eastAsiaLanguage,
          eastAsiaCharset,
        );
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
          kerning: request.kerning,
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
