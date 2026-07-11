/**
 * Horizontal advance corrections for document fonts whose real metrics differ
 * materially from the browser fallback used when the requested face is absent.
 *
 * Values in this module describe fonts, not samples. A family belongs here only
 * when its source metric has a documented provenance; unknown families remain a
 * strict 1/0 no-op so existing layout is byte-stable.
 */

export type FontAdvanceScriptClass =
  | 'ideograph'
  | 'hiragana'
  | 'katakana'
  | 'fullwidthPunctuation'
  | 'fullwidthLatin'
  | 'other';

export interface FontAdvanceRun {
  readonly text: string;
  readonly scale: number;
}

interface FontAdvanceProfile {
  readonly family: string;
  readonly scale: Readonly<Record<FontAdvanceScriptClass, number>>;
}

const NO_ADVANCE_CORRECTION: Readonly<Record<FontAdvanceScriptClass, number>> = {
  ideograph: 1,
  hiragana: 1,
  katakana: 1,
  fullwidthPunctuation: 1,
  fullwidthLatin: 1,
  other: 1,
};

/**
 * Meiryo UI Regular horizontal advances harvested from `meiryo.ttc` with
 * fontTools (`hmtx`, unitsPerEm 2048). The values are class means over the
 * representative repertoire used for the harvest:
 *
 * - ideographs and fullwidth Latin: 1.0000 em;
 * - Hiragana: 0.7775 em (individual glyphs 0.50–0.89 em);
 * - Katakana: 0.7438 em (individual glyphs 0.69–0.82 em);
 * - fullwidth punctuation `、。「」（）`: 0.7214 em.
 *
 * Plain Meiryo is deliberately absent: its corresponding glyphs are uniformly
 * full-width. These class means correct a full-width substitute toward the real
 * requested face; they are not visual-regression tuning constants.
 */
const FONT_ADVANCE_PROFILES: ReadonlyArray<FontAdvanceProfile> = [
  {
    family: 'meiryo ui',
    scale: {
      ideograph: 1,
      hiragana: 0.7775,
      katakana: 0.7438,
      fullwidthPunctuation: 0.7214,
      fullwidthLatin: 1,
      other: 1,
    },
  },
  {
    family: 'メイリオ ui',
    scale: {
      ideograph: 1,
      hiragana: 0.7775,
      katakana: 0.7438,
      fullwidthPunctuation: 0.7214,
      fullwidthLatin: 1,
      other: 1,
    },
  },
];

interface FontBiasProfile {
  readonly test: (family: string) => boolean;
  readonly biasEm: number;
}

/**
 * Canvas-vs-Word horizontal advance bias in em per glyph. This is independent
 * of the real-font script profiles above: it is a line-fit allowance for the
 * gap between Canvas `measureText` advances and Word's own layout advances for
 * the SAME face, not a glyph transform. The allowance is backend-agnostic by
 * design; the committed Georgia value below is calibrated on the Chromium VRT.
 *
 * Georgia: the tracked public demo's justified body face (issue #794). The
 * The Chromium VRT's Canvas-vs-Word accumulated excess measures at roughly
 * 0.1–0.3 px per glyph
 * at the demo's 10–11 px body em — an em-fraction band of ~0.009–0.028. The
 * committed value is fixed INSIDE that measured band by the public demo's
 * Word-reference wrap positions (demo/sample-1 fidelity ratchet, scanned at
 * 0.009/0.0105/0.0115/0.012/0.013/0.02): 0.009 under-admits words Word keeps
 * (pages 4–5 regress) and 0.012+ over-admits words Word wraps (pages 3/5);
 * 0.0105 reproduces every Word-verified wrap. Times New Roman (measures
 * ~0.03 px/glyph, effectively zero), CSS generics, and unknown families
 * intentionally fall through to zero.
 */
const FONT_BIAS_PROFILES: ReadonlyArray<FontBiasProfile> = [
  {
    test: (family) => family === 'georgia',
    biasEm: 0.0105,
  },
];

function normalizeFamily(family: string | null | undefined): string {
  return (family ?? '')
    .trim()
    .replace(/^["']|["']$/g, '')
    .replace(/\s+/g, ' ')
    .toLowerCase();
}

function advanceProfile(family: string | null | undefined): FontAdvanceProfile | null {
  const normalized = normalizeFamily(family);
  return FONT_ADVANCE_PROFILES.find((profile) => profile.family === normalized) ?? null;
}

/** Classify one Unicode code point for a horizontal font-advance profile. */
export function fontAdvanceScriptClass(char: string): FontAdvanceScriptClass {
  const cp = char.codePointAt(0);
  if (cp === undefined) return 'other';
  if (cp >= 0x3040 && cp <= 0x309f) return 'hiragana';
  if (cp >= 0x30a0 && cp <= 0x30ff) return 'katakana';
  if ('、。「」（）'.includes(char)) return 'fullwidthPunctuation';
  if (
    (cp >= 0x3400 && cp <= 0x4dbf)
    || (cp >= 0x4e00 && cp <= 0x9fff)
    || (cp >= 0xf900 && cp <= 0xfaff)
    || (cp >= 0x20000 && cp <= 0x323af)
  ) return 'ideograph';
  if ((cp >= 0xff21 && cp <= 0xff3a) || (cp >= 0xff41 && cp <= 0xff5a)) {
    return 'fullwidthLatin';
  }
  return 'other';
}

/** Real requested-face advance scale for a character, or 1 for a no-op. */
export function fontScriptAdvanceScale(
  family: string | null | undefined,
  char: string,
): number {
  const profile = advanceProfile(family);
  if (profile === null) return 1;
  return profile.scale[fontAdvanceScriptClass(char)] ?? NO_ADVANCE_CORRECTION.other;
}

/**
 * Split text into maximal runs with one uniform requested-font advance scale.
 * Untabled families are returned untouched as one no-op run.
 */
export function splitFontAdvanceRuns(
  family: string | null | undefined,
  text: string,
): FontAdvanceRun[] {
  const profile = advanceProfile(family);
  if (profile === null || text.length === 0) return [{ text, scale: 1 }];
  const runs: FontAdvanceRun[] = [];
  let currentText = '';
  let currentScale: number | null = null;
  for (const char of text) {
    const scale = profile.scale[fontAdvanceScriptClass(char)];
    if (currentScale === null || scale === currentScale) {
      currentText += char;
      currentScale = scale;
    } else {
      runs.push({ text: currentText, scale: currentScale });
      currentText = char;
      currentScale = scale;
    }
  }
  if (currentText.length > 0) runs.push({ text: currentText, scale: currentScale ?? 1 });
  return runs;
}

/** Expected Canvas-over-Word advance bias, in em per glyph. */
export function fontAdvanceBiasEm(family: string | null | undefined): number {
  const normalized = normalizeFamily(family);
  for (const profile of FONT_BIAS_PROFILES) {
    if (profile.test(normalized)) return profile.biasEm;
  }
  return 0;
}
