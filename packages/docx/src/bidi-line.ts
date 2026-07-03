// Per-line bidi ordering for the docx renderer.
//
// `buildSegments` splits each run's text into space-delimited word pieces, so
// WITHIN a run Arabic joining never crosses a segment boundary (a mid-word
// run split — e.g. one letter bolded — still seams at the run boundary; that
// pre-existing limitation is tracked for Phase 4). That lets us reorder at SEGMENT
// granularity (1:1 with the laid-out segments — every per-segment property is
// preserved) using the shared UAX#9 engine, and let Canvas shape/mirror each
// segment internally when it is drawn with `ctx.direction` set to the segment's
// resolved direction. Inline objects (image / math / tab) participate as a
// single neutral object-replacement character.

import {
  getDefaultBidiEngine,
  hasStrongRtl,
  OBJECT_PLACEHOLDER,
  buildVisualOrder,
  type BidiClass,
} from '@silurus/ooxml-core';

/** A laid-out segment as seen here: only its optional text matters for bidi.
 *  Typed as `unknown` element so the renderer's LayoutSeg union (whose image /
 *  math / tab members carry no `text`) assigns cleanly. */
const segText = (s: unknown): string | undefined => {
  const t = (s as { text?: unknown }).text;
  return typeof t === 'string' ? t : undefined;
};

/** Does this segment carry a run-level `<w:rtl>` (ECMA-376 §17.3.2.30)? */
const segRtl = (s: unknown): boolean => (s as { rtl?: unknown }).rtl === true;

/** Should this segment's European digits be classified AN (Word's Arabic
 *  complex-script digit ordering)? Set by the renderer from `w:lang w:bidi`
 *  (§17.3.2.20). See {@link computeLineVisualOrder}. */
const segDigitsAsAN = (s: unknown): boolean =>
  (s as { digitsAsAN?: unknown }).digitsAsAN === true;

/**
 * Cheap gate: does this run of segments need the bidi pass? True when any
 * segment contains a strong-RTL character OR carries a run-level `<w:rtl>`
 * mark (§17.3.2.30 — e.g. a digits-only run that must resolve RTL).
 */
export function segmentsHaveRtl(segments: readonly unknown[]): boolean {
  for (const s of segments) {
    if (segRtl(s)) return true;
    const t = segText(s);
    if (t !== undefined && hasStrongRtl(t)) return true;
  }
  return false;
}

export interface LineVisualOrder {
  /** Logical segment indices in visual (left-to-right) order. */
  order: number[];
  /** Per-LOGICAL-index resolved direction (true = RTL) for `ctx.direction`. */
  rtl: boolean[];
}

/** Punctuation/symbols — the "ambiguous" characters a run-level `<w:rtl>`
 *  resolves to RTL (§17.3.2.30). Whitespace is deliberately EXCLUDED: an
 *  inter-word space classified R would reverse the mutual order of English
 *  words inside an rtl-marked run, which Word does not do. Letters/digits are
 *  excluded too — strong chars keep their class, and digits are handled by the
 *  separate `digitsAsAN` (w:lang-gated) override. */
const AMBIGUOUS_CHAR = /[\p{P}\p{S}]/u;

/**
 * Compute the visual draw order of a line's segments under `baseRtl`. Text
 * segments contribute their text; non-text segments contribute one neutral
 * placeholder so they take the surrounding direction. Each segment is assigned
 * the embedding level of its first code unit (segments are single-script in
 * practice because they are space-split); Canvas resolves any residual
 * intra-segment bidi when the slice is drawn with the matching `ctx.direction`.
 *
 * A run-level `<w:rtl>` (§17.3.2.30) gives the run's AMBIGUOUS characters
 * right-to-left characteristics. This is modelled as a UAX#9 §4.3 HL1
 * Bidi_Class override (punctuation/symbols → R), NOT as an RLE…PDF embedding:
 * an embedding raises the run's level above the paragraph base, which strands
 * sibling base-level content on the wrong side (e.g. a trailing "." run after
 * "2022" in an RTL paragraph was over-embedded to level 3 and reordered to
 * "2022." where Word renders ".2022"). With the class override the run's
 * neutrals resolve RTL at the BASE level, exactly like Word: a literal "1. "
 * prefix mirrors to ".1", while strong-Latin content keeps its even (LTR)
 * level so English words in an rtl-marked run keep their LTR word order.
 */
export function computeLineVisualOrder(
  segments: readonly unknown[],
  baseRtl: boolean,
): LineVisualOrder {
  const n = segments.length;
  if (n === 0) return { order: [], rtl: [] };

  // Concatenate every segment into one logical string for the bidi algorithm.
  let full = '';
  const segStart: number[] = new Array(n);
  const segEnd: number[] = new Array(n);
  // UAX#9 §4.3 HL1 per-code-unit Bidi_Class override (see engine.computeLevels).
  //  - `digitsAsAN` segments: European digits → AN, so a logical "28-02-2026"
  //    reorders to Word's "2026-02-28" under an RTL base (§17.3.2.20).
  //  - rtl-marked segments (§17.3.2.30): punctuation/symbols → R, so the run's
  //    ambiguous characters resolve RTL at the base level (see doc above).
  // `undefined` until any segment opts in, so the pure algorithm runs for
  // ordinary lines.
  let classOverride: (BidiClass | null)[] | undefined;
  const ensureOverride = (): (BidiClass | null)[] => {
    if (!classOverride) classOverride = [];
    while (classOverride.length < full.length) classOverride.push(null);
    return classOverride;
  };
  for (let i = 0; i < n; i++) {
    const t = segText(segments[i]) ?? '';
    segStart[i] = full.length;
    full += t.length > 0 ? t : OBJECT_PLACEHOLDER;
    segEnd[i] = full.length;

    if (t.length > 0 && (segDigitsAsAN(segments[i]) || segRtl(segments[i]))) {
      const ov = ensureOverride();
      const digitsAN = segDigitsAsAN(segments[i]);
      const rtlMarked = segRtl(segments[i]);
      for (let k = segStart[i]; k < segEnd[i]; k++) {
        const c = full.charCodeAt(k);
        if (digitsAN && c >= 0x30 && c <= 0x39) {
          ov[k] = 'AN';
        } else if (rtlMarked && AMBIGUOUS_CHAR.test(full[k])) {
          ov[k] = 'R';
        }
      }
    }
  }
  if (classOverride) while (classOverride.length < full.length) classOverride.push(null);

  const engine = getDefaultBidiEngine();
  const { levels, paragraphLevel } = engine.computeLevels(
    full,
    baseRtl ? 'rtl' : 'ltr',
    classOverride,
  );

  // Ordering: each segment takes the RESOLVED level of its first real code unit
  // (255-removed → paragraph level), then UAX#9 L2 permutes them — the shared
  // `buildVisualOrder` back half. No level forcing for rtl-marked segments:
  // Latin letters keep their even (LTR) level (so English words in an rtl-marked
  // run keep their mutual LTR order, as Word renders them), while the run's
  // punctuation resolves to the odd level via the §17.3.2.30 class override
  // above.
  const { order } = buildVisualOrder(levels, paragraphLevel, segStart);

  // Direction hint = "does the segment contain ANY odd-level unit". A
  // digits-with-punctuation slice like "1. " has its "." resolve to the odd
  // level, so the slice draws with ctx.direction rtl and Canvas mirrors it to
  // ".1" exactly as Word does — whereas a pure-Latin slice is all-even and
  // keeps its LTR rendering. This is docx-specific (pptx/xlsx use plain
  // first-unit level parity), so it stays here rather than in buildVisualOrder.
  const rtl: boolean[] = new Array(n);
  for (let i = 0; i < n; i++) {
    // Scan excludes the segment's TRAILING whitespace: an inter-word space's
    // level is seam context (N2 gives it the embedding level between
    // opposite-direction neighbours), not segment content — only the content
    // decides whether the slice must mirror.
    let scanEnd = segEnd[i];
    while (scanEnd > segStart[i] && full[scanEnd - 1] === ' ') scanEnd--;
    let anyOdd = false;
    for (let k = segStart[i]; k < scanEnd; k++) {
      const l = levels[k];
      if (l !== 255 && (l & 1) === 1) {
        anyOdd = true;
        break;
      }
    }
    rtl[i] = anyOdd;
  }

  return { order, rtl };
}
/** Physical edge a line aligns to, resolving logical start/end against base direction. */
export type AlignEdge = 'left' | 'right' | 'center' | 'justify';

/**
 * Resolve a paragraph's `w:jc` value (and base direction) to a physical edge.
 * ALL edge values are logical in WordprocessingML: `start`/`end` by definition
 * (§17.18.44), and the transitional `left`/`right` are defined as
 * "semantically equivalent to start/end" (ECMA-376 Part 4 §14.11.2) — so every
 * edge flips under an RTL base. An unset alignment defaults to the leading
 * (logical-start) edge.
 */
export function resolveAlignEdge(alignment: string | undefined, baseRtl: boolean): AlignEdge {
  switch (alignment) {
    case 'center':
      return 'center';
    case 'both':
    case 'justify':
    case 'distribute':
    // ECMA-376 §17.18.44: the three kashida settings (lowKashida / mediumKashida /
    // highKashida) and thaiDistribute are all forms of full justification between
    // both text margins — they differ only in HOW the extra space is distributed
    // (Arabic kashida elongation, Thai per-character spacing). The physical edge is
    // "justify" for all of them. NOTE: this is a MAPPING to the existing
    // inter-word/inter-character justification.
    // TODO(§17.18.44): true kashida (U+0640 tatweel) elongation — tracked in
    // https://github.com/yukiyokotani/office-open-xml-viewer/issues/724
    case 'lowKashida':
    case 'mediumKashida':
    case 'highKashida':
    case 'thaiDistribute':
      return 'justify';
    case 'end':
    case 'right':
      return baseRtl ? 'left' : 'right';
    case 'start':
    case 'left':
    case undefined:
    default:
      return baseRtl ? 'right' : 'left';
  }
}

/** ECMA-376 §17.18.44 — whether a `w:jc` value fully justifies each line by
 *  expanding inter-word (and, for distribute/thaiDistribute, inter-character)
 *  spacing. Covers `both` / `justify` / `distribute` plus the kashida and Thai
 *  variants, which this renderer maps onto the same slack-distribution kernel
 *  (see {@link resolveAlignEdge}; true kashida elongation is a follow-up). */
export function jcIsFullyJustified(alignment: string | undefined): boolean {
  switch (alignment) {
    case 'both':
    case 'justify':
    case 'distribute':
    case 'lowKashida':
    case 'mediumKashida':
    case 'highKashida':
    case 'thaiDistribute':
      return true;
    default:
      return false;
  }
}

/** ECMA-376 §17.18.44 — whether a `w:jc` value also stretches the paragraph's
 *  LAST line (unlike `both`, whose final line is left as-is). `distribute` and
 *  its Thai optimization `thaiDistribute` both spread every line, including the
 *  last. */
export function jcStretchesLastLine(alignment: string | undefined): boolean {
  return alignment === 'distribute' || alignment === 'thaiDistribute';
}
