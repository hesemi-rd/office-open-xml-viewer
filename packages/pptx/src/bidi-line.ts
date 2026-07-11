// Per-line bidi ordering for the pptx renderer.
//
// Lines are laid out as a list of style/word segments. We reorder at SEGMENT
// granularity (1:1 with the laid-out segments — every per-segment property is
// preserved) using the shared UAX#9 engine (rule L2), and let Canvas
// shape/mirror each segment internally when it is drawn with `ctx.direction`
// set to the segment's resolved direction. The whole segment string is drawn in
// one fillText, so Canvas resolves any residual intra-segment bidi. Inline
// objects (math / image) participate as a single neutral object-replacement
// character.

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

/** A DrawingML tab is UAX#9 Bidi_Class S, not a neutral inline object. */
const segIsTab = (s: unknown): boolean => 'isTab' in (s as object);

/** Cheap test: does this run of segments contain any strong-RTL character? */
export function segmentsHaveRtl(segments: readonly unknown[]): boolean {
  for (const s of segments) {
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

/**
 * Compute the visual draw order of a line's segments under `baseRtl`. Text
 * segments contribute their text; non-text segments contribute one neutral
 * placeholder so they take the surrounding direction. Each segment is assigned
 * the embedding level of its first code unit. pptx segments are style-merged
 * (not word-split), so a single-style segment CAN span a direction boundary;
 * its internal order is still resolved correctly by Canvas (whole segment in
 * one fillText with the matching `ctx.direction`), while its position among
 * neighbours uses the first-unit level — a documented approximation, exact
 * splitting at level boundaries is tracked as a follow-up.
 */
export function computeLineVisualOrder(
  segments: readonly unknown[],
  baseRtl: boolean,
): LineVisualOrder {
  const n = segments.length;
  if (n === 0) return { order: [], rtl: [] };

  let full = '';
  const segStart: number[] = new Array(n);
  let classOverride: (BidiClass | null)[] | undefined;
  for (let i = 0; i < n; i++) {
    segStart[i] = full.length;
    const t = segText(segments[i]) ?? '';
    full += t.length > 0 ? t : OBJECT_PLACEHOLDER;
    if (segIsTab(segments[i])) {
      classOverride ??= [];
      while (classOverride.length < full.length) classOverride.push(null);
      classOverride[segStart[i]] = 'S';
    }
  }
  if (classOverride) while (classOverride.length < full.length) classOverride.push(null);

  const { levels, paragraphLevel } = getDefaultBidiEngine().computeLevels(
    full,
    baseRtl ? 'rtl' : 'ltr',
    classOverride,
  );

  const { order, segLevels } = buildVisualOrder(levels, paragraphLevel, segStart);
  const rtl: boolean[] = new Array(n);
  for (let i = 0; i < n; i++) rtl[i] = (segLevels[i] & 1) === 1;
  return { order, rtl };
}
