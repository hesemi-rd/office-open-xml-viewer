// UAX #14 — Unicode Line Breaking Algorithm. The generated table supplies the
// effective Line_Break property after the default LB1 resolutions; this module
// is the small typed surface shared by the OOXML renderers.

import {
  LB_CLASS_NAMES,
  LB_RANGE_STARTS,
  LB_RANGE_CLASS,
  EAW_FWH_STARTS,
  EAW_FWH_ENDS,
} from './line-break-class.generated.js';

export { LINE_BREAK_UNICODE_VERSION } from './line-break-class.generated.js';

export type LBClass = (typeof LB_CLASS_NAMES)[number];

/**
 * Effective UAX #14 Line_Break class after the default LB1 resolutions.
 *
 * @param cp An integer Unicode code point in [0, 0x10FFFF].
 */
export function lineBreakClass(cp: number): LBClass {
  let lo = 0;
  let hi = LB_RANGE_STARTS.length - 1;
  while (lo < hi) {
    const mid = (lo + hi + 1) >> 1;
    if (LB_RANGE_STARTS[mid] <= cp) lo = mid;
    else hi = mid - 1;
  }
  return LB_CLASS_NAMES[LB_RANGE_CLASS[lo]];
}

/**
 * East_Asian_Width ∈ {Fullwidth, Wide, Halfwidth} (UAX #11). This is exactly
 * the `$EastAsian` set that UAX #14 LB30 excludes from its OP/CP operands.
 */
export function isEastAsianFWH(cp: number): boolean {
  let lo = 0;
  let hi = EAW_FWH_STARTS.length - 1;
  if (hi < 0 || cp < EAW_FWH_STARTS[0]) return false;
  while (lo < hi) {
    const mid = (lo + hi + 1) >> 1;
    if (EAW_FWH_STARTS[mid] <= cp) lo = mid;
    else hi = mid - 1;
  }
  return cp < EAW_FWH_ENDS[lo];
}

/**
 * A conservative, one-way UAX #14 no-break predicate for segment boundaries.
 *
 * `true` proves that a supported rule prohibits the boundary. `false` means
 * only "not proved here" and MUST NOT be interpreted as permission to break.
 *
 * Supported rules (Unicode 17.0.0 numbering), restricted to what one directly
 * adjacent code-point pair can prove — no earlier rule in the UAX #14 chain can
 * force a break at any of these boundaries, so each pair below is sound in
 * every context:
 *
 * - LB14  `OP SP* ×` — nothing breaks directly after an opening bracket. Only
 *   the zero-space instance is decidable from a pair; seams that carry the
 *   interior spaces stay deferred.
 * - LB23  `(AL | HL) × NU`, `NU × (AL | HL)`
 * - LB23a `PR × (ID | EB | EM)`, `(ID | EB | EM) × PO`
 * - LB24  `(PR | PO) × (AL | HL)`, `(AL | HL) × (PR | PO)`
 * - LB25  zero-repetition instances of the number regex: `NU × (NU | PO | PR)`,
 *   `(PO | PR | HY | IS) × NU`. The `NU (SY|IS)* CL/CP × PO/PR` lines need the
 *   left context and stay deferred.
 * - LB28  `(AL | HL) × (AL | HL)`
 * - LB30  `(AL | HL | NU) × OP`, `CP × (AL | HL | NU)`, both gated on the
 *   bracket NOT being East_Asian_Width F/W/H (`$EastAsian`).
 */
export function isUax14NoBreakPair(prevCp: number, nextCp: number): boolean {
  const prev = lineBreakClass(prevCp);
  // LB14 — unconditional after an opening bracket, whatever follows.
  if (prev === 'OP') return true;

  const next = lineBreakClass(nextCp);
  const prevAlpha = prev === 'AL' || prev === 'HL';
  const nextAlpha = next === 'AL' || next === 'HL';

  // LB28 / LB23 / LB24 — letters with letters, digits, prefixes and postfixes.
  if (prevAlpha && (nextAlpha || next === 'NU' || next === 'PR' || next === 'PO')) return true;
  if (nextAlpha && (prev === 'NU' || prev === 'PR' || prev === 'PO')) return true;

  // LB23a — numeric prefix before, or numeric postfix after, an ideograph.
  if (prev === 'PR' && (next === 'ID' || next === 'EB' || next === 'EM')) return true;
  if ((prev === 'ID' || prev === 'EB' || prev === 'EM') && next === 'PO') return true;

  // LB25 — the adjacency-provable core of a number.
  if (prev === 'NU' && (next === 'NU' || next === 'PO' || next === 'PR')) return true;
  if (next === 'NU' && (prev === 'PO' || prev === 'PR' || prev === 'HY' || prev === 'IS')) {
    return true;
  }

  // LB30 — letters/digits against non-East-Asian parentheses.
  if ((prevAlpha || prev === 'NU') && next === 'OP' && !isEastAsianFWH(nextCp)) return true;
  if (prev === 'CP' && !isEastAsianFWH(prevCp) && (nextAlpha || next === 'NU')) return true;

  return false;
}
