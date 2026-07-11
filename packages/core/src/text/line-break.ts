// UAX #14 — Unicode Line Breaking Algorithm. The generated table supplies the
// effective Line_Break property after the default LB1 resolutions; this module
// is the small typed surface shared by the OOXML renderers.

import {
  LB_CLASS_NAMES,
  LB_RANGE_STARTS,
  LB_RANGE_CLASS,
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
 * A conservative, one-way UAX #14 no-break predicate for segment boundaries.
 *
 * `true` proves that v1's supported rule prohibits the boundary. `false` means
 * only "not proved here" and MUST NOT be interpreted as permission to break.
 * V1 implements LB28: (AL | HL) × (AL | HL).
 */
export function isUax14NoBreakPair(prevCp: number, nextCp: number): boolean {
  const prev = lineBreakClass(prevCp);
  const next = lineBreakClass(nextCp);
  return (prev === 'AL' || prev === 'HL') && (next === 'AL' || next === 'HL');
}
