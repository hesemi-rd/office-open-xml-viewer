import {
  JT_NAMES,
  JT_RANGE_STARTS,
  JT_RANGE_VALUE,
} from './arabic-joining.generated.js';

export type JoiningType = 'U' | 'C' | 'D' | 'L' | 'R' | 'T';

/** Return the Unicode Joining_Type for one code point. */
export function joiningType(cp: number): JoiningType {
  let lo = 0;
  let hi = JT_RANGE_STARTS.length - 1;
  let rangeIndex = -1;

  while (lo <= hi) {
    const mid = lo + ((hi - lo) >> 1);
    if (JT_RANGE_STARTS[mid] <= cp) {
      rangeIndex = mid;
      lo = mid + 1;
    } else {
      hi = mid - 1;
    }
  }

  if (rangeIndex < 0) return 'U';
  return JT_NAMES[JT_RANGE_VALUE[rangeIndex]] ?? 'U';
}

export function joinsToFollowing(cp: number): boolean {
  const jt = joiningType(cp);
  return jt === 'D' || jt === 'L' || jt === 'C';
}

export function joinsToPreceding(cp: number): boolean {
  const jt = joiningType(cp);
  return jt === 'D' || jt === 'R' || jt === 'C';
}

const LAM = 0x0644;
const LAM_ALEF_PARTNERS = new Set([0x0627, 0x0622, 0x0623, 0x0625, 0x0671]);

/**
 * Return original code-point offsets before which U+0640 may be inserted.
 * Transparent combining marks remain attached to the preceding joining letter.
 */
export function kashidaInsertionPoints(text: string): number[] {
  const cps = [...text].map((ch) => ch.codePointAt(0)!);
  const points: number[] = [];
  let preceding = cps.length > 0 && joiningType(cps[0]) !== 'T' ? 0 : -1;

  for (let k = 1; k < cps.length; k++) {
    const current = cps[k];
    if (joiningType(current) === 'T') continue;

    if (preceding >= 0) {
      const previous = cps[preceding];
      const lamAlef = previous === LAM && LAM_ALEF_PARTNERS.has(current);
      if (
        !lamAlef &&
        joinsToFollowing(previous) &&
        joinsToPreceding(current)
      ) {
        points.push(k);
      }
    }
    preceding = k;
  }

  return points;
}
