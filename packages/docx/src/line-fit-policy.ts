/**
 * ECMA-376 §17.4.6 and §17.3.1.44-.45 line-break selection: returns the
 * largest `end` in (start, limitEnd] whose caller-model cumulative height for
 * [start, end) fits `available`. The walk stops at the first non-fitting end.
 */
export function selectLargestFittingEnd(
  start: number,
  limitEnd: number,
  available: number,
  heightAt: (end: number) => number,
): { end: number; height: number } {
  let end = start;
  let height = 0;
  for (let candidate = start + 1; candidate <= limitEnd; candidate++) {
    const candidateHeight = heightAt(candidate);
    if (!(candidateHeight <= available)) break;
    end = candidate;
    height = candidateHeight;
  }
  return { end, height };
}

/**
 * ECMA-376 §17.3.1.44-.45 widow-orphan adjustment following greedy line
 * selection. This pure decision mirrors the existing body pagination policy;
 * §17.4.6 governs whether the analogous containing table row may split.
 */
export function adjustForWidowOrphan(input: {
  widowControl: boolean;
  start: number;
  end: number;
  totalLines: number;
  belowColumnTop: boolean;
}): { kind: 'keep' } | { kind: 'dropLastLine' } | { kind: 'relocate' } {
  if (!input.widowControl || input.end >= input.totalLines) return { kind: 'keep' };
  if (input.totalLines - input.end === 1 && input.end - input.start >= 2) {
    return { kind: 'dropLastLine' };
  }
  if (input.start === 0 && input.end - input.start === 1 && input.belowColumnTop) {
    return { kind: 'relocate' };
  }
  return { kind: 'keep' };
}
