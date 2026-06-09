// Renderer-facing layer ABOVE the BidiEngine seam: turns a line's logical
// styled runs into visual-ordered draw segments, and resolves base direction.
//
// NOTE: Task 1 scaffold ships a trivial logical-order passthrough. The real
// implementation (level-boundary ∪ shape-style-boundary cutting + L2 visual
// reordering, and char-data-backed first-strong resolution) lands in Task 5.

import type { BaseDirection, StyledRun, VisualSegment } from './types.js';

/**
 * Resolve a format direction flag to a concrete base direction. When the flag
 * is undefined or 'auto', use UAX#9 first-strong over `text` (P2-P3).
 */
export function resolveBaseDirection(
  flag: boolean | 'auto' | undefined,
  text: string,
): BaseDirection {
  if (flag === true) return 'rtl';
  if (flag === false) return 'ltr';
  // Placeholder first-strong: real version (char-data-backed) lands in Task 5.
  for (const ch of text) {
    const cp = ch.codePointAt(0)!;
    // Hebrew + Arabic strong-RTL blocks (rough; replaced in Task 5).
    if (
      (cp >= 0x0590 && cp <= 0x08ff) ||
      (cp >= 0xfb1d && cp <= 0xfdff) ||
      (cp >= 0xfe70 && cp <= 0xfeff)
    ) {
      return 'rtl';
    }
    if ((cp >= 0x0041 && cp <= 0x005a) || (cp >= 0x0061 && cp <= 0x007a)) {
      return 'ltr';
    }
  }
  return 'ltr';
}

/**
 * Turn a line's logical styled runs into visual-ordered draw segments.
 * Placeholder: passthrough in logical order. Real implementation in Task 5.
 */
export function toVisualSegments(
  runs: StyledRun[],
  base: BaseDirection,
): VisualSegment[] {
  const isRTL = base === 'rtl';
  return runs.map((run) => ({
    text: run.text,
    isRTL,
    level: isRTL ? 1 : 0,
    run,
  }));
}
