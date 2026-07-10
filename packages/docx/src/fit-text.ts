export interface FitTextRun {
  /** §17.3.2.14 w:val — target region width in TWIPS. undefined ⇒ not a fitText run. */
  fitTextValTwips?: number;
  /** §17.3.2.14 w:id — links CONSECUTIVE fitText runs into ONE region. undefined ⇒
   *  standalone region (an id-less fitText run never links to any neighbour). */
  fitTextId?: number | string;
  /** Code-point count of the run. */
  charCount: number;
  /** Natural glyph-advance SUM of the run in px (at the layout scale), BEFORE the
   *  §17.3.2.43 w:w scale and WITHOUT any §17.3.2.35 w:spacing pitch. */
  naturalWidthPx: number;
  /** §17.3.2.43 w:w horizontal glyph scale as a FRACTION (0.66 = 66%). undefined ⇒ 1. */
  charScale?: number;
}

export interface FitTextRegion {
  start: number;
  end: number;
  targetPx: number;
  naturalPx: number;
  charCount: number;
  perGapPx: number;
  trailingPadPx: number;
}

/** ECMA-376 §17.3.2.14. `scale` is px-per-pt. */
export function groupFitTextRegions(runs: FitTextRun[], scale: number): FitTextRegion[] {
  const regions: FitTextRegion[] = [];

  for (let start = 0; start < runs.length; ) {
    const first = runs[start];
    if (first.fitTextValTwips === undefined) {
      start += 1;
      continue;
    }

    let end = start + 1;
    if (first.fitTextId !== undefined) {
      while (
        end < runs.length &&
        runs[end].fitTextValTwips !== undefined &&
        runs[end].fitTextId === first.fitTextId
      ) {
        end += 1;
      }
    }

    let naturalPx = 0;
    let charCount = 0;
    for (let index = start; index < end; index += 1) {
      const run = runs[index];
      naturalPx += run.naturalWidthPx * (run.charScale ?? 1);
      charCount += run.charCount;
    }

    const targetPx = (first.fitTextValTwips / 20) * scale;
    // ECMA-376 §17.3.2.14 requires the region to occupy exactly w:val and
    // describes compression as “decreasing the size of each character”.
    // Word-observed multi-character expansion uses inter-character gaps; until
    // compression ground truth is available, keep every glyph at its natural
    // width and use the same gap formula. A one-character region has no gap, so
    // its residual becomes cell padding AFTER the glyph instead of stretching or
    // shrinking the glyph. A future ground-truth sample may replace that padding
    // (and negative-gap compression) with character scaling.
    const perGapPx = charCount > 1 ? (targetPx - naturalPx) / (charCount - 1) : 0;
    const trailingPadPx = targetPx - naturalPx - Math.max(0, charCount - 1) * perGapPx;

    regions.push({ start, end, targetPx, naturalPx, charCount, perGapPx, trailingPadPx });
    start = end;
  }

  return regions;
}
