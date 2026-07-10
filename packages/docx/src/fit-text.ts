export interface FitTextRun {
  /** §17.3.2.14 w:val — target region width in TWIPS. undefined ⇒ not a fitText run. */
  fitTextValTwips?: number;
  /** §17.3.2.14 w:id — links CONSECUTIVE fitText runs into ONE region. undefined ⇒
   *  standalone region (an id-less fitText run never links to any neighbour). */
  fitTextId?: number;
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
    // ECMA-376 §17.3.2.14 describes compression as “decreasing the size of each
    // character”. Word-observed expansion uses inter-character gaps; until a
    // compression ground truth is available, use the same general gap formula.
    // A future ground-truth sample may require changing compression to char scale.
    const perGapPx = charCount > 1 ? (targetPx - naturalPx) / (charCount - 1) : 0;

    regions.push({ start, end, targetPx, naturalPx, charCount, perGapPx });
    start = end;
  }

  return regions;
}
