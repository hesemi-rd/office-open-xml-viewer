/** One entry of a line's LOGICAL-order segment sequence, as seen by
 *  {@link resolveTabWidths}: a tab (its gap is computed) or content (its
 *  measured advance in px). */
export interface TabItem {
  isTab: boolean;
  width: number;
}

/** An `a:tab` from `a:pPr > a:tabLst` (§21.1.2.1.x) in reading-frame px:
 *  `pos` is the distance from the LEADING text-inset edge (logical, not
 *  physical — the right inset under an RTL base), `algn` is `l|ctr|r|dec`. */
export interface TabStopPx {
  pos: number;
  algn: string;
}

/**
 * ECMA-376 §21.1.2.1.x — resolve each inline TAB segment's GAP width against
 * the paragraph's custom stop grid, in READING-frame px measured from the
 * leading text-inset edge (issue #916, generalising the #913 single-cell
 * semantics to N cells; mirrors docx `layoutBidiTabStops`).
 *
 * The walk is base-agnostic: the pen advances through `items` in LOGICAL
 * order from `startPen` (the leading indent, plus any first-line indent); a
 * tab jumps to the nearest stop strictly past the pen, placing the FOLLOWING
 * cell (the content run up to the next tab / line end) so it ends on (`r`,
 * `dec`), centres on (`ctr`), or starts at (`l`) the stop. Under an RTL base
 * the SAME reading-frame gaps apply — UAX#9 L2 (tabs are Bidi_Class S, see
 * bidi-line.ts) reverses cells and tabs together, so each logical gap IS its
 * visual gap and the cumulative draw reproduces the mirrored stops.
 *
 * `dec` aligns like `r` (frac 1): true decimal-point splitting is not
 * implemented — the same approximation docx uses. DrawingML defines no
 * automatic tab grid (there is no `defTabSz` support in this parser), so a
 * tab with no reachable stop degrades to `noStopGap` (a space width,
 * preserving the pre-#916 behaviour). Two clamps mirror docx: a cell whose
 * far edge would cross `limit` (the trailing text-inset edge) is pinned to it
 * (#835 — Word never pushes a cell off the text area), and a tab never moves
 * the pen backwards.
 *
 * @returns one width per LOGICAL index (non-tabs keep their input width).
 */
export function resolveTabWidths(
  items: readonly TabItem[],
  stops: readonly TabStopPx[],
  startPen: number,
  limit: number,
  noStopGap: number,
): number[] {
  const widths = items.map((item) => item.width);
  const followW = (from: number): number => {
    let width = 0;
    for (let i = from; i < items.length && !items[i].isTab; i++) {
      width += widths[i];
    }
    return width;
  };

  let pen = startPen;
  for (let i = 0; i < items.length; i++) {
    if (!items[i].isTab) {
      pen += widths[i];
      continue;
    }

    let stop: TabStopPx | null = null;
    for (const candidate of stops) {
      if (candidate.pos > pen && (stop === null || candidate.pos < stop.pos)) {
        stop = candidate;
      }
    }
    if (stop === null) {
      widths[i] = noStopGap;
      pen += noStopGap;
      continue;
    }

    const followingWidth = followW(i + 1);
    const fraction = stop.algn === 'ctr'
      ? 0.5
      : stop.algn === 'r' || stop.algn === 'dec'
        ? 1
        : 0;
    let target = stop.pos - followingWidth * fraction;
    if (target + followingWidth > limit) target = limit - followingWidth;
    if (target < pen) target = pen;
    widths[i] = target - pen;
    pen = target;
  }

  return widths;
}
