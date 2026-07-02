/**
 * Virtualization range math for the continuous-scroll viewers (DocxScrollViewer /
 * PptxScrollViewer). Pure and DOM-free — the viewer owns the scroll surface and
 * the slot pool (design §6); this only answers "given the heights, the gap, and
 * the current scrollTop, which item indices must be mounted, where does each
 * item sit, and how tall is the whole scroll region?". Prefix-sum offsets +
 * binary search of the first visible index. See design §5.1
 * (docs/dev-notes/2026-07-01-scroll-viewer-design.md).
 */
export interface VisibleRange {
  /** First index to mount (inclusive, includes overscan). `start > end` ⇒ nothing
   *  to mount (empty input, or a 0-height viewport whose top sits exactly on an
   *  item boundary) — mount loops over `[start, end]` naturally run zero times. */
  start: number;
  /** Last index to mount (inclusive, includes overscan). Empty input ⇒ -1. */
  end: number;
  /** First item intersecting the viewport top (EXCLUDES overscan) — for
   *  onVisiblePageChange. A viewport top strictly inside the gap BETWEEN items i
   *  and i+1 is attributed to item i (gap = trailing padding of the preceding
   *  item — the standard virtualization convention; mount-safe, and flips to i+1
   *  exactly at `offsets[i+1]`). */
  topIndex: number;
  /** Top offset (px) of every item i: `leading` + Σ heights[0..i-1] + i*gap.
   *  length = heights.length. With no padding (`leading` 0) this reduces to the
   *  bare prefix-sum. */
  offsets: number[];
  /** `leading` + Σ heights + (n-1)*gap + `trailing` (gap between items only, none
   *  after the last) → spacer height. With no padding this reduces to
   *  Σ heights + (n-1)*gap. */
  totalHeight: number;
}

/** Optional leading/trailing padding (px) added OUTSIDE the item run — the desk
 *  margin a PDF reader leaves above the first item and below the last. Distinct
 *  from `gap`, which only sits BETWEEN adjacent items. Both default 0, so an
 *  omitted `pad` is exactly the pre-padding behaviour (fully backward-compatible). */
export interface VisibleRangePad {
  /** px above the FIRST item (shifts every offset down by this amount). Default 0. */
  leading?: number;
  /** px below the LAST item (added to totalHeight only). Default 0. */
  trailing?: number;
}

/** Clamp `v` to `[lo, hi]` (hi < lo yields lo — only reached when n === 0, guarded upstream). */
function clamp(v: number, lo: number, hi: number): number {
  return v < lo ? lo : v > hi ? hi : v;
}

/**
 * Compute which item slots to mount for a vertical virtualized scroll region.
 *
 * @param heights        per-item extent along the scroll axis (px), in order.
 * @param gap            px between adjacent items (contributes BETWEEN items only,
 *                       never before the first nor after the last).
 * @param scrollTop      current scroll offset (px); negative is treated as 0 via
 *                       the search / clamps.
 * @param viewportHeight visible height of the scroll surface (px).
 * @param overscan       extra items kept mounted beyond the viewport on each side.
 * @param pad            optional {@link VisibleRangePad} desk margin OUTSIDE the
 *                       item run: `leading` px above the first item shifts every
 *                       offset down; `trailing` px below the last item extends
 *                       totalHeight. Both default 0 — an omitted `pad` is exactly
 *                       the pre-padding behaviour (backward-compatible). A viewport
 *                       top inside the leading pad (scrollTop < leading, so below
 *                       every offset) yields topIndex 0 via the existing clamp.
 * @returns a {@link VisibleRange}. Empty `heights` ⇒
 *          `{ start: 0, end: -1, topIndex: 0, offsets: [], totalHeight: 0 }`
 *          (an empty mount range: `start > end`). NOTE: with n === 0 the padding
 *          is DELIBERATELY NOT applied — an empty document shows no desk padding,
 *          consistent with the viewers' empty-doc no-op contract (they never mount
 *          a spacer for a zero-item document). `pad` only takes effect once there
 *          is at least one item.
 */
export function computeVisibleRange(
  heights: number[],
  gap: number,
  scrollTop: number,
  viewportHeight: number,
  overscan: number,
  pad?: VisibleRangePad,
): VisibleRange {
  const n = heights.length;
  if (n === 0) {
    // Empty doc: no items ⇒ no desk padding (leading/trailing deliberately ignored).
    // Preserves the exact pre-padding empty result the viewers rely on.
    return { start: 0, end: -1, topIndex: 0, offsets: [], totalHeight: 0 };
  }

  const leading = pad?.leading ?? 0;
  const trailing = pad?.trailing ?? 0;

  // Prefix-sum offsets: offsets[i] = leading + Σ heights[0..i-1] + i*gap.
  const offsets = new Array<number>(n);
  let acc = 0;
  for (let i = 0; i < n; i++) {
    offsets[i] = leading + acc + i * gap;
    acc += heights[i];
  }
  // leading + Σ heights + (n-1) gaps + trailing — no gap after the last item; the
  // desk padding brackets the run (leading above the first, trailing below the last).
  const totalHeight = leading + acc + (n - 1) * gap + trailing;

  // topIndex = largest i with offsets[i] <= scrollTop (the item under the
  // viewport top), clamped to [0, n-1]. Binary search over the non-decreasing
  // offsets: find the first index whose offset EXCEEDS scrollTop, minus one.
  let lo = 0;
  let hi = n; // exclusive upper bound
  while (lo < hi) {
    const mid = (lo + hi) >>> 1;
    if (offsets[mid] <= scrollTop) lo = mid + 1;
    else hi = mid;
  }
  const topIndex = clamp(lo - 1, 0, n - 1);

  // lastVisible = last index whose item TOP begins within the viewport
  // (offsets[i] < scrollTop + viewportHeight). Same binary search on the
  // viewport bottom edge.
  const bottom = scrollTop + viewportHeight;
  lo = 0;
  hi = n;
  while (lo < hi) {
    const mid = (lo + hi) >>> 1;
    if (offsets[mid] < bottom) lo = mid + 1;
    else hi = mid;
  }
  const lastVisible = clamp(lo - 1, 0, n - 1);

  const start = clamp(topIndex - overscan, 0, n - 1);
  const end = clamp(lastVisible + overscan, 0, n - 1);

  return { start, end, topIndex, offsets, totalHeight };
}
