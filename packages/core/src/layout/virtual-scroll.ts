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
  /** Top offset (px) of every item i: Σ heights[0..i-1] + i*gap. length = heights.length. */
  offsets: number[];
  /** Σ heights + (n-1)*gap (gap between items only, none after the last) → spacer height. */
  totalHeight: number;
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
 * @returns a {@link VisibleRange}. Empty `heights` ⇒
 *          `{ start: 0, end: -1, topIndex: 0, offsets: [], totalHeight: 0 }`
 *          (an empty mount range: `start > end`).
 */
export function computeVisibleRange(
  heights: number[],
  gap: number,
  scrollTop: number,
  viewportHeight: number,
  overscan: number,
): VisibleRange {
  const n = heights.length;
  if (n === 0) {
    return { start: 0, end: -1, topIndex: 0, offsets: [], totalHeight: 0 };
  }

  // Prefix-sum offsets: offsets[i] = Σ heights[0..i-1] + i*gap.
  const offsets = new Array<number>(n);
  let acc = 0;
  for (let i = 0; i < n; i++) {
    offsets[i] = acc + i * gap;
    acc += heights[i];
  }
  // Σ heights + (n-1) gaps — no trailing gap after the last item.
  const totalHeight = acc + (n - 1) * gap;

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
