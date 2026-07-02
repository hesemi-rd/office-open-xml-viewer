// Width-based label truncation for chart text. Chart labels (axis titles,
// legend entries, category labels) can be arbitrarily long; the renderers used
// to cap them with a fixed `slice(0, N)` character count, which neither
// measures the real rendered width (a CJK label of 6 chars is far wider than a
// Latin label of 6 chars) nor signals truncation to the reader. `elideToWidth`
// replaces that with a measured fit that appends an ellipsis.

const ELLIPSIS = '…';

/**
 * Return `text` trimmed to fit within `maxPx`, appending an ellipsis ('…') when
 * it is shortened. The longest prefix whose `prefix + '…'` width is `<= maxPx`
 * is found by binary search over `ctx.measureText`.
 *
 * Contract:
 * - Uses `ctx.font` **as set by the caller at call time** — set the font (and
 *   any `letterSpacing`) before calling so the measurement matches the eventual
 *   `fillText`. This function does not mutate `ctx` state.
 * - Returns `text` unchanged when it already fits (no ellipsis added).
 * - Returns `''` when `maxPx` cannot even hold the ellipsis (or is <= 0), so a
 *   collapsed slot renders nothing rather than a lone '…'.
 * - An empty `text` returns `''`.
 *
 * @param ctx   canvas context whose current `font` is used for measurement
 * @param text  the full label text
 * @param maxPx the maximum width in device pixels the label may occupy
 */
export function elideToWidth(
  ctx: CanvasRenderingContext2D,
  text: string,
  maxPx: number,
): string {
  if (text === '') return '';
  if (maxPx <= 0) return '';
  if (ctx.measureText(text).width <= maxPx) return text;

  // Not enough room even for the ellipsis alone → render nothing.
  const ellipsisW = ctx.measureText(ELLIPSIS).width;
  if (ellipsisW > maxPx) return '';

  // Binary-search the largest prefix length `k` (in code units) such that
  // `text.slice(0, k) + '…'` fits. `k` ranges over [0, text.length - 1]; the
  // full string is already known not to fit. `k === 0` yields a bare ellipsis,
  // which is the correct minimal indicator when a single glyph won't fit.
  let lo = 0;
  let hi = text.length - 1;
  let best = 0;
  while (lo <= hi) {
    const mid = (lo + hi) >> 1;
    const w = ctx.measureText(text.slice(0, mid) + ELLIPSIS).width;
    if (w <= maxPx) {
      best = mid;
      lo = mid + 1;
    } else {
      hi = mid - 1;
    }
  }
  return text.slice(0, best) + ELLIPSIS;
}
