// Tabular (fixed-advance) digit rendering for Canvas 2D.
//
// The media playback time ("m:ss / m:ss") is drawn with a proportional UI font,
// so as the digits tick (e.g. "0:01" → "0:08" → "0:11") their differing widths
// shift the layout and the text appears to jitter. CSS solves this with
// `font-variant-numeric: tabular-nums`, but Canvas 2D exposes no way to set it.
//
// Instead we lay each digit out in a fixed-width cell (the widest digit's
// advance) and center the glyph within it — the same visual result as tabular
// figures, while keeping the natural (non-monospace) typeface. Punctuation
// (':', '/', spaces) keeps its natural advance.
//
// Only `measureText`/`fillText`/`textAlign` are used, so any object providing
// them works (the production caller passes a CanvasRenderingContext2D).
type TextCtx = Pick<CanvasRenderingContext2D, 'measureText' | 'fillText' | 'textAlign'>;

const isDigit = (ch: string): boolean => ch >= '0' && ch <= '9';

/** Widest advance among the digits 0–9 for the context's current font. */
export function tabularDigitWidth(ctx: TextCtx): number {
  let w = 0;
  for (let d = 0; d < 10; d++) w = Math.max(w, ctx.measureText(String(d)).width);
  return w;
}

/** Width of `text` when every digit occupies a fixed `digitW` cell and other
 *  glyphs keep their natural advance. Equal for any two strings with the same
 *  digit/non-digit shape — which is what removes the jitter. */
export function tabularTextWidth(ctx: TextCtx, text: string, digitW: number): number {
  let w = 0;
  for (const ch of text) w += isDigit(ch) ? digitW : ctx.measureText(ch).width;
  return w;
}

/** Draw `text` from left edge `leftX`, centering each digit inside a fixed
 *  `digitW` cell (tabular figures) while keeping the natural font. Honors the
 *  current baseline/fill; saves and restores `textAlign`. */
export function drawTabularText(
  ctx: TextCtx,
  text: string,
  leftX: number,
  y: number,
  digitW: number,
): void {
  const prevAlign = ctx.textAlign;
  ctx.textAlign = 'left';
  let cx = leftX;
  for (const ch of text) {
    if (isDigit(ch)) {
      const gw = ctx.measureText(ch).width;
      ctx.fillText(ch, cx + (digitW - gw) / 2, y); // center the glyph in its cell
      cx += digitW;
    } else {
      ctx.fillText(ch, cx, y);
      cx += ctx.measureText(ch).width;
    }
  }
  ctx.textAlign = prevAlign;
}
