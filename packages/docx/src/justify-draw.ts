// Glyph placement for a justified segment that carries inter-CJK pitch.
//
// ECMA-376 §17.18.44 `both`/`distribute` fills a CJK line by adding pitch at
// inter-ideograph boundaries (see text-distribute.ts). The renderer therefore
// slices a justified segment at those boundaries and draws the pieces with the
// pitch inserted between them. The SUBTLE part is *where* each piece is drawn.
//
// ── Why pieces must be anchored to the WHOLE-string advance ──────────────────
// A canvas `measureText`/`fillText` over the WHOLE segment applies the browser's
// contextual CJK metrics — most visibly 約物半角, the half-width collapse of
// punctuation like （「」。） when adjacent to kana/kanji. Measuring the sliced
// pieces in ISOLATION loses that collapse, so `Σ measureText(piece)` runs WIDER
// than `measureText(whole)`. The segment's box on the line is `measuredWidth` =
// `measureText(whole)` (+ the internal pitch), and the NEXT segment is drawn at
// that box edge. If we positioned each piece by summing the isolated advances
// (`penX += measureText(piece) + perGap`) the drawn glyphs would overrun the box
// by that drift, and the following run — especially visible at a CJK→Latin
// boundary — would be painted ON TOP of this segment's tail.
//
// Anchoring each piece to the whole-string cumulative advance
// (`measure(prefix)`) instead reproduces exactly the positions `fillText(whole)`
// would use (so 約物半角 is honoured) and lands the final glyph on the box edge
// (`measure(whole) + nGaps·perGap`), so nothing overlaps. The pitch is the only
// thing added on top, and only at the gap offsets the kernel selected.

/** One piece of a sliced justified segment, with the x-offset (from the
 *  segment's origin) at which to draw it. */
export interface JustifiedPiece {
  /** The substring (code points joined) to `fillText`. */
  text: string;
  /** px offset from the segment origin `x` at which to draw `text`. */
  dx: number;
}

/**
 * Compute draw offsets for a justified segment's pieces.
 *
 * The segment's code points are sliced at `splitBefore` (the same code-point
 * offsets the distribute kernel reported as internal gaps). Each resulting piece
 * is anchored to the whole-string cumulative advance up to its start — via
 * `measure` on the prefix — plus `perGap` for every gap that precedes it. This
 * keeps the drawn glyphs aligned with the segment's `measureText(whole)`-based
 * box (so the final glyph reaches `measure(whole) + splitBefore.length·perGap`
 * and the next segment never overlaps), while honouring the contextual CJK
 * metrics baked into `measure`.
 *
 * @param cps        The segment's code points (e.g. `[...seg.text]`).
 * @param splitBefore Ascending code-point offsets (1..len-1) after which an
 *                    internal gap falls; from the distribute kernel.
 * @param perGap     px added at each gap.
 * @param measure    Contextual width of a string — `ctx.measureText(s).width`
 *                   with the segment's font already selected on `ctx`.
 * @returns One {@link JustifiedPiece} per slice, in draw order.
 */
export function justifiedPiecePositions(
  cps: readonly string[],
  splitBefore: readonly number[],
  perGap: number,
  measure: (s: string) => number,
): JustifiedPiece[] {
  const pieces: JustifiedPiece[] = [];
  let from = 0;
  let gapsSeen = 0;
  const emit = (cut: number): void => {
    // Anchor to the whole-string advance up to this piece's first code point,
    // NOT the running sum of isolated piece advances — see the module header.
    const dx = measure(cps.slice(0, from).join('')) + gapsSeen * perGap;
    pieces.push({ text: cps.slice(from, cut).join(''), dx });
    from = cut;
    gapsSeen++;
  };
  for (const cut of splitBefore) emit(cut);
  emit(cps.length); // final piece (no trailing gap of its own)
  return pieces;
}
