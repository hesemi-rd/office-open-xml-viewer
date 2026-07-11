// Justified text alignment for the pptx renderer.
//
// DrawingML offers four "fill the column" alignments (ECMA-376 §20.1.10.59
// ST_TextAlignType): `just` and `justLow` justify a line by widening its
// inter-word and inter-CJK gaps to reach the column width, leaving the
// paragraph's LAST line natural (the spec notes it "does not justify
// sentences which are short"); `dist` and `thaiDist` distribute the same way
// but justify EVERY line, the last included ("Distributes the text words
// across an entire text line"). `just`≡`justLow` (the low-quality variant only
// changes the Arabic kashida sophistication, which Canvas cannot reproduce).
// `thaiDist` distributes like `dist` but at a FINER granularity for Thai/Lao/
// Khmer: the spec says "each character is treated as a word" (§20.1.10.59), so
// its slack falls at every grapheme-CLUSTER boundary of a space-free SEA span,
// not only at inter-word spaces / inter-CJK boundaries. It is therefore NOT an
// alias of `dist` (which leaves such SEA text ragged); see the `thaiDist` mode.
//
// `just` and `dist` use the SAME gap selection (inter-word AND inter-CJK) and
// differ ONLY in the last line (just: natural, dist: filled). This is
// intentionally NOT WordprocessingML's ST_Jc model (§17.18.44), where `both`
// stretches inter-word space only and `distribute` adds inter-character pitch:
// DrawingML/PowerPoint justifies CJK under `just` too — a pure-CJK `just` line
// fills to the column edge in PowerPoint (its same-name PDF is the ground truth
// here; §17.18.44 governs Word, not PowerPoint). The spec's "smart … does not
// justify sentences which are short" clause is realized as the last-line rule
// only; no length threshold is invented for other short lines (the spec gives
// no metric, and a guess would be a heuristic).
//
// The gap geometry itself lives in the shared kernel
// `@silurus/ooxml-core` → `distributeLineSlack` (packages/core/src/text/
// line-distribute.ts), shared with the docx justifier. This module is a thin
// pptx adapter: it owns the DrawingML last-line policy and the slack threshold,
// injects PowerPoint's whitespace predicate (every JS `\s` char is an inter-word
// space — wider than Word's U+0020/U+3000 set), and re-expresses the kernel's
// per-segment result as stretch annotations the renderer applies in place.
//
// Why character-level, not segment-level: the layout merges adjacent
// same-style tokens into ONE segment (see layoutParagraph's `push`/`sameMeta`),
// so a plain paragraph line is usually a single segment holding the whole
// string. Justification therefore has to look INSIDE each segment's text: a
// stretch opportunity is an inter-word space OR an inter-CJK boundary anywhere
// in the line's character stream (boundaries are evaluated across segment
// edges too, so a colour change mid-word doesn't suppress a gap).
//
// `justifyLine` annotates each input segment with `jext` (px to advance AFTER
// the segment when its trailing boundary is a gap) and, when the segment has
// internal gaps, `splitBefore` (code-point offsets) + `perGap` (px). The
// renderer draws each segment in place: if `splitBefore` is empty it does one
// `fillText(seg.text)` and advances by `measureText(seg.text) + jext`; if not,
// it uses `justifiedPiecePositions` (from `@silurus/ooxml-core`) to anchor each
// piece to the whole-string prefix advance — required to avoid 約物半角
// measurement drift, where `Σ measureText(piece)` would exceed
// `measureText(seg.text)` because isolated punctuation loses its contextual
// half-width collapse. The sum of all `jext` plus all `perGap × splits` equals
// the slack, so the painted line reaches `availWidth`.

import { distributeLineSlack, isCjkBreakChar, type DistributeSeg } from '@silurus/ooxml-core';

/** A laid-out segment as seen by the justifier: structurally the shared core
 *  `DistributeSeg`, aliased (not re-declared) so there is a single structural
 *  source. Only the optional `text` matters; an undefined `text` marks an inline
 *  object that bears no stretch of its own (a CJK neighbour can still open a gap
 *  against it). The generic `T` lets the renderer pass its full LayoutSeg and
 *  get the segment back with `jext` (and optionally `splitBefore`/`perGap`). */
export type JustifySeg = DistributeSeg;

export type JustifyMode = 'just' | 'dist' | 'thaiDist';

/** Justification annotation added to each segment of a line. The renderer reads
 *  these to widen the segment's contribution to the line: `jext` advances the
 *  pen AFTER the segment is drawn; `splitBefore` (when present and non-empty)
 *  marks internal CJK gaps at which `perGap` px should be inserted within the
 *  segment's drawn glyphs. */
export type Justified = {
  /** px to advance after drawing this segment (0 when its trailing boundary is
   *  not a gap). */
  jext: number;
  /** Ascending code-point offsets at which internal gaps fall inside the
   *  segment's text. Empty / undefined for inline objects and for segments with
   *  no interior gaps. */
  splitBefore?: number[];
  /** px to insert at each gap inside the segment. Set when `splitBefore` is
   *  non-empty; undefined otherwise. */
  perGap?: number;
};

/** PowerPoint counts EVERY JS-`\s` code point as an inter-word space (\t, \n, the
 *  ideographic space U+3000, …), wider than Word's U+0020/U+3000 set. Injected
 *  into the shared kernel so the pptx gap selection is unchanged by the
 *  extraction. (U+3000 is whitespace here, so it is stretched as one inter-word
 *  gap and never reaches the CJK predicate — never double-counted.) */
const isWsCp = (cp: number): boolean => /\s/.test(String.fromCodePoint(cp));

/**
 * Annotate one laid-out line's segments with the stretch (`jext`,
 * `splitBefore`, `perGap`) needed to fill `availWidth`.
 *
 * @param segments    The line's segments in logical order (LTR only — the
 *                    renderer disables justification under bidi).
 * @param availWidth  Content width to fill, px.
 * @param naturalWidth Sum of the segments' natural advance widths, px.
 * @param mode        'just' (last line stays natural), 'dist' (every line), or
 *                    'thaiDist' ('dist' plus Thai/Lao/Khmer grapheme-cluster gaps
 *                    — §20.1.10.59 "each character is treated as a word").
 * @param isLastLine  Whether this is the paragraph's final line.
 * @returns The same segments (shallow copies) with `jext`/`splitBefore`/`perGap`
 *          added, or `null` when no justification applies (last line under
 *          `just`, no slack, no stretchable gap, …).
 */
export function justifyLine<T extends JustifySeg>(
  segments: readonly T[],
  availWidth: number,
  naturalWidth: number,
  mode: JustifyMode,
  isLastLine: boolean,
): (T & Justified)[] | null {
  // `just`/`justLow` leave the paragraph's last line natural.
  if (mode === 'just' && isLastLine) return null;

  // No room to fill (or already overflowing) → nothing to distribute.
  const slack = availWidth - naturalWidth;
  if (slack <= 0.5) return null;

  // Shared gap selection. pptx justifies from the line start (no leading-indent
  // skip), opens inter-CJK boundaries (expansion), and uses PowerPoint's wide
  // whitespace predicate. It does NOT exclude any segment by index: the final
  // glyph's gap is already suppressed by the kernel's content-span trim (the
  // boundary AFTER the last non-whitespace unit opens no gap), so unlike docx —
  // which draws its final segment separately and excludes it — pptx must let
  // gaps open across EVERY segment boundary, including inside the last segment.
  // `lastDrawnSi = segments.length` is therefore a sentinel that matches no real
  // segment, reproducing the original code-point walk exactly.
  const dist = distributeLineSlack(segments, slack, {
    firstContentSi: 0,
    lastDrawnSi: segments.length,
    isGapChar: isCjkBreakChar,
    isWhitespace: isWsCp,
    // `thaiDist` also opens a gap at every Thai/Lao/Khmer grapheme-cluster
    // boundary, so space-free SEA text is distributed per §20.1.10.59.
    seaClusterGaps: mode === 'thaiDist',
  });
  if (!dist) return null;

  const { perGap, perSeg } = dist;

  // Annotate each segment in place: `jext` widens the trailing boundary;
  // `splitBefore`/`perGap` describe the internal gaps. Σ jext + Σ (perGap ×
  // splits) == slack, so the painted line reaches availWidth.
  const out: (T & Justified)[] = [];
  for (let si = 0; si < segments.length; si++) {
    const seg = segments[si];
    const s = perSeg.get(si);
    const trailing = s?.trailingGap ? perGap : 0;
    const splits = s?.splitBefore;
    if (splits && splits.length > 0) {
      out.push({ ...seg, jext: trailing, splitBefore: [...splits], perGap });
    } else {
      out.push({ ...seg, jext: trailing });
    }
  }
  return out;
}
