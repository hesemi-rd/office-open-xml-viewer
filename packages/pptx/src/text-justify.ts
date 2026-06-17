// Justified text alignment for the pptx renderer.
//
// DrawingML offers four "fill the column" alignments (ECMA-376 §20.1.10.59
// ST_TextAlignType): `just` and `justLow` justify a line by widening its
// inter-word and inter-CJK gaps to reach the column width, leaving the
// paragraph's LAST line natural (the spec notes it "does not justify
// sentences which are short"); `dist` and `thaiDist` distribute the same way
// but justify EVERY line, the last included ("Distributes the text words
// across an entire text line"). justLow/thaiDist are the low-quality variants
// — same distribution, only the kashida/glyph-stretch sophistication differs,
// which Canvas cannot reproduce — so we treat just≡justLow and dist≡thaiDist.
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
// per-segment result as the renderer's draw pieces.
//
// Why character-level, not segment-level: the layout merges adjacent
// same-style tokens into ONE segment (see layoutParagraph's `push`/`sameMeta`),
// so a plain paragraph line is usually a single segment holding the whole
// string. Justification therefore has to look INSIDE each segment's text: a
// stretch opportunity is an inter-word space OR an inter-CJK boundary anywhere
// in the line's character stream (boundaries are evaluated across segment
// edges too, so a colour change mid-word doesn't suppress a gap).
//
// `justifyLine` returns the line re-expressed as draw pieces: each input
// segment is split at the gap positions, and every piece carries `jext`, the
// px to advance AFTER drawing it. The renderer simply draws each piece and adds
// its `jext` — the existing glyph paths are untouched. The sum of all `jext`
// equals the slack, so the painted line reaches `availWidth`.
//
// Splits land only at gap positions (inter-word spaces and inter-CJK
// boundaries), never inside a Latin word, so a split never cuts a kerning pair
// or ligature. The per-piece advance widths the renderer re-measures therefore
// sum to the whole-line `naturalWidth` passed in, and the painted line lands on
// `availWidth` without measurement drift.

import { distributeLineSlack, isCjkBreakChar, type DistributeSeg } from '@silurus/ooxml-core';

/** A laid-out segment as seen by the justifier: structurally the shared core
 *  `DistributeSeg`, aliased (not re-declared) so there is a single structural
 *  source. Only the optional `text` matters; an undefined `text` marks an inline
 *  object that bears no stretch of its own (a CJK neighbour can still open a gap
 *  against it). The generic `T` lets the renderer pass its full LayoutSeg and
 *  get pieces that keep every style field, plus `jext`. */
export type JustifySeg = DistributeSeg;

export type JustifyMode = 'just' | 'dist';

/** PowerPoint counts EVERY JS-`\s` code point as an inter-word space (\t, \n, the
 *  ideographic space U+3000, …), wider than Word's U+0020/U+3000 set. Injected
 *  into the shared kernel so the pptx gap selection is unchanged by the
 *  extraction. (U+3000 is whitespace here, so it is stretched as one inter-word
 *  gap and never reaches the CJK predicate — never double-counted.) */
const isWsCp = (cp: number): boolean => /\s/.test(String.fromCodePoint(cp));

/**
 * Split one laid-out line into draw pieces that fill `availWidth` when each
 * piece's `jext` is added to the pen after it is drawn.
 *
 * @param segments    The line's segments in logical order (LTR only — the
 *                    renderer disables justification under bidi).
 * @param availWidth  Content width to fill, px.
 * @param naturalWidth Sum of the segments' natural advance widths, px.
 * @param mode        'just' (last line stays natural) or 'dist' (every line).
 * @param isLastLine  Whether this is the paragraph's final line.
 * @returns The pieces (each a shallow copy of its source segment with sliced
 *          `text` and a `jext` advance), or `null` when no justification
 *          applies (last line under `just`, no slack, no stretchable gap, …).
 */
export function justifyLine<T extends JustifySeg>(
  segments: readonly T[],
  availWidth: number,
  naturalWidth: number,
  mode: JustifyMode,
  isLastLine: boolean,
): (T & { jext: number })[] | null {
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
  });
  if (!dist) return null;

  const { perGap, perSeg } = dist;

  // Re-express the per-segment stretch as the renderer's draw pieces. For each
  // segment: an inline object is one piece (a trailing gap → jext = perGap);
  // a text segment is sliced at the kernel's interior split offsets, each piece
  // before a split advancing perGap, and the final piece advancing perGap iff
  // the boundary AFTER the segment is a gap (trailingGap). This matches the old
  // code-point walk exactly — Σ jext == slack, final glyph reaches availWidth.
  const out: (T & { jext: number })[] = [];
  for (let si = 0; si < segments.length; si++) {
    const seg = segments[si];
    const s = perSeg.get(si);
    if (seg.text === undefined) {
      out.push({ ...seg, jext: s?.trailingGap ? perGap : 0 });
      continue;
    }
    const cps = [...seg.text]; // code points (handles surrogate pairs)
    const splits = s?.splitBefore ?? [];
    let from = 0;
    for (const cut of splits) {
      out.push({ ...seg, text: cps.slice(from, cut).join(''), jext: perGap });
      from = cut;
    }
    // Final piece of this segment: trailingGap → perGap, else 0.
    out.push({ ...seg, text: cps.slice(from).join(''), jext: s?.trailingGap ? perGap : 0 });
  }
  return out;
}
