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

import { isCjkBreakChar } from '@silurus/ooxml-core';

/** A boundary is a stretch opportunity when either side is a CJK / ideographic
 *  glyph (see core's `isCjkBreakChar` for the canonical ranges). U+3000
 *  (ideographic space) is classified as whitespace below, so it never reaches
 *  this test — which is why sharing the wrap-side predicate (it includes U+3000)
 *  is safe here: a U+3000 unit is stretched as an inter-word gap, never reaching
 *  the CJK test, so it is never double-counted. */
const isCjk = (ch: string): boolean => isCjkBreakChar(ch.codePointAt(0) ?? 0);

/** A laid-out segment as seen by the justifier. Only the optional text matters;
 *  an undefined `text` marks an inline object (math / image), which is one
 *  opaque unit that bears no stretch of its own (a CJK neighbour can still open
 *  a gap against it). The generic `T` lets the renderer pass its full LayoutSeg
 *  and get pieces that keep every style field, plus `jext`. */
export interface JustifySeg {
  text?: string;
}

export type JustifyMode = 'just' | 'dist';

const isWsChar = (ch: string): boolean => /\s/.test(ch);

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

  // Flatten to a unit stream: each code point of a text segment is a unit; an
  // inline object is a single (non-char, non-ws) unit.
  type Unit = { ch?: string; ws: boolean };
  const units: Unit[] = [];
  for (const seg of segments) {
    if (seg.text === undefined) {
      units.push({ ws: false });
    } else {
      for (const ch of seg.text) units.push({ ch, ws: isWsChar(ch) });
    }
  }

  // Content span: leading/trailing whitespace must not stretch, and a single
  // content unit (one word / one glyph) has no inner gap.
  let first = -1;
  let last = -1;
  for (let k = 0; k < units.length; k++) {
    if (!units[k].ws) {
      if (first === -1) first = k;
      last = k;
    }
  }
  if (first === -1 || first === last) return null;

  // Mark the gap AFTER each qualifying unit. Whitespace → inter-word (one gap
  // per space, Word stretches each). Non-whitespace → an inter-CJK gap only
  // when the boundary to the next NON-whitespace unit touches a CJK glyph; a
  // boundary into whitespace is already counted by that whitespace, so it is
  // never double-counted here.
  const gapAfter = new Array<boolean>(units.length).fill(false);
  let total = 0;
  for (let k = first; k < last; k++) {
    const u = units[k];
    if (u.ws) {
      gapAfter[k] = true;
      total++;
    } else {
      const nx = units[k + 1];
      if (!nx.ws) {
        const lc = u.ch;
        const rc = nx.ch;
        if ((lc !== undefined && isCjk(lc)) || (rc !== undefined && isCjk(rc))) {
          gapAfter[k] = true;
          total++;
        }
      }
    }
  }
  if (total === 0) return null; // e.g. a single long Latin word that wrapped alone

  const perGap = slack / total;

  // Re-walk the SAME unit order, splitting each segment at gap positions.
  const out: (T & { jext: number })[] = [];
  let k = 0;
  for (const seg of segments) {
    if (seg.text === undefined) {
      out.push({ ...seg, jext: gapAfter[k] ? perGap : 0 });
      k++;
      continue;
    }
    let buf = '';
    for (const ch of seg.text) {
      buf += ch;
      const g = gapAfter[k];
      k++;
      if (g) {
        out.push({ ...seg, text: buf, jext: perGap });
        buf = '';
      }
    }
    if (buf !== '') out.push({ ...seg, text: buf, jext: 0 });
  }
  return out;
}
