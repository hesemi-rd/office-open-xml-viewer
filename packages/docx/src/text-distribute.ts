// Justified-text slack distribution for the docx renderer.
//
// ECMA-376 / ISO-29500 §17.18.44 ST_Jc defines the "fill the line" alignments:
//
//   both (Justified)  "This type of justification shall only affect the inter-
//                      word spacing on each line, and not the inter-character
//                      spacing within each word when justifying its contents."
//   distribute        "This type of justification shall equally affect the
//                      inter-word spacing on each line as well as the inter-
//                      character spacing between each word ... an equal amount of
//                      additional character pitch shall be added to all
//                      characters on the line."
//
// The `both` clause "inter-word, not inter-character within each word" is written
// for LATIN script, where a word is a run of letters delimited by spaces. CJK has
// no inter-word spaces and each ideograph behaves as its own word, so an
// inter-CJK boundary IS an inter-word boundary, not an intra-word one. Word
// confirms this empirically: in the sample-9 Word PDF a pure-CJK `both` line that
// wrapped — with NO ASCII spaces at all — is filled exactly to the right text
// margin (pdftotext -bbox: the 0-space line reaches xMax 523.0pt, the page's
// right margin, while the paragraph's natural last line stops short at 358.0pt).
// The only mechanism that can fill a space-free line is adding pitch at inter-CJK
// boundaries, so Word's `both` widens BOTH inter-word ASCII spaces AND inter-CJK
// boundaries; it does NOT widen the boundary between two Latin letters of one
// word.
//
// `both` and `distribute` therefore select the SAME gap opportunities and differ
// ONLY in the last line: `both` leaves the paragraph's final line natural;
// `distribute` fills it too. (The renderer decides whether to call this kernel
// for a given line; the kernel just stretches whatever line it is handed.) This
// mirrors the pptx text-justify kernel — see packages/pptx/src/text-justify.ts —
// which reached the same gap model from the PowerPoint PDFs.
//
// Why a per-segment / per-character model: docx layout splits a paragraph into
// segments at spaces and style boundaries (splitTextForLayout), so a CJK phrase
// like "観察することで" is ONE segment with no internal spaces. Its inter-CJK
// gaps fall INSIDE the segment, between code points — the old "advance after
// trailing ASCII spaces only" model could never open them. This kernel walks the
// line's whole code-point stream (boundaries evaluated across segment edges, so a
// colour change mid-phrase doesn't swallow a gap) and reports, per segment, where
// each gap falls: internal split points (to slice the glyph drawing) and a
// trailing-edge flag (the inter-segment boundary). The renderer applies `perGap`
// at each.

/** Single-code-point CJK test, sharing the ranges hasCJKBreakOpportunity uses in
 *  renderer.ts (CJK symbols/punctuation + Unified incl. Ext-A via 0x3000–0x9FFF,
 *  CJK Compatibility Ideographs, Hangul syllables, and Halfwidth/Fullwidth
 *  forms). A boundary between two non-space code points is a stretch opportunity
 *  when either side is one of these. NOTE: U+3000 (ideographic space) falls in
 *  the 0x3000–0x9FFF block but the gap walker classifies whitespace FIRST (see
 *  isJustifyWhitespace), so an ideographic space is treated as one inter-word gap
 *  and never reaches this test — matching the pptx CJK_RE which starts at U+3001. */
export function isCJKCodePoint(cp: number): boolean {
  return (
    (cp >= 0x3000 && cp <= 0x9fff) ||
    (cp >= 0xf900 && cp <= 0xfaff) ||
    (cp >= 0xac00 && cp <= 0xd7af) ||
    (cp >= 0xff00 && cp <= 0xffef)
  );
}

/** Whitespace that participates as an INTER-WORD gap: the ASCII space (U+0020)
 *  and the ideographic space (U+3000). Both are stretched as one gap each, like
 *  any inter-word space, rather than as CJK boundaries around them. (JS `\s`
 *  matches U+3000 too, so this stays consistent with the renderer's `/\S/`
 *  leading-indent detection.) */
const isJustifyWhitespace = (cp: number): boolean => cp === 0x20 || cp === 0x3000;

/** A laid-out segment as the distributor sees it. Only the optional text matters;
 *  an undefined `text` marks a non-text inline atom (image / math / tab) — one
 *  opaque unit bearing no stretch of its own, though a CJK neighbour can still
 *  open a gap against its edge (the atom counts as a non-CJK, non-space unit). */
export interface DistributeSeg {
  text?: string;
}

/** Per-text-segment instructions for applying the line's slack.
 *
 *  `splitBefore` lists the code-point offsets (1..len-1, counted in code points,
 *  NOT UTF-16 units) at which an INTERNAL gap falls — the glyphs before the
 *  offset are drawn, the pen advances by `perGap`, then drawing resumes.
 *  `trailingGap` marks that the boundary AFTER this segment's last code point
 *  (the inter-segment boundary) is a stretch opportunity. */
export interface SegStretch {
  /** Code-point offsets inside the segment after which to insert `perGap`. */
  splitBefore: number[];
  /** Whether to advance `perGap` after the whole segment (inter-segment gap). */
  trailingGap: boolean;
  /** px added strictly INSIDE the segment = splitBefore.length * perGap.
   *  Decorations (highlight / underline / strike / ruby centring / onTextRun
   *  width) should span measuredWidth + internalStretch. */
  internalStretch: number;
}

/** Result of distributing a line's slack across its gap opportunities. */
export interface DistributeResult {
  /** px added at each gap (negative when the line is compressed). */
  perGap: number;
  /** Per-segment stretch, keyed by the segment's index in `segments`. Segments
   *  with no gap (and non-text atoms) are absent. */
  perSeg: Map<number, SegStretch>;
}

/**
 * Distribute `slack` px equally across a justified line's gap opportunities
 * (inter-word ASCII spaces AND inter-CJK boundaries), per ECMA-376 §17.18.44.
 *
 * A gap is opened AFTER a code point whose owning segment is in
 * `[firstContentSi, lastDrawnSi)` — i.e. the LEFT side of every gap is eligible,
 * while the gap's right side may be the first code point of the final segment (so
 * a 2-token line "the quick" still widens the space before "quick", and a CJK
 * line split into [観察][結果] still widens the 察|結 boundary). The final segment
 * thus opens no gap of its own and is never split internally, but the boundary
 * INTO it stretches like any other — all slack lands to the final glyph's left,
 * which reaches the margin (Σgaps == slack). Also excluded, matching the
 * renderer's pre-existing rules:
 *   - segments before `firstContentSi` (a paragraph's 字下げ leading-indent
 *     whitespace) stay fixed;
 *   - leading whitespace before the first content unit, and the gap after the
 *     last content unit, never stretch.
 *
 * @param segments       The line's segments in LOGICAL (reading) order.
 * @param slack          availWidth - naturalWidth, px. >0 stretches; <0 compresses.
 * @param firstContentSi Index of the first segment holding non-whitespace content;
 *                       earlier (leading-indent) whitespace segments are fixed.
 *                       Pass 0 to disable the skip (e.g. under bidi).
 * @param lastDrawnSi    Index of the visually-final segment; it and the boundary
 *                       into it get no gap.
 * @param minPerGap      Lower bound on a (negative) `perGap` when compressing
 *                       (slack < 0), so we never eat more than a capped amount per
 *                       gap. Ignored when slack >= 0.
 * @param includeCJK     When true (default), inter-CJK boundaries open gaps too
 *                       (the §17.18.44 "additional character pitch shall be added"
 *                       behaviour, used for EXPANSION). When false, only inter-word
 *                       spaces open gaps — used for COMPRESSION (negative slack),
 *                       where shrinking a space glyph is legitimate but overlapping
 *                       two ideographs is not, so the renderer passes
 *                       includeCJK = slack > 0.
 * @returns The distribution, or `null` when nothing stretches (no eligible gap,
 *          or |slack| below the 0.5px noise floor).
 */
export function distributeLineSlack(
  segments: readonly DistributeSeg[],
  slack: number,
  firstContentSi: number,
  lastDrawnSi: number,
  minPerGap = -Infinity,
  includeCJK = true,
): DistributeResult | null {
  if (Math.abs(slack) <= 0.5) return null;

  // Flatten the eligible segments to a code-point stream. Segments before
  // `firstContentSi` (leading 字下げ indent) are skipped; the rest — INCLUDING
  // the final segment — are flattened, because a gap may have its RIGHT side in
  // the final segment even though its left side may not be. Each unit carries its
  // owning segment, the code-point offset within that segment, the code point,
  // and an inter-word-whitespace flag (ASCII + ideographic space; see
  // isJustifyWhitespace). Non-text atoms become one unit with cp=undefined.
  type Unit = { si: number; off: number; cp?: number; ws: boolean };
  const units: Unit[] = [];
  for (let si = firstContentSi; si < segments.length; si++) {
    const seg = segments[si];
    if (seg === undefined) continue;
    if (seg.text === undefined) {
      units.push({ si, off: 0, ws: false });
      continue;
    }
    let off = 0;
    for (const ch of seg.text) {
      const cp = ch.codePointAt(0)!;
      units.push({ si, off, cp, ws: isJustifyWhitespace(cp) });
      off++;
    }
  }

  // Content span: the first and last NON-whitespace units. Leading/trailing
  // spaces never stretch, and a single content unit has no interior gap.
  let first = -1;
  let last = -1;
  for (let k = 0; k < units.length; k++) {
    if (!units[k].ws) {
      if (first === -1) first = k;
      last = k;
    }
  }
  if (first === -1 || first === last) return null;

  // Mark a gap AFTER unit k, for first <= k < last, when k's owning segment is
  // eligible to OPEN a gap (si < lastDrawnSi — the final segment opens none and
  // is never split, but the boundary INTO it still stretches). Whitespace → one
  // inter-word gap (Word stretches each space). Non-space → an inter-CJK gap only
  // when the boundary to the next NON-space unit touches a CJK glyph; a boundary
  // INTO whitespace is already counted by that whitespace, so it is not
  // double-counted.
  const gapAfter = new Array<boolean>(units.length).fill(false);
  let total = 0;
  for (let k = first; k < last; k++) {
    const u = units[k];
    if (u.si >= lastDrawnSi) continue; // final segment opens no gap
    if (u.ws) {
      gapAfter[k] = true;
      total++;
      continue;
    }
    if (!includeCJK) continue; // compression path: only spaces open gaps
    const nx = units[k + 1];
    if (nx.ws) continue; // counted by that space when we reach it
    const lc = u.cp;
    const rc = nx.cp;
    if (
      (lc !== undefined && isCJKCodePoint(lc)) ||
      (rc !== undefined && isCJKCodePoint(rc))
    ) {
      gapAfter[k] = true;
      total++;
    }
  }
  if (total === 0) return null;

  let perGap = slack / total;
  if (slack < 0 && perGap < minPerGap) perGap = minPerGap;

  // Per-segment code-point length, to tell an internal gap (off < len-1) from the
  // trailing inter-segment gap (off === len-1).
  const segLen = new Map<number, number>();
  for (const u of units) {
    if (u.cp !== undefined) segLen.set(u.si, (segLen.get(u.si) ?? 0) + 1);
  }

  const perSeg = new Map<number, SegStretch>();
  for (let k = 0; k < units.length; k++) {
    if (!gapAfter[k]) continue;
    const u = units[k];
    let s = perSeg.get(u.si);
    if (!s) {
      s = { splitBefore: [], trailingGap: false, internalStretch: 0 };
      perSeg.set(u.si, s);
    }
    const len = segLen.get(u.si) ?? 0;
    if (u.cp === undefined || u.off === len - 1) {
      s.trailingGap = true; // inter-segment boundary
    } else {
      s.splitBefore.push(u.off + 1); // split BEFORE code point off+1
      s.internalStretch += perGap;
    }
  }
  return { perGap, perSeg };
}
