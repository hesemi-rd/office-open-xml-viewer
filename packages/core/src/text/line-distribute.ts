// Shared "fill the line" slack distributor for justified text — the single gap
// model behind WordprocessingML §17.18.44 (`both`/`distribute`) and DrawingML
// §20.1.10.59 (`just`/`dist`). Given a laid-out line's segments (logical order)
// and its slack, it reports — per segment — where each gap falls and the per-gap
// px, so a renderer can slice the glyph drawing and advance the pen.
//
// ── Why one kernel for both formats ─────────────────────────────────────────
// Word and PowerPoint reach the SAME gap selection: a stretch opportunity is an
// inter-word space OR an inter-CJK boundary (either side a CJK / ideographic
// glyph), evaluated across the whole line's code-point stream — boundaries are
// tested across segment edges too, so a colour change mid-phrase doesn't swallow
// a gap. They differ only in policy the CALLER owns and which this kernel does
// NOT encode:
//   • last-line policy (Word `both` / PowerPoint `just` leave the final line
//     natural; `distribute` / `dist` fill it). The caller decides whether to
//     call the kernel for a given line — the kernel just stretches whatever line
//     it is handed.
//   • the whitespace predicate. PowerPoint treats every JS `\s` char as an
//     inter-word space; Word counts only U+0020 and U+3000. Each caller injects
//     its own `isWhitespace` so the kernel reproduces that format's behaviour.
//   • the gap predicate. EXPANSION opens inter-CJK boundaries (`isGapChar` =
//     core.isCjkBreakChar); a COMPRESSION path that may shrink spaces but must
//     not overlap ideographs injects `isGapChar: () => false`.
//
// Why a per-segment / per-character model: layout merges adjacent same-style
// tokens into ONE segment, so a CJK phrase like "観察することで" is a single
// segment with no internal spaces — its inter-CJK gaps fall INSIDE the segment,
// between code points. The kernel walks the line's whole code-point stream and
// reports, per segment, the interior split offsets and a trailing-edge flag.
//
// This file is format-agnostic: no ECMA-376 §-specific policy lives here, only
// the gap geometry both renderers share. See packages/pptx/src/text-justify.ts
// and packages/docx/src/text-distribute.ts for the format adapters.

import { isCjkBreakChar } from './cjk-ranges.js';

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

/** Tuning for {@link distributeLineSlack}; every field is optional and defaults
 *  to the WordprocessingML expansion behaviour. */
export interface DistributeOptions {
  /** Index of the first segment holding non-whitespace content; earlier
   *  (leading-indent / 字下げ whitespace) segments are fixed and never open a
   *  gap. Default 0 (no skip — e.g. under bidi, or for callers with no indent
   *  concept). */
  firstContentSi?: number;
  /** Index of the VISUALLY-final segment; it and the boundary into it open no
   *  gap. Default `segments.length - 1`. The match is EXACT, not `>=`: under
   *  bidi this is the visually-last segment's LOGICAL index, which is not the
   *  maximum si, so `>=` would wrongly suppress every logically-later segment
   *  (a pure-RTL line would skip the whole line → no justification). Pass a value
   *  ≥ segments.length (e.g. `segments.length`) to exclude NO segment: the pptx
   *  adapter does this because it draws every segment in one loop and relies on
   *  the content-span trim to suppress only the final glyph's gap. */
  lastDrawnSi?: number;
  /** Lower bound on a (negative) `perGap` when compressing (slack < 0), so a
   *  compression never eats more than a capped amount per gap. Default
   *  -Infinity (uncapped). Ignored when slack >= 0. */
  minPerGap?: number;
  /** A boundary between two non-space code points opens an inter-CJK gap when
   *  EITHER side satisfies this predicate. Default core.isCjkBreakChar (open a
   *  gap when either side is a CJK / ideographic glyph). Pass `() => false` for
   *  a compression path that must not overlap ideographs (only spaces stretch). */
  isGapChar?: (cp: number) => boolean;
  /** Classifies a code point as inter-word whitespace: it becomes ONE gap and is
   *  never treated as a CJK boundary. Default `cp === 0x20 || cp === 0x3000`
   *  (the WordprocessingML set). PowerPoint injects a wider JS-`\s` predicate.
   *  Whitespace is classified FIRST, so an ideographic space is one inter-word
   *  gap and never reaches `isGapChar`. */
  isWhitespace?: (cp: number) => boolean;
}

/** Default inter-word whitespace (WordprocessingML): ASCII space + ideographic
 *  space. */
const defaultIsWhitespace = (cp: number): boolean => cp === 0x20 || cp === 0x3000;

/**
 * Distribute `slack` px equally across a justified line's gap opportunities
 * (inter-word spaces AND, by default, inter-CJK boundaries).
 *
 * A gap is opened AFTER a code point whose owning segment is eligible to OPEN a
 * gap (its index is in `[firstContentSi, segments.length)` and is NOT
 * `lastDrawnSi`). The LEFT side of every gap is thus eligible, while the gap's
 * right side may be the first code point of the final segment — so a 2-token
 * line "the quick" widens the space before "quick", and a CJK line split into
 * [観察][結果] widens the 察|結 boundary. The final segment opens no gap of its
 * own and is never split internally, but the boundary INTO it stretches like any
 * other; all slack lands to the final glyph's left, which reaches the margin
 * (Σgaps == slack). Leading whitespace before the first content unit and the gap
 * after the last content unit never stretch.
 *
 * @param segments The line's segments in LOGICAL (reading) order.
 * @param slack    availWidth - naturalWidth, px. >0 stretches; <0 compresses.
 * @param opts     See {@link DistributeOptions}; omitted fields take the
 *                 WordprocessingML-expansion defaults.
 * @returns The distribution, or `null` when nothing stretches (no eligible gap,
 *          or |slack| below the 0.5px noise floor).
 */
export function distributeLineSlack<T extends DistributeSeg>(
  segments: readonly T[],
  slack: number,
  opts: DistributeOptions = {},
): DistributeResult | null {
  if (Math.abs(slack) <= 0.5) return null;

  const firstContentSi = opts.firstContentSi ?? 0;
  const lastDrawnSi = opts.lastDrawnSi ?? segments.length - 1;
  const minPerGap = opts.minPerGap ?? -Infinity;
  const isGapChar = opts.isGapChar ?? isCjkBreakChar;
  const isWhitespace = opts.isWhitespace ?? defaultIsWhitespace;

  // Flatten the eligible segments to a code-point stream. Segments before
  // `firstContentSi` (leading indent) are skipped; the rest — INCLUDING the
  // final segment — are flattened, because a gap may have its RIGHT side in the
  // final segment even though its left side may not be. Each unit carries its
  // owning segment, the code-point offset within that segment, the code point,
  // and an inter-word-whitespace flag. Non-text atoms become one unit with
  // cp=undefined.
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
      units.push({ si, off, cp, ws: isWhitespace(cp) });
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
  // eligible to OPEN a gap (its si is not lastDrawnSi — the visually-last
  // segment opens none and is never split, but the boundary INTO it still
  // stretches). Whitespace → one inter-word gap (each space stretches).
  // Non-space → an inter-CJK gap only when the boundary to the next NON-space
  // unit satisfies `isGapChar` on either side; a boundary INTO whitespace is
  // already counted by that whitespace, so it is not double-counted.
  const gapAfter = new Array<boolean>(units.length).fill(false);
  let total = 0;
  for (let k = first; k < last; k++) {
    const u = units[k];
    if (u.si === lastDrawnSi) continue; // the visually-last segment opens no gap
    if (u.ws) {
      gapAfter[k] = true;
      total++;
      continue;
    }
    const nx = units[k + 1];
    if (nx.ws) continue; // counted by that space when we reach it
    const lc = u.cp;
    const rc = nx.cp;
    if (
      (lc !== undefined && isGapChar(lc)) ||
      (rc !== undefined && isGapChar(rc))
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
