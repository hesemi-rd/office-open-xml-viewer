// Dictionary-based line breaking for Southeast-Asian (SEA) scripts that write
// without inter-word spaces: Thai, Lao and Khmer (GitHub issue #797). Word
// boundaries in these scripts are not marked by whitespace and cannot be derived
// from Unicode ranges alone; they require a dictionary. We use the platform ICU
// dictionary via `Intl.Segmenter({ granularity: 'word' })` to enumerate the
// break opportunities INTERIOR to each SEA-script span, and expose a pure,
// package-agnostic kernel that every renderer's wrap loop consumes.
//
// Design invariants (shared with the docx/pptx/xlsx breakers):
//   • ADD break opportunities only; never change glyphs or advances. The run is
//     split only at the ACTUAL line break, so within a line the text stays one
//     contiguous draw (measure==paint).
//   • Restrict added opportunities to boundaries INTERIOR to a maximal SEA span,
//     between two dictionary WORDS. A boundary that touches non-SEA text (Latin,
//     digits, CJK, whitespace) is left to the existing whitespace/Latin/CJK
//     logic, so non-SEA content stays byte-identical.
//   • Graceful fallback: when `Intl.Segmenter` is unavailable (old runtimes, a
//     `small-icu` build without the SEA dictionaries) or throws, we return no
//     offsets and the caller keeps its current cluster/character behaviour.
//
// Scope: Thai U+0E00–0E7F, Lao U+0E80–0EFF, Khmer U+1780–17FF — exactly the
// ranges named in the issue. Myanmar/Tibetan remain out of scope (follow-up
// #961). The no-space SEA↔Latin/CJK/digit script-transition boundary (#960) is
// added by {@link seaTransitionOffsets}; {@link seaMixedBreakOffsets} unions the
// dictionary, transition and CJK per-character opportunities for a single run
// that mixes SEA with Latin/digits/CJK.

import { isCjkBreakChar } from './cjk-ranges.js';

/** Southeast-Asian dictionary-break script tags used to pick the ICU dictionary.
 *  ICU dispatches the dictionary by the character's SCRIPT, not by this locale,
 *  so it is only a hint; we still pick per-span for clarity/forward-safety. */
export type SeaScript = 'th' | 'lo' | 'km';

/**
 * True when `cp` belongs to one of the no-inter-word-space SEA scripts (Thai,
 * Lao, Khmer). This is a SCRIPT-membership test, not a break predicate: the
 * blocks include combining marks, tone marks, digits and punctuation that are
 * NOT independently breakable — the actual break opportunities come from
 * {@link seaWordBreakOffsets}. Used to detect SEA spans and as the cheap gate in
 * {@link containsSeaScript}.
 *
 * @param cp A Unicode scalar value (e.g. from `String.prototype.codePointAt`).
 */
export function isSeaScriptCodePoint(cp: number): boolean {
  return (
    (cp >= 0x0e00 && cp <= 0x0e7f) || // Thai
    (cp >= 0x0e80 && cp <= 0x0eff) || // Lao
    (cp >= 0x1780 && cp <= 0x17ff) // Khmer
  );
}

/**
 * True when `cp` is a Southeast-Asian (Thai/Lao/Khmer) combining mark that
 * EXTENDS the preceding grapheme cluster rather than starting a new one — i.e.
 * its UAX#29 Grapheme_Cluster_Break is `Extend` or `SpacingMark` (above/below
 * vowel signs, tone marks, the Khmer COENG subscript-former, etc.). These are
 * exactly the code points a `thaiDistribute` justifier must NOT open a gap
 * before, so a base consonant keeps its marks glued.
 *
 * For a boundary between two SEA code points this predicate is sufficient to
 * decide grapheme clustering: the SEA blocks contain no `Prepend` characters
 * (verified against UCD), so a break falls before every SEA code point EXCEPT an
 * Extend/SpacingMark. Ranges derived from the Unicode Character Database
 * (Grapheme_Cluster_Break), cross-checked against `Intl.Segmenter` grapheme
 * output — so this stays correct even where the platform `Intl.Segmenter` is
 * unavailable (the {@link graphemeClusterOffsets} code-point fallback is NOT
 * cluster-safe and must not gate SEA distribution).
 */
export function isSeaGraphemeExtend(cp: number): boolean {
  return (
    // Thai (U+0E00–0E7F): MAI HAN-AKAT; SARA AM..PHINTHU; MAITAIKHU..YAMAKKAN.
    cp === 0x0e31 ||
    (cp >= 0x0e33 && cp <= 0x0e3a) ||
    (cp >= 0x0e47 && cp <= 0x0e4e) ||
    // Lao (U+0E80–0EFF): MAI KAN; SIGN PALI VIRAMA..semivowel signs; tone marks.
    cp === 0x0eb1 ||
    (cp >= 0x0eb3 && cp <= 0x0ebc) ||
    (cp >= 0x0ec8 && cp <= 0x0ece) ||
    // Khmer (U+1780–17FF): vowel signs, COENG, robat, signs (17B4–17D3) + ATTHACAN.
    (cp >= 0x17b4 && cp <= 0x17d3) ||
    cp === 0x17dd
  );
}

/** The SEA script of a code point already known to be SEA (see
 *  {@link isSeaScriptCodePoint}). Used to pick the per-span ICU dictionary. */
function seaScriptOf(cp: number): SeaScript {
  if (cp <= 0x0e7f) return 'th';
  if (cp <= 0x0eff) return 'lo';
  return 'km';
}

/**
 * Cheap presence gate: true when any code point of `text` is SEA-script. Used by
 * the renderers to skip all segmentation work for the overwhelmingly common
 * non-SEA input (zero `Intl.Segmenter` cost), mirroring docx's
 * `hasCJKBreakOpportunity`.
 */
export function containsSeaScript(text: string): boolean {
  for (const ch of text) {
    if (isSeaScriptCodePoint(ch.codePointAt(0)!)) return true;
  }
  return false;
}

// ── Segmenter plumbing (cached + test-injectable) ────────────────────────────

/** Minimal shape of one `Intl.Segments` entry we rely on. */
interface SeaSegment {
  index: number;
  isWordLike?: boolean;
}
/** A word segmenter: text (a single SEA span) → its word segments. Injected in
 *  tests for determinism / to exercise the unavailable + throwing paths. */
export type SeaWordSegmenter = (text: string, script: SeaScript) => Iterable<SeaSegment>;

// `undefined` = use the platform ICU segmenter; `null` = force unavailable
// (test); a function = use it (test).
let wordSegmenterOverride: SeaWordSegmenter | null | undefined;

const intlWordCache = new Map<SeaScript, Intl.Segmenter | null>();

function intlWordSegmenter(script: SeaScript): Intl.Segmenter | null {
  const cached = intlWordCache.get(script);
  if (cached !== undefined) return cached;
  let seg: Intl.Segmenter | null = null;
  try {
    if (typeof Intl !== 'undefined' && typeof Intl.Segmenter === 'function') {
      seg = new Intl.Segmenter(script, { granularity: 'word' });
    }
  } catch {
    seg = null;
  }
  intlWordCache.set(script, seg);
  return seg;
}

function wordSegmenterFor(script: SeaScript): ((span: string) => Iterable<SeaSegment>) | null {
  if (wordSegmenterOverride === null) return null; // forced unavailable (test)
  if (typeof wordSegmenterOverride === 'function') {
    const fn = wordSegmenterOverride;
    return (span) => fn(span, script);
  }
  const intl = intlWordSegmenter(script);
  return intl ? (span) => intl.segment(span) as Iterable<SeaSegment> : null;
}

/** Test seam: override the word segmenter. `null` forces the unavailable path;
 *  a function supplies deterministic segments; call {@link resetSeaSegmenterForTest}
 *  to restore the platform ICU segmenter. Not for production use. */
export function setSeaWordSegmenterForTest(fn: SeaWordSegmenter | null): void {
  wordSegmenterOverride = fn;
}
/** Test seam: restore the default platform ICU segmenter and clear caches. */
export function resetSeaSegmenterForTest(): void {
  wordSegmenterOverride = undefined;
  intlWordCache.clear();
  graphemeSegmenter = undefined;
}

// ── Word break offsets ───────────────────────────────────────────────────────

/**
 * Enumerate dictionary word-break opportunities in `text`, as UTF-16 offsets `i`
 * such that a line break is permitted immediately BEFORE `text[i]`.
 *
 * Only boundaries that are (1) interior to a maximal SEA-script span and (2) the
 * START of a word-like segment are returned — so we never break before intra-SEA
 * punctuation (Khmer `។`, Thai `๚`) nor at a SEA↔non-SEA transition (left to the
 * existing logic). The offsets are always at grapheme-cluster boundaries (ICU
 * word breaks never split a base + combining mark).
 *
 * Returns `[]` (graceful fallback) when `text` has no SEA character, when
 * `Intl.Segmenter` is unavailable, or when segmentation throws. Segment ONCE per
 * logical string and cache/slice the result; do not call this per line.
 */
export function seaWordBreakOffsets(text: string): number[] {
  if (!containsSeaScript(text)) return [];
  const offsets: number[] = [];
  const len = text.length;
  let i = 0;
  while (i < len) {
    const cp = text.codePointAt(i)!;
    const step = cp > 0xffff ? 2 : 1;
    if (!isSeaScriptCodePoint(cp)) {
      i += step;
      continue;
    }
    // Maximal SEA span [i, j). The span's dictionary is picked from its first
    // character; ICU still dispatches per-script internally within the span.
    const script = seaScriptOf(cp);
    let j = i + step;
    while (j < len) {
      const c = text.codePointAt(j)!;
      if (!isSeaScriptCodePoint(c)) break;
      j += c > 0xffff ? 2 : 1;
    }
    const segFn = wordSegmenterFor(script);
    if (segFn) {
      const span = text.slice(i, j);
      // Accumulate this span's offsets in a temp buffer and commit them only
      // after the iterator completes, so a mid-iteration `segment()` failure
      // leaves NO partial offsets (all-or-nothing graceful fallback per span).
      const spanOffsets: number[] = [];
      try {
        for (const g of segFn(span)) {
          // index>0 keeps the offset interior to the span (never span-start,
          // which is a SEA↔non-SEA edge or the string start). isWordLike keeps
          // it a break-before-a-WORD (never before punctuation). Fake segmenters
          // may omit isWordLike → treat as word-like.
          if (g.index > 0 && (g.isWordLike ?? true)) spanOffsets.push(i + g.index);
        }
        for (const o of spanOffsets) offsets.push(o);
      } catch {
        // segment() failure on this span → add no offsets (graceful fallback).
      }
    }
    i = j;
  }
  return offsets;
}

// ── Script-transition break opportunities (issue #960) ───────────────────────

/**
 * Break-before offsets at the EDGES of every maximal SEA span — the SEA↔non-SEA
 * script transitions that carry NO intervening whitespace (issue #960). Word
 * treats these transitions as line-break opportunities (adjudicated against the
 * Word-exported ground truth sample-45.pdf: a line broke exactly at the Thai→
 * Latin boundary `…ของ | Thailand`, carrying the Latin word whole to the next
 * line; digit groups `1250`/`990` break away from the surrounding Thai the same
 * way). {@link seaWordBreakOffsets} deliberately restricts itself to boundaries
 * INTERIOR to a SEA span, so the transition edges were previously unbreakable —
 * a spaceless mix like `เมืองBangkok…Thailand` had no legal break at the script
 * seam. This adds exactly those seams.
 *
 * Returns UTF-16 offsets `i` (a break is permitted immediately BEFORE `text[i]`),
 * strictly interior (`0 < i < text.length`), ascending and unique. A transition
 * where either side is whitespace is skipped — a space is already an inter-word
 * break, so the seam beside it needs no extra opportunity. Returns `[]` when
 * `text` has no SEA character (byte-identical non-SEA path).
 */
export function seaTransitionOffsets(text: string): number[] {
  if (!containsSeaScript(text)) return [];
  const offsets: number[] = [];
  const len = text.length;
  let prevCp = text.codePointAt(0)!;
  let i = prevCp > 0xffff ? 2 : 1;
  while (i < len) {
    const cp = text.codePointAt(i)!;
    const step = cp > 0xffff ? 2 : 1;
    const prevSea = isSeaScriptCodePoint(prevCp);
    const curSea = isSeaScriptCodePoint(cp);
    // A SEA↔non-SEA change is a transition. Skip when either side is whitespace
    // (the space itself is the break opportunity, handled by the caller).
    if (prevSea !== curSea && !isBreakInertSpaceCp(prevCp) && !isBreakInertSpaceCp(cp)) {
      offsets.push(i);
    }
    prevCp = cp;
    i += step;
  }
  return offsets;
}

/** Code points beside which a SEA↔non-SEA seam must NOT gain a break: real
 *  whitespace (the space itself already IS the break) and the non-breaking space
 *  family (NBSP U+00A0, figure space U+2007, narrow NBSP U+202F, word joiner
 *  U+2060, ZWNBSP/BOM U+FEFF), which are explicitly NON-breaking (UAX#14 GL/WJ). */
function isBreakInertSpaceCp(cp: number): boolean {
  return (
    cp === 0x20 || cp === 0x09 || cp === 0x0a || cp === 0x0d || cp === 0x3000 ||
    cp === 0x00a0 || cp === 0x2007 || cp === 0x202f || cp === 0x2060 || cp === 0xfeff
  );
}

/** Kinsoku sets used to drop line-break positions that would leave a
 *  line-start-forbidden char at a line head or a line-end-forbidden char at a
 *  line tail (ECMA-376 §17.15.1.58–.60). Shape-compatible with the renderers'
 *  `KinsokuRules`; only the two membership sets are consulted. */
export interface SeaMixedKinsoku {
  /** When `false`, the document turned kinsoku OFF (§17.3.1.16 `<w:kinsoku
   *  w:val="0"/>`) and NO position is dropped — matching the CJK path's own
   *  `if (!rules.enabled) return` short-circuit. Absent/`true` ⇒ filter. */
  enabled?: boolean;
  lineStartForbidden: ReadonlySet<number>;
  lineEndForbidden: ReadonlySet<number>;
}

/**
 * The UNIFIED break-before offset set for a single run/segment that contains SEA
 * script but may also mix Latin, digits and CJK (issue #960). Unions three
 * opportunity sources so each script keeps its own rule inside one contiguous
 * token:
 *   • SEA dictionary word boundaries — {@link seaWordBreakOffsets};
 *   • SEA↔non-SEA transitions — {@link seaTransitionOffsets};
 *   • CJK per-character boundaries when `opts.cjk` — every position whose char,
 *     or the char before it, is a {@link isCjkBreakChar} ideograph, so the CJK
 *     side breaks at each character exactly as its dedicated path does.
 *
 * When `opts.kinsoku` is supplied, positions that would orphan a
 * line-start-forbidden char at a line head, or leave a line-end-forbidden char
 * at a line tail, are removed (the offset-set equivalent of the CJK path's
 * retract). This is what lets a mixed CJK+SEA token — which previously fell
 * wholly onto the CJK path and so treated its SEA interior as unbreakable — wrap
 * with BOTH the CJK per-char opportunities and the SEA dictionary/transition
 * opportunities merged.
 *
 * Offsets are ascending, unique, strictly interior. For a pure-SEA token with no
 * kinsoku char this equals {@link seaWordBreakOffsets} (byte-identical wrap).
 */
export function seaMixedBreakOffsets(
  text: string,
  opts?: { cjk?: boolean; kinsoku?: SeaMixedKinsoku },
): number[] {
  if (!containsSeaScript(text)) return [];
  const len = text.length;
  // Dictionary boundaries are ALWAYS grapheme-cluster boundaries (ICU never
  // splits a base + mark), so seed the set with them directly.
  const set = new Set<number>(seaWordBreakOffsets(text));
  // The transition and CJK offsets are enumerated per code point, so they could
  // land INSIDE a grapheme cluster (a base + variation selector / ZWJ / tone
  // mark). Collect them separately and keep only those that coincide with a
  // grapheme boundary — a code-point split there would tear the cluster.
  const graphemeSafe = new Set<number>(graphemeClusterOffsets(text));
  const addGraphemeSafe = (o: number): void => {
    if (graphemeSafe.has(o)) set.add(o);
  };
  for (const o of seaTransitionOffsets(text)) addGraphemeSafe(o);
  if (opts?.cjk) {
    let prevCp = text.codePointAt(0)!;
    let i = prevCp > 0xffff ? 2 : 1;
    while (i < len) {
      const cp = text.codePointAt(i)!;
      // A break may fall before an ideograph or after one (CJK is break-eligible
      // on both edges; the grapheme test above drops a variation-selector tear,
      // and the kinsoku pass below removes the illegal ones).
      if (isCjkBreakChar(cp) || isCjkBreakChar(prevCp)) addGraphemeSafe(i);
      prevCp = cp;
      i += cp > 0xffff ? 2 : 1;
    }
  }
  const k = opts?.kinsoku;
  // §17.3.1.16 — a document that turned kinsoku OFF drops NO position (mirrors
  // the CJK split path's `if (!rules.enabled) return`).
  const filterKinsoku = k != null && k.enabled !== false;
  const out: number[] = [];
  for (const o of set) {
    if (o <= 0 || o >= len) continue;
    if (filterKinsoku) {
      const at = text.codePointAt(o)!;
      const before = codePointEndingAt(text, o);
      if (before !== undefined && k!.lineEndForbidden.has(before)) continue;
      if (k!.lineStartForbidden.has(at)) continue;
    }
    out.push(o);
  }
  out.sort((a, b) => a - b);
  return out;
}

/** The code point that ENDS immediately before UTF-16 offset `o` (handles a
 *  trailing surrogate pair). `undefined` when `o` is at or before the start. */
function codePointEndingAt(text: string, o: number): number | undefined {
  if (o <= 0) return undefined;
  const lo = text.charCodeAt(o - 1);
  if (lo >= 0xdc00 && lo <= 0xdfff && o >= 2) {
    const hi = text.charCodeAt(o - 2);
    if (hi >= 0xd800 && hi <= 0xdbff) return text.codePointAt(o - 2)!;
  }
  return lo;
}

// ── Greedy whole-word line fit (shared kernel) ───────────────────────────────

/**
 * Greedy whole-word line fit for SEA dictionary breaking, shared by all three
 * renderers so the break decision is defined once.
 *
 * Returns the largest `end` with `start < end <= text.length` such that
 *   (a) `end === text.length` OR `end` is a member of `offsets` (a legal word
 *       boundary), AND
 *   (b) `measure(text.slice(start, end)) <= avail`.
 * If not even the first candidate word fits within `avail`, returns `start`
 * (no progress) so the caller can wrap to a fresh line first, or — on an already
 * empty line — fall back to a grapheme-safe emergency split.
 *
 * Every candidate is measured (no early return): advance width is NOT guaranteed
 * monotone in the prefix length because Word/PowerPoint allow NEGATIVE character
 * spacing (`w:spacing` / `a:spc`), so a longer prefix can fit after a shorter one
 * overflowed. We keep the largest boundary (or the full remainder) that fits.
 *
 * `offsets` MUST be sorted ascending (as produced by {@link seaWordBreakOffsets}
 * or {@link graphemeClusterOffsets}). `measure` supplies the caller's exact
 * advance model (so measure==paint). Pure — no canvas/DOM.
 */
export function fitSeaWordPrefix(
  text: string,
  offsets: readonly number[],
  start: number,
  avail: number,
  measure: (sub: string) => number,
): number {
  const len = text.length;
  if (start >= len) return start;
  let best = start;
  for (const b of offsets) {
    if (b <= start || b >= len) continue;
    if (measure(text.slice(start, b)) <= avail) best = b;
  }
  if (measure(text.slice(start, len)) <= avail) best = len;
  return best;
}

// ── Grapheme clusters (emergency over-long-word split) ───────────────────────

let graphemeSegmenter: Intl.Segmenter | null | undefined;

function getGraphemeSegmenter(): Intl.Segmenter | null {
  if (graphemeSegmenter !== undefined) return graphemeSegmenter;
  let seg: Intl.Segmenter | null = null;
  try {
    if (typeof Intl !== 'undefined' && typeof Intl.Segmenter === 'function') {
      seg = new Intl.Segmenter(undefined, { granularity: 'grapheme' });
    }
  } catch {
    seg = null;
  }
  graphemeSegmenter = seg;
  return seg;
}

/**
 * Grapheme-cluster break-before offsets (UTF-16, interior — excludes 0 and
 * `text.length`). Used for the last-resort emergency split of a single SEA word
 * wider than the whole line/cell: a code-point split would tear a base +
 * combining/tone mark (both are BMP, so "SEA is BMP therefore safe" is false),
 * whereas a grapheme boundary keeps the cluster intact. Falls back to code-point
 * boundaries when `Intl.Segmenter` is unavailable.
 */
export function graphemeClusterOffsets(text: string): number[] {
  const out: number[] = [];
  const seg = getGraphemeSegmenter();
  if (seg) {
    for (const g of seg.segment(text)) {
      if (g.index > 0) out.push(g.index);
    }
    return out;
  }
  for (let i = 0; i < text.length; ) {
    const cp = text.codePointAt(i)!;
    i += cp > 0xffff ? 2 : 1;
    if (i < text.length) out.push(i);
  }
  return out;
}
