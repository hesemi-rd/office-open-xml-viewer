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
// ranges named in the issue. Myanmar/Tibetan and the no-space Thai↔Latin/CJK
// script-transition boundary are intentionally out of scope (follow-ups).

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
