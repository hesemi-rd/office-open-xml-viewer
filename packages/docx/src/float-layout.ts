// Float-wrap geometry for DOCX anchor images (ECMA-376 §20.4.2.x).
//
// Pure layout math: given the floats active on a page (as FloatRect exclusion
// boxes) it answers "where may this line sit?" and "where must this new float
// be re-seated to avoid a clash?". No canvas/drawing or document-model deps, so
// it can be unit-reasoned and shared by the renderer and the paginator.
//
// IMPORTANT: which parts of the behavior here are ECMA-376-mandated vs
// implementation-defined (Word-mimicking) HEURISTICS is documented inline on
// resolveFloatOverlap and resolveLineFloatWindow. Do not "tighten" the
// heuristics toward a specific sample — see packages/docx/CLAUDE.md.

/** Anchor image float that affects text wrap on the current page. */
export interface FloatRect {
  /** What kind of object reserved this float. Used to scope overlap avoidance:
   *  ECMA-376 §17.4.56 (tblOverlap="never") only forbids a floating table from
   *  overlapping OTHER FLOATING TABLES — not DrawingML anchors (§20.4.2.3) or
   *  text frames. resolveFloatOverlap reads this to limit a never-overlap
   *  table's blockers to kind==='table'. 'shape' = DrawingML wp:anchor shape,
   *  'frame' = <w:framePr> text frame; both also cover anchor images. */
  kind: 'table' | 'shape' | 'frame';
  mode: 'square' | 'topAndBottom';
  /** Hex key of the image bitmap (used to defer drawing until final Y is known). */
  imageKey: string;
  /** Absolute canvas X of the image box (without dist padding). */
  imageX: number;
  imageY: number;
  imageW: number;
  imageH: number;
  /** Padded exclusion rectangle for text wrap. */
  xLeft: number;
  xRight: number;
  yTop: number;
  yBottom: number;
  /** wrapText: "bothSides" | "left" | "right" | "largest" (only square uses this). */
  side: string;
  /** dist* padding (px) — needed when displacing a float to keep its exclusion
   *  padding when re-seating next to a blocking float (ECMA-376 §20.4.2.x). */
  distLeft: number;
  distRight: number;
  distTop: number;
  distBottom: number;
  /** Identifier of the anchoring paragraph. Used only by the implementation-
   *  defined (HEURISTIC) overlap avoidance under allowOverlap=true: floats with
   *  the SAME paraId never displace each other, different-paragraph floats do.
   *  ECMA-376 does not define this scoping; it mirrors Word/LibreOffice. See
   *  resolveFloatOverlap. */
  paraId: number;
  /** true once the image itself has been drawn (drawn after its paragraph lays out). */
  drawn: boolean;
}

/** A horizontal interval [l, r] in absolute canvas px. */
export interface Gap {
  l: number;
  r: number;
}

// ── Float-layout tolerances (px) ──────────────────────────────────────────────
// Sub-pixel slack used so floating-point coordinate noise (margin/anchor/dist
// arithmetic at the current scale) doesn't read as a real overlap or a real gap.

/** Overlap epsilon: two exclusion rects must overlap by MORE than this to count
 *  as intersecting, so coincident/touching edges (and FP noise) are not a clash. */
export const FLOAT_OVERLAP_EPS = 0.01;

/** Slack added to the page-right edge when testing whether a displaced float
 *  still fits horizontally — a float ending within this many px of the page edge
 *  is treated as fitting (it would otherwise be pushed down by FP rounding).
 *  Looser than FLOAT_OVERLAP_EPS because it guards a half-pixel rounding of a
 *  full-width displacement, not an edge-touch test. */
export const FLOAT_PAGE_RIGHT_SLACK = 0.5;

/** Minimum horizontal space (pt) a free side-gap must have before Word will
 *  START a CONTENT (text / inline-object) line beside a float, rather than
 *  flowing it below the float band. Measured — NOT from ECMA-376, which mandates
 *  no side-gap minimum (§20.4.2.17 only says text wraps around the rectangle;
 *  §17.18.3 `<w:br w:clear>` is the sole spec-mandated flow onto a float-free
 *  region). Word's rule is exactly 1 inch (1440 twips): established from the
 *  fixture set private/sample-19/20/22, Word-exported PDF, pdftotext bbox — a gap
 *  of 70pt flows the line below, a gap of 72pt starts the line beside. For a
 *  content line the threshold is INDEPENDENT of the line's text, of font size
 *  (8/12/24pt all switch at 72pt), and of line spacing (single/1.5/double all
 *  switch at 72pt) — i.e. it is an absolute width, not an em- or
 *  line-height-proportional quantity. See issue #676. Callers convert to px with
 *  `WORD_MIN_LINE_START_PT * scale` (renderer scale is px/pt).
 *
 *  SCOPE — content lines only. An EMPTY paragraph-mark line (a literally-empty or
 *  anchor-only paragraph's pilcrow, no width-bearing content) does NOT obey this
 *  1-inch rule: Word keeps such a mark beside a float whenever the gap can hold
 *  the pilcrow itself, dropping it below only when the gap is narrower than that
 *  (effectively a full-width float band). Grounded from sample-9 p.4 (full-width
 *  band → mark drops below, carrying its wrapNone anchor image) AND sample-12 p.2
 *  (a ~62pt side-gap, below 1 inch, where the nine authoring blank-line marks
 *  after the figure stay BESIDE the float — flowing them below at 1 inch (the
 *  regression #676 introduced) pushed the caption + CONCLUSION onto the next
 *  page). The narrow threshold is the paragraph-mark em; it governs the
 *  literally-empty paths — the paint pass `resolveEmptyMarkTop` and paginator
 *  mirror `flowMarkLine` — plus the anchorHost-only metric line inside
 *  `layoutLines`. Every CONTENT line (including a content paragraph's
 *  trailing-break empty final line) passes `wordMinLineStartPx` below. */
export const WORD_MIN_LINE_START_PT = 72;

/** Tolerance (pt) subtracted from the 1-inch requirement when testing a side
 *  gap, to make Word's INCLUSIVE ≥ 1-inch boundary robust to coordinate noise.
 *  Word places a line beside a float at a gap of exactly 1 inch (issue #676 /
 *  sample-22 page 7: a frame authored so the gap is 72.0pt is beside). But a gap
 *  that is nominally 1 inch is computed as content-width − frame-width through
 *  twip→EMU→px conversions and lands slightly under 72: this renderer computes
 *  71.963716pt for the 72.0pt frame (a ~0.036pt deficit — sub-twip conversion
 *  rounding, not pure IEEE-754). Without tolerance the inclusive boundary
 *  flips to below and disagrees with Word. One twip (1/20 pt = 0.05pt) is the
 *  authoring granularity of a frame width, so a gap short of 1 inch by less
 *  than one twip is treated as exactly 1 inch. One twip covers the observed
 *  0.036pt deficit (a half twip, 0.025pt, would NOT) yet is ≪ the 2pt step
 *  that discriminates the fixtures (70pt stays below, 72pt goes beside), so it
 *  never promotes a genuinely sub-inch gap. Applied in the render's px space
 *  as `× scale` (see resolveLineFloatWindow). Same rationale as
 *  FLOAT_PAGE_RIGHT_SLACK: a tolerance sized to the coordinate-rounding
 *  granularity it absorbs. */
export const LINE_START_GAP_EPS_PT = 0.05; // one twip (1/20 pt)

/** The `requiredWidth` (px) every CONTENT-line caller passes to
 *  `resolveLineFloatWindow` for a line-start probe: Word's 1-inch minimum
 *  side-gap, minus the one-twip rounding tolerance, at the render scale (px/pt).
 *  Single source of truth so the paint pass and both paginator mirrors agree
 *  bit-for-bit on the flow/beside decision. Empty paragraph-mark lines use the
 *  narrower `paragraphMarkEmPx` threshold instead (see WORD_MIN_LINE_START_PT's
 *  SCOPE note). See WORD_MIN_LINE_START_PT and LINE_START_GAP_EPS_PT (issue
 *  #676). */
export function wordMinLineStartPx(scale: number): number {
  return (WORD_MIN_LINE_START_PT - LINE_START_GAP_EPS_PT) * scale;
}

/** Minimum width (px) a free side-gap must have to hold a line start. Internal
 *  defensive floor of `resolveLineFloatWindow`: it floors a zero-width probe so
 *  a `requiredWidth === 0` call still rejects sub-pixel slivers between
 *  full-width floats. In the docx renderer this floor no longer GOVERNS the
 *  line-start decision — every caller now passes `wordMinLineStartPx(scale)`
 *  (≈ 54px even at scale 0.75), far above 1px, so `Math.max` always yields the
 *  1-inch requirement (issue #676). Kept because `resolveLineFloatWindow` is an
 *  exported pure function unit-tested on its own contract: a hypothetical
 *  `requiredWidth = 0` caller must still not wedge a line into a coordinate-noise
 *  sliver. */
export const MIN_LINE_GAP = 1;

export function isWrapFloat(mode?: string | null): boolean {
  return mode === 'square' || mode === 'topAndBottom' || mode === 'tight' || mode === 'through';
}

/**
 * Does float `f`'s horizontal extent overlap the paragraph/column text band
 * [paraXLeft, paraXRight]? Touching edges (within FLOAT_OVERLAP_EPS) do not count.
 *
 * ECMA-376 §20.4.2.17 (wrapSquare) and §20.4.2.20 (wrapTopAndBottom) both exclude
 * text only where the object is horizontally placed ("text shall wrap around …
 * THIS OBJECT"). Floats are registered in ABSOLUTE page coordinates and the page
 * float set is shared across a section's newspaper columns (§17.6.4), so a float
 * anchored in one column must be filtered out for a line laid out in another
 * column that it does not horizontally overlap. Both wrap modes route through
 * this one predicate so they share identical column-scoping semantics.
 */
export function floatOverlapsColumnX(
  f: FloatRect,
  paraXLeft: number,
  paraXRight: number,
): boolean {
  return f.xRight > paraXLeft + FLOAT_OVERLAP_EPS && f.xLeft < paraXRight - FLOAT_OVERLAP_EPS;
}

/** Two exclusion rects intersect (strict overlap, touching edges allowed). */
export function rectsOverlap(
  aL: number, aR: number, aT: number, aB: number,
  bL: number, bR: number, bT: number, bB: number,
): boolean {
  return aL < bR - FLOAT_OVERLAP_EPS && aR > bL + FLOAT_OVERLAP_EPS &&
    aT < bB - FLOAT_OVERLAP_EPS && aB > bT + FLOAT_OVERLAP_EPS;
}

/**
 * Widest free horizontal interval within [left, right] after removing the
 * `blocked` spans. Returns null when nothing is free. Factored out of
 * resolveLineFloatWindow so the caller holds a properly-typed Gap (the previous
 * inline closure form forced TS to narrow `best` to never, requiring casts).
 */
export function widestFreeGap(blocked: Gap[], left: number, right: number): Gap | null {
  const spans = blocked.slice().sort((a, b) => a.l - b.l);
  let cursor = left;
  let best: Gap | null = null;
  const consider = (l: number, r: number): void => {
    // Adopt only when strictly wider than the current best (0 when none yet),
    // so a zero/negative-width gap never becomes `best`. Matches the prior inline
    // form `r - l > (best ? best.r - best.l : 0)`.
    if (r - l > (best ? best.r - best.l : 0)) best = { l, r };
  };
  for (const b of spans) {
    if (b.l > cursor) consider(cursor, Math.min(b.l, right));
    cursor = Math.max(cursor, Math.min(b.r, right));
    if (cursor >= right) break;
  }
  if (cursor < right) consider(cursor, right);
  return best;
}

/**
 * Resolve where a single line box may sit relative to the page's active floats.
 *
 * Given the line's intended top Y and the minimum horizontal width it needs to
 * be placeable (`requiredWidth`), this returns the Y at which the line actually
 * starts plus the horizontal sub-window it may use. Every docx caller passes
 * `wordMinLineStartPx(scale)` for `requiredWidth` — Word's measured 1-inch
 * minimum side-gap less a half-twip rounding tolerance (see that helper,
 * WORD_MIN_LINE_START_PT, and issue #676), applied uniformly to empty
 * paragraph-mark lines and text lines alike.
 *
 * Two ECMA-376 wrap rules are applied, in order:
 *   1. topAndBottom floats (§20.4.2.16): a line intersecting one is pushed below
 *      it — text never sits beside a topAndBottom object.
 *   2. square floats (§20.4.2.17): text wraps around the float's rect + dist
 *      padding. When several squares cover a row we take the WIDEST free
 *      horizontal gap. If no gap is wide enough for `requiredWidth` the line
 *      cannot sit here at all and flows below the obstruction (advance to the
 *      lowest blocking float bottom and re-evaluate). The square/topAndBottom
 *      geometry is spec-defined; the `requiredWidth` gate — how much clear space
 *      a line needs before Word starts it beside a float rather than below —
 *      is NOT in ECMA-376 (§17.18.3 `<w:br w:clear>` is the only spec-mandated
 *      flow onto a float-free region). It is Word's observed 1-inch rule, a
 *      GROUNDED runtime measurement (private/sample-19/20/22 PDF bbox, issue
 *      #676), not a fitted constant: 70pt → below, 72pt → beside, independent of
 *      content / font size / line spacing. A line placed beside the float whose
 *      first word overruns the (≥1-inch) gap is force-broken there — Word breaks
 *      the word rather than refusing the gap (observed "AFTE"/"R-10" wrap); the
 *      caller's over-long-word char-break handles that.
 */
export function resolveLineFloatWindow(
  topY: number,
  requiredWidth: number,
  probeH: number,
  paraX: number,
  maxWidth: number,
  floats: FloatRect[],
  // The paragraph's RAW COLUMN band, distinct from the indented text band
  // [paraX, paraX + maxWidth]. Step 1 (topAndBottom) gates by the COLUMN band —
  // §20.4.2.20 blocks the FULL column where the object sits, including the
  // paragraph's indent margins — while step 2 (square side-gap) keeps gating by
  // the narrower indented text band (§20.4.2.17). Defaults to the indented band
  // so a direct unit caller that has no separate column band stays correct.
  columnXLeftPt: number = paraX,
  columnXRightPt: number = paraX + maxWidth,
): { topY: number; xOffset: number; maxWidth: number } {
  const paraXLeft = paraX;
  const paraXRight = paraX + maxWidth;

  // 1. Keep pushing past any topAndBottom block we sit inside.
  //
  // §20.4.2.20 (wrapTopAndBottom) excludes text only where THIS OBJECT is
  // horizontally placed. A page-scoped float anchored in another newspaper
  // column (§17.6.4) must not push this column's line below an unrelated
  // vertical band, so it is gated by `floatOverlapsColumnX` against the RAW
  // COLUMN band — NOT the indented text band step 2 uses. §20.4.2.20 blocks the
  // whole column where the object sits ("text must wrap around neither side of
  // this object"), so a topAndBottom float in this column's indent margin still
  // pushes an indented paragraph's lines below it even though it does not overlap
  // the narrower indented band. Within a column the object DOES overlap it blocks
  // the full width (topY is advanced to the band bottom, no side gap is computed).
  //
  // ASSUMPTION (M3): step 1 runs ONCE, before the square sweep, and is not
  // re-checked after a square push in step 2. This relies on "no topAndBottom
  // float sits below a square float in the same column band." topAndBottom
  // objects span the full column width where they sit (ECMA-376 §20.4.2.20) and
  // are normally anchored above/below the square-wrapped region, so the square
  // push in step 2 lands in float-free space below the band. If a document ever
  // places a topAndBottom strictly below a square the pushed line could clip into
  // it; the spec-correct fix is to make steps 1+2 a single fixpoint loop. Left as
  // a documented assumption here to preserve current behavior (no observed sample
  // exercises the inverted ordering).
  for (let guard = 0; guard < 16; guard++) {
    const lineBot = topY + probeH;
    let skip: number | null = null;
    for (const f of floats) {
      if (f.mode !== 'topAndBottom') continue;
      if (!floatOverlapsColumnX(f, columnXLeftPt, columnXRightPt)) continue;
      if (lineBot > f.yTop && topY < f.yBottom) {
        skip = skip === null ? f.yBottom : Math.max(skip, f.yBottom);
      }
    }
    if (skip === null) break;
    topY = skip;
  }

  // 2. Horizontal constraint from square floats.
  // A gap must be at least `requiredWidth` wide to host a line start. Every docx
  // caller passes Word's 1-inch rule already reduced by its rounding tolerance,
  // `wordMinLineStartPx(scale)` (= (WORD_MIN_LINE_START_PT − LINE_START_GAP_EPS_PT)
  // × scale). MIN_LINE_GAP floors a degenerate `requiredWidth === 0` probe against
  // sub-pixel slivers; it does not govern when a real caller passes ≥ ~1 inch.
  const usableGap = Math.max(requiredWidth, MIN_LINE_GAP);
  let xOffset = 0;
  let lineMaxWidth = maxWidth;
  for (let guard = 0; guard < 64; guard++) {
    const lineBot = topY + probeH;
    const blocked: { l: number; r: number }[] = [];
    const intersecting: FloatRect[] = [];
    for (const f of floats) {
      if (f.mode !== 'square') continue;
      if (lineBot <= f.yTop || topY >= f.yBottom) continue;
      // §20.4.2.17 excludes text only where the square wrap rectangle overlaps
      // its line area. A page-scoped float in another newspaper column must not
      // turn this column's full width into a sub-inch "side gap" and push the
      // line below an unrelated vertical band. Unlike step 1 (which gates by the
      // raw COLUMN band for a full-column block), the square side-gap math is
      // relative to the actual text band, so it gates by the INDENTED band
      // [paraXLeft, paraXRight].
      if (!floatOverlapsColumnX(f, paraXLeft, paraXRight)) continue;
      intersecting.push(f);
      switch (f.side) {
        // Text may sit only on the LEFT of the float ⇒ everything from the
        // float's left edge to the column right is unavailable.
        case 'left':  blocked.push({ l: f.xLeft, r: paraXRight }); break;
        // Text may sit only on the RIGHT ⇒ column left .. float right blocked.
        case 'right': blocked.push({ l: paraXLeft, r: f.xRight }); break;
        case 'largest':
        case 'bothSides':
        default:      blocked.push({ l: f.xLeft, r: f.xRight }); break;
      }
    }
    if (intersecting.length === 0) {
      xOffset = 0;
      lineMaxWidth = maxWidth;
      break;
    }
    // Widest free horizontal gap between the blocked spans across the column.
    // The caller's `requiredWidth` already carries the 1-inch tolerance
    // (wordMinLineStartPx), so a gap meant to be exactly 1 inch but computed a
    // hair under 72 (twip/px rounding noise) still meets Word's inclusive
    // ≥ 1-inch boundary (issue #676, sample-22 p.7).
    const best = widestFreeGap(blocked, paraXLeft, paraXRight);
    if (best && best.r - best.l >= usableGap) {
      xOffset = Math.max(0, best.l - paraXLeft);
      lineMaxWidth = Math.min(maxWidth - xOffset, best.r - best.l);
      if (lineMaxWidth < 0) lineMaxWidth = 0;
      break;
    }

    // No usable gap on this row: the intersecting floats cover the whole column
    // width here. The line must clear them all — advance to the LOWEST blocking
    // float bottom (max yBottom) so it sits centred below the band rather than
    // squeezed into a side sliver beside a still-active float.
    const nextY = Math.max(...intersecting.map((f) => f.yBottom));
    if (nextY <= topY) {
      // Degenerate guard (shouldn't happen): keep full width to avoid a stall.
      xOffset = 0;
      lineMaxWidth = maxWidth;
      break;
    }
    topY = nextY;
  }
  return { topY, xOffset, maxWidth: lineMaxWidth };
}

/**
 * Multi-float collision resolution for a NEW wrap float, against floats already
 * registered on the page.
 *
 * What ECMA-376 actually mandates here is narrow, and the mandate differs by
 * what kind of object forbids overlap:
 *   - A DrawingML anchor with @allowOverlap="false" (Part 1 §20.4.2.3): an
 *     object that "cannot overlap other DrawingML object … shall be
 *     repositioned when displayed to prevent this overlap" — i.e. it must avoid
 *     OTHER DRAWINGML OBJECTS. (We never pass allowOverlap=false for shapes/
 *     images today; the default is "true", which only *permits* overlap.)
 *   - A floating table with <w:tblOverlap w:val="never"/> (§17.4.56): the table
 *     "cannot overlap with OTHER FLOATING TABLES in the document." It does NOT
 *     mandate avoiding DrawingML anchors (§20.4.2.3) or text frames — those keep
 *     their own §20.4.2.3 behavior. So a never-overlap table must only avoid
 *     blockers with kind==='table'.
 * allowOverlap="true"/omitted (the default, §20.4.2.3 / §17.4.56) only *permits*
 * overlap; the spec is silent on whether a renderer may avoid it. So:
 *   - allowOverlap === false → spec-mandated avoidance. Scoped by `kind`: a
 *     table avoids only other tables (§17.4.56); any other kind would avoid all
 *     (§20.4.2.3) — not currently exercised, see above.
 *   - allowOverlap === true  → implementation-defined avoidance of floats
 *     anchored in OTHER paragraphs only.
 *
 * EVERYTHING ELSE in this function is implementation-defined — ECMA-376 Part 1
 * does NOT specify it. This is a HEURISTIC chosen to match Word's observed
 * layout (e.g. sample-9 figure 9), not a spec requirement:
 *   - the move DIRECTION (right first, then down),
 *   - WHICH float moves (the later/document-order float is the "new" one),
 *   - the "same-paragraph floats never displace each other" gate (the paraId
 *     scoping under allowOverlap=true above),
 *   - and the move AMOUNT. Note the dist* padding reused below is, per
 *     §20.4.2.3/§20.4.2.17, the minimum distance between the float and *text*
 *     (wrapSquare geometry) — it is NOT spec-defined as a float-to-float gap.
 *     Using it to seat one float beside another is our own choice.
 *
 * If the §20.4.2.3 "shall be repositioned" requirement is ever satisfiable in
 * more than one way, the particular re-seating below remains a Word-mimicking
 * heuristic; keep it as such until a spec-grounded placement rule is found.
 *
 * We re-seat horizontally to the right of the blocking float(s) first (margins
 * may be used — Word lets a displaced float sit in the page margin), and only
 * fall back to a vertical push when no horizontal room remains.
 *
 * Coordinates are page-absolute px. (x,y) is the image box origin (no dist).
 * `pageRight` is the page width in px; `floats` is the page's active float set.
 */
export function resolveFloatOverlap(
  x: number, y: number, w: number, h: number,
  dl: number, dr: number, dt: number, db: number,
  paraId: number, allowOverlap: boolean,
  kind: FloatRect['kind'],
  pageRight: number, floats: FloatRect[],
): { x: number; y: number } {
  for (let guard = 0; guard < 16; guard++) {
    const exL = x - dl, exR = x + w + dr, exT = y - dt, exB = y + h + db;
    // Which already-registered floats block the moving float:
    //   - allowOverlap === false (spec-mandated avoidance): scope by the moving
    //     float's kind. A floating table with tblOverlap="never" (§17.4.56) may
    //     only overlap-avoid OTHER FLOATING TABLES, so it blocks on kind==='table'
    //     alone; DrawingML anchors / frames keep their own §20.4.2.3 placement.
    //     Any other kind would avoid all intersecting floats (§20.4.2.3), but no
    //     caller passes allowOverlap=false for non-tables today.
    //   - allowOverlap === true (Word de-facto cross-paragraph avoidance): only
    //     floats anchored in OTHER paragraphs block, regardless of kind.
    const blockers = floats.filter(
      (f) =>
        (allowOverlap ? f.paraId !== paraId : kind !== 'table' || f.kind === 'table') &&
        rectsOverlap(exL, exR, exT, exB, f.xLeft, f.xRight, f.yTop, f.yBottom),
    );
    if (blockers.length === 0) return { x, y };

    // Horizontal: re-seat just right of the right-most blocker. Setting our
    // left exclusion edge (x - dl) flush against the blocker's right exclusion
    // edge means x = maxRight + dl (gap = blocker.distRight + our distLeft).
    const maxRight = Math.max(...blockers.map((f) => f.xRight));
    const newX = maxRight + dl;
    if (newX + w + dr <= pageRight + FLOAT_PAGE_RIGHT_SLACK) {
      x = newX;
      continue; // re-check against all other-paragraph floats
    }

    // No horizontal room: push below the lowest blocker (smallest displacement
    // that clears them). Our top exclusion edge (y - dt) flush with the
    // blocker's bottom exclusion edge ⇒ y = maxBottom + dt.
    const maxBottom = Math.max(...blockers.map((f) => f.yBottom));
    y = maxBottom + dt;
  }
  return { x, y };
}

/**
 * If y is inside a topAndBottom float that horizontally overlaps the paragraph's
 * column band [paraXLeft, paraXRight], return that float's bottom; otherwise
 * return y. Mirrors `resolveLineFloatWindow` step 1: §20.4.2.20 excludes text
 * only where THIS OBJECT is placed, so a float anchored in another newspaper
 * column (§17.6.4) — the page float set is shared across columns — is filtered
 * out via the shared `floatOverlapsColumnX` predicate.
 */
export function skipPastTopAndBottom(
  y: number,
  floats: FloatRect[],
  paraXLeft: number,
  paraXRight: number,
): number {
  for (let guard = 0; guard < 16; guard++) {
    let next = y;
    for (const f of floats) {
      if (f.mode !== 'topAndBottom') continue;
      if (!floatOverlapsColumnX(f, paraXLeft, paraXRight)) continue;
      if (y >= f.yTop && y < f.yBottom) next = Math.max(next, f.yBottom);
    }
    if (next === y) return y;
    y = next;
  }
  return y;
}
