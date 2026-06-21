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

/** Minimum width (px) a free side-gap must have to hold a line start. Floors the
 *  required-width probe so a zero-width probe (an empty line with no content yet)
 *  still rejects sub-pixel slivers between full-width floats. */
export const MIN_LINE_GAP = 1;

export function isWrapFloat(mode?: string | null): boolean {
  return mode === 'square' || mode === 'topAndBottom' || mode === 'tight' || mode === 'through';
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
 * be placeable (`requiredWidth` — the width of its first atomic token, or, for
 * an empty paragraph-mark line, one em of the mark font), this returns the Y at
 * which the line actually starts plus the horizontal sub-window it may use.
 *
 * Two ECMA-376 wrap rules are applied, in order:
 *   1. topAndBottom floats (§20.4.2.16): a line intersecting one is pushed below
 *      it — text never sits beside a topAndBottom object.
 *   2. square floats (§20.4.2.17): text wraps around the float's rect + dist
 *      padding. When several squares cover a row we take the WIDEST free
 *      horizontal gap. If no gap is wide enough for `requiredWidth` the line
 *      cannot sit here at all and flows below the obstruction (advance to the
 *      lowest blocking float bottom and re-evaluate). The square/topAndBottom
 *      geometry itself is spec-defined; but routing empty / anchor-only
 *      paragraph-mark lines through this with `requiredWidth` = one em (so they
 *      drop below a full-width float band) is an implementation-defined
 *      HEURISTIC (see resolveEmptyMarkTop): ECMA-376 does not require a
 *      float-free row for a paragraph mark — only `<w:br w:clear>` (§17.18.3)
 *      mandates flowing onto a float-free region.
 */
export function resolveLineFloatWindow(
  topY: number,
  requiredWidth: number,
  probeH: number,
  paraX: number,
  maxWidth: number,
  floats: FloatRect[],
): { topY: number; xOffset: number; maxWidth: number } {
  // 1. Keep pushing past any topAndBottom block we sit inside.
  //
  // ASSUMPTION (M3): step 1 runs ONCE, before the square sweep, and is not
  // re-checked after a square push in step 2. This relies on "no topAndBottom
  // float sits below a square float in the same column band." topAndBottom
  // objects span the full content width (ECMA-376 §20.4.2.16) and are normally
  // anchored above/below the square-wrapped region, so the square push in step 2
  // lands in float-free space below the band. If a document ever places a
  // topAndBottom strictly below a square the pushed line could clip into it; the
  // spec-correct fix is to make steps 1+2 a single fixpoint loop. Left as a
  // documented assumption here to preserve current behavior (no observed sample
  // exercises the inverted ordering).
  for (let guard = 0; guard < 16; guard++) {
    const lineBot = topY + probeH;
    let skip: number | null = null;
    for (const f of floats) {
      if (f.mode !== 'topAndBottom') continue;
      if (lineBot > f.yTop && topY < f.yBottom) {
        skip = skip === null ? f.yBottom : Math.max(skip, f.yBottom);
      }
    }
    if (skip === null) break;
    topY = skip;
  }

  // 2. Horizontal constraint from square floats.
  const paraXLeft = paraX;
  const paraXRight = paraX + maxWidth;
  // A gap must fit the line's first atomic token to be usable. Floor so a
  // zero-width probe (no content yet) still rejects sub-pixel slivers.
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
 * What ECMA-376 actually mandates here is narrow. Part 1 §20.4.2.3
 * (wp:anchor/@allowOverlap) says an object that "cannot overlap other DrawingML
 * object … shall be repositioned when displayed to prevent this overlap" — i.e.
 * allowOverlap="false" REQUIRES repositioning. allowOverlap="true" (the default
 * when omitted, §20.4.2.3) only *permits* overlap; the spec is silent on whether
 * a renderer may avoid it. So:
 *   - allowOverlap === false → spec-mandated avoidance of ALL other floats.
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
  pageRight: number, floats: FloatRect[],
): { x: number; y: number } {
  for (let guard = 0; guard < 16; guard++) {
    const exL = x - dl, exR = x + w + dr, exT = y - dt, exB = y + h + db;
    // Blocking floats whose exclusion rects intersect ours. When the moving
    // float forbids overlap (§20.4.2.3 allowOverlap="false") EVERY intersecting
    // float blocks; otherwise only floats from OTHER paragraphs block (Word
    // de-facto cross-paragraph avoidance).
    const blockers = floats.filter(
      (f) => (allowOverlap ? f.paraId !== paraId : true) &&
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

/** If y is inside a topAndBottom float, return the float bottom; otherwise return y. */
export function skipPastTopAndBottom(y: number, floats: FloatRect[]): number {
  for (let guard = 0; guard < 16; guard++) {
    let next = y;
    for (const f of floats) {
      if (f.mode !== 'topAndBottom') continue;
      if (y >= f.yTop && y < f.yBottom) next = Math.max(next, f.yBottom);
    }
    if (next === y) return y;
    y = next;
  }
  return y;
}
