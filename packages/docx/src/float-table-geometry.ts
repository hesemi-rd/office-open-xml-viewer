// Floating-table placement geometry (ECMA-376 §17.4.57 `<w:tblpPr>` /
// §17.4.56 `<w:tblOverlap>`).
//
// Pure placement math: given a `<w:tblpPr>` and the section geometry on
// `RenderState`, resolve the table box (canvas px), decide which side body text
// wraps on, and push the wrap-exclusion FloatRect onto `state.floats`. The
// anchor / alignment semantics line up 1:1 with a `<w:framePr>` text frame, so
// this module reuses frame-geometry's frameXContainer / frameYContainer /
// resolveAlignedPosH / resolveAlignedPosV and the shared pushFloatRect builder.
// Extracted from renderer.ts so the placement logic can be unit-reasoned in
// isolation (see float-table-geometry.test.ts / measure-column-geometry.test.ts).
// Only `RenderState` is imported as a type (erased at runtime), so there is no
// import cycle with renderer.ts.

import type { TblpPr } from './types.js';
import type { RenderState } from './renderer.js';
import {
  frameXContainer,
  frameYContainer,
  resolveAlignedPosH,
  resolveAlignedPosV,
  clampAbsBoxIntoContainer,
  pushFloatRect,
} from './frame-geometry.js';

/** Resolved top-left placement (canvas px) of a floating table (`<w:tblpPr>`,
 *  ECMA-376 §17.4.57). `w`/`h` are the rendered table extent (sum of column
 *  widths × row heights). Exported for unit tests only — not package API. */
export interface FloatTableBox {
  x: number;
  y: number;
  w: number;
  h: number;
}

/**
 * Resolve a floating table's top-left placement (canvas px) from its
 * `<w:tblpPr>` (ECMA-376 §17.4.57). `tableW`/`tableH` are the already-laid-out
 * table extent in px. The anchor / alignment semantics line up 1:1 with a
 * `<w:framePr>` text frame (horzAnchor↔hAnchor, vertAnchor↔vAnchor,
 * tblpXSpec↔xAlign, tblpYSpec↔yAlign, tblpX/tblpY↔x/y), so this mirrors
 * {@link computeFrameBox}'s placement math exactly — `frameXContainer` /
 * `frameYContainer` give the same per-anchor bands (text→column band,
 * margin→page content margin, page→physical edges), which is the #513
 * column-integrity guarantee.
 *
 * Exported for unit tests only (the float-table-geometry table) — not package API.
 */
export function computeFloatTableBox(
  tp: TblpPr,
  state: RenderState,
  paraTop: number,
  tableW: number,
  tableH: number,
): FloatTableBox {
  const sc = state.scale;
  // §17.4.57 + §17.18.35: horzAnchor's literal default is "page", which would
  // pin a floating table to the physical page edge (left margin). But when the
  // source `<w:tblpPr>` gave NO horizontal positioning at all (no horzAnchor, no
  // tblpX, no tblpXSpec), Word anchors the table at the anchor paragraph's text/
  // column left — the in-flow position it was converted from — NOT the page edge.
  // (Word runtime behavior; the spec-literal page default does not match Word
  // here. See calibre sample-11 p.3 "ITEM/NEEDED" float.) We force the text band
  // for that case so the table aligns with the body column, then let tblpX=0
  // place it at that band's left.
  const hx = tp.horzSpecified
    ? frameXContainer(tp.horzAnchor, state)
    : frameXContainer('text', state);
  // Vertical band of the vertAnchor (the "anchor object", §22.9.2.20). The
  // "text" band end uses the table height; tblpYSpec is gated out for "text" so
  // only band.start (= paraTop) is consumed there.
  const vBand = frameYContainer(tp.vertAnchor, paraTop, tableH, state);

  // Horizontal: tblpXSpec (ST_XAlign) supersedes the absolute tblpX offset
  // (§17.4.57). Mirrors computeFrameBox's xAlign handling.
  let x: number;
  if (tp.tblpXSpec) {
    x = resolveAlignedPosH(tp.tblpXSpec, hx.left, hx.right, tableW);
  } else {
    // §17.4.57 tblpX: absolute signed offset from the horzAnchor left edge.
    x = hx.left + tp.tblpX * sc;
  }

  // Vertical: tblpYSpec (ST_YAlign) supersedes the absolute tblpY offset
  // (§17.4.57) — EXCEPT when vertAnchor="text", where relative vertical
  // positioning is not allowed and tblpYSpec is ignored (fall back to tblpY).
  // Mirrors computeFrameBox's yAlign handling (ignored when vAnchor="text").
  let y: number;
  if (tp.tblpYSpec && tp.vertAnchor !== 'text') {
    y = resolveAlignedPosV(tp.tblpYSpec, vBand, tableH);
  } else {
    // §17.4.57 tblpY: absolute signed offset from the vertAnchor band start.
    y = vBand.start + tp.tblpY * sc;
  }

  // Word ground truth (sample-18 Sec B): a vertAnchor=page/margin floating table
  // whose bottom would fall past its container is shifted UP to sit flush on the
  // container bottom (physical page edge for page-anchored: measured top 741.9pt =
  // 841.9 − 100 for a 100pt table), not left overflowing. vertAnchor="text" is
  // excluded — its overflow is handled by the paginator's row-split (the floating
  // analogue of splitTableAcrossPages), and its band rides the flow cursor, so
  // clamping here would be wrong. Mirrors computeFrameBox exactly (§17.3.1.11).
  if (tp.vertAnchor === 'page' || tp.vertAnchor === 'margin') {
    y = clampAbsBoxIntoContainer(y, tableH, vBand);
  }

  return { x, y, w: tableW, h: tableH };
}

/**
 * Push the wrap-exclusion FloatRect for a resolved floating-table box onto
 * `state.floats` so following body text flows around the table (§17.4.57). The
 * exclusion is the table box padded by the *FromText dist values. Overlap
 * avoidance (§17.4.56) runs FIRST (mirroring registerShapeFloat: resolve x/y,
 * THEN build the exclusion from the resolved x/y) with allowOverlap =
 * `table.overlap !== 'never'` (default true ⇒ overlap permitted).
 *
 * `side` (which side text wraps on) is computed by the caller from the resolved
 * box vs the column band: the table sits to one side and text fills the other
 * (§17.4.57). The x-range is built from the box (which for horzAnchor="text" is
 * column-relative via frameXContainer), so resolveLineFloatWindow only
 * constrains the matching column (#513), consistent with registerFrameFloat.
 *
 * Exported for unit tests only (the float-table-geometry table) — not package API.
 */
export function registerTableFloat(
  box: FloatTableBox,
  tp: TblpPr,
  state: RenderState,
  side: string,
  allowOverlap: boolean,
): void {
  if (box.w <= 0 || box.h <= 0) return;
  const sc = state.scale;
  const dl = tp.leftFromText * sc;
  const dr = tp.rightFromText * sc;
  const dt = tp.topFromText * sc;
  const db = tp.bottomFromText * sc;
  const paraId = state.floatParaSeq++;

  // §17.4.56: resolve overlap against already-registered page floats before
  // fixing the exclusion rect. allowOverlap=false (tblOverlap="never") forces
  // avoidance of OTHER FLOATING TABLES only — §17.4.56 scopes "never" to other
  // floating tables, NOT DrawingML anchors (§20.4.2.3) or text frames. The
  // kind==='table' tag below makes resolveFloatOverlap limit blockers to tables.
  // allowOverlap=true only avoids OTHER paragraphs' floats (the implementation-
  // defined scope documented on resolveFloatOverlap).
  pushFloatRect(state, {
    x: box.x,
    y: box.y,
    w: box.w,
    h: box.h,
    dl, dr, dt, db,
    kind: 'table',
    mode: 'square',
    side,
    imageKey: '', // non-image float: the table is painted by renderFloatTable.
    drawn: true, // painted by renderFloatTable; deferred image path must skip it.
    paraId,
    avoidOverlap: true,
    allowOverlap,
  });
}

/**
 * §17.4.57 — which side the body text wraps on, from the resolved float box vs
 * the current COLUMN band [contentX, contentX+contentW]. A floating table sits
 * to ONE side of the column and text fills the OTHER:
 *   - float's right edge at/left-of the column centre ⇒ float on the LEFT ⇒
 *     text wraps on the RIGHT (side='right').
 *   - float's left edge at/right-of the column centre ⇒ float on the RIGHT ⇒
 *     text wraps on the LEFT (side='left').
 *   - otherwise the float straddles the centre ⇒ 'bothSides' (resolveLineFloat-
 *     Window then takes the widest free gap on either flank).
 * Coordinates are page-absolute px; the comparison is against the column band so
 * a per-column floating table wraps within its own column (#513).
 *
 * NOTE: comparing against the column CENTRE is a simplified approximation of
 * Word's side selection — it picks a wrap side from where the float's box falls
 * relative to the column midpoint rather than the precise free-space-on-each-
 * flank measurement Word performs.
 *
 * Exported for unit tests only (the float-table-geometry table) — not package API.
 */
export function floatTableWrapSide(box: FloatTableBox, state: RenderState): string {
  const colLeft = state.contentX;
  const colRight = state.contentX + state.contentW;
  const center = (colLeft + colRight) / 2;
  if (box.x + box.w <= center) return 'right';
  if (box.x >= center) return 'left';
  return 'bothSides';
}
