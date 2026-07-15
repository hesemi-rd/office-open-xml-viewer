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
import type {
  FloatingTablePlacementLayout,
  FloatingTablePositionInput,
  FloatingTableReferenceFramesPt,
  LayoutRect,
  ResolvedFloatingTablePlacementLayout,
} from './layout/types.js';
import {
  frameXContainer,
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

function frameEndX(frame: LayoutRect): number {
  return frame.xPt + frame.widthPt;
}

function frameEndY(frame: LayoutRect): number {
  return frame.yPt + frame.heightPt;
}

function resolvePointSpaceFloatBox(
  positioning: FloatingTablePositionInput,
  frames: FloatingTableReferenceFramesPt,
  widthPt: number,
  heightPt: number,
  skipVClamp = false,
): FloatTableBox {
  const horizontalFrame = positioning.horzSpecified
    ? positioning.horzAnchor === 'page'
      ? frames.page
      : positioning.horzAnchor === 'margin'
        ? frames.margin
        : frames.text
    : frames.text;
  const verticalFrame = positioning.vertAnchor === 'page'
    ? frames.page
    : positioning.vertAnchor === 'margin'
      ? frames.margin
      : frames.text;
  const x = positioning.xAlign
    ? resolveAlignedPosH(
        positioning.xAlign,
        horizontalFrame.xPt,
        frameEndX(horizontalFrame),
        widthPt,
      )
    : horizontalFrame.xPt + positioning.xPt;
  let y = positioning.yAlign && positioning.vertAnchor !== 'text'
    ? resolveAlignedPosV(
        positioning.yAlign,
        { start: verticalFrame.yPt, end: frameEndY(verticalFrame) },
        heightPt,
      )
    : verticalFrame.yPt + positioning.yPt;
  if (!skipVClamp && (positioning.vertAnchor === 'page' || positioning.vertAnchor === 'margin')) {
    y = clampAbsBoxIntoContainer(
      y,
      heightPt,
      { start: verticalFrame.yPt, end: frameEndY(verticalFrame) },
    );
  }
  return { x, y, w: widthPt, h: heightPt };
}

/** Resolve parser-independent §17.4.57 positioning facts in point space. */
export function resolveFloatingTableBoxPt(
  positioning: FloatingTablePositionInput,
  frames: FloatingTableReferenceFramesPt,
  widthPt: number,
  heightPt: number,
): FloatTableBox {
  return resolvePointSpaceFloatBox(positioning, frames, widthPt, heightPt);
}

/** Register one point-space floating-table exclusion through the shared float registry. */
export function registerFloatingTableBoxPt(
  box: FloatTableBox,
  positioning: FloatingTablePositionInput,
  overlap: 'never' | 'overlap',
  state: RenderState,
  side = 'bothSides',
): ReturnType<typeof pushFloatRect> {
  const sc = state.scale;
  return pushFloatRect(state, {
    x: box.x * sc,
    y: box.y * sc,
    w: box.w * sc,
    h: box.h * sc,
    dl: positioning.leftFromTextPt * sc,
    dr: positioning.rightFromTextPt * sc,
    dt: positioning.topFromTextPt * sc,
    db: positioning.bottomFromTextPt * sc,
    kind: 'table',
    mode: 'square',
    side,
    imageKey: '',
    drawn: true,
    paraId: state.floatParaSeq++,
    avoidOverlap: true,
    allowOverlap: overlap !== 'never',
  });
}

/** Resolve retained §17.4.57 facts against explicit page-local reference frames. */
export function resolveFloatingTablePlacement(
  placement: FloatingTablePlacementLayout,
  frames: FloatingTableReferenceFramesPt,
): ResolvedFloatingTablePlacementLayout {
  const widthPt = placement.child.columnWidthsPt.reduce((sum, width) => sum + width, 0);
  const heightPt = placement.child.advancePt;
  const rawBox = resolveFloatingTableBoxPt(
    placement.positioning,
    frames,
    widthPt,
    heightPt,
  );
  const usesTextX = !placement.positioning.horzSpecified
    || (placement.positioning.horzAnchor !== 'page'
      && placement.positioning.horzAnchor !== 'margin');
  const usesTextY = placement.positioning.vertAnchor !== 'page'
    && placement.positioning.vertAnchor !== 'margin';
  const box = {
    ...rawBox,
    x: usesTextX && placement.acquiredTextOffsetPt
      ? frames.text.xPt + placement.acquiredTextOffsetPt.xPt
      : rawBox.x,
    y: usesTextY && placement.acquiredTextOffsetPt
      ? frames.text.yPt + placement.acquiredTextOffsetPt.yPt
      : rawBox.y,
  };
  const { positioning } = placement;
  const bounds = Object.freeze({
    xPt: box.x,
    yPt: box.y,
    widthPt: box.w,
    heightPt: box.h,
  });
  const exclusionBounds = Object.freeze({
    xPt: box.x - positioning.leftFromTextPt,
    yPt: box.y - positioning.topFromTextPt,
    widthPt: box.w + positioning.leftFromTextPt + positioning.rightFromTextPt,
    heightPt: box.h + positioning.topFromTextPt + positioning.bottomFromTextPt,
  });
  return Object.freeze({
    kind: 'resolved-floating-table-placement',
    occurrenceId: placement.occurrenceId,
    xPt: box.x,
    yPt: box.y,
    bounds,
    exclusionBounds,
    overlap: placement.overlap,
    child: placement.child,
    source: placement,
  });
}

/** Register the retained exclusion and return the collision-adjusted paint box. */
export function registerResolvedFloatingTablePlacement(
  placement: ResolvedFloatingTablePlacementLayout,
  state: RenderState,
): ResolvedFloatingTablePlacementLayout {
  const sc = state.scale;
  const positioning = placement.source.positioning;
  const rect = registerFloatingTableBoxPt({
    x: placement.xPt,
    y: placement.yPt,
    w: placement.bounds.widthPt,
    h: placement.bounds.heightPt,
  }, positioning, placement.overlap, state);
  const xPt = rect.imageX / sc;
  const yPt = rect.imageY / sc;
  return Object.freeze({
    ...placement,
    xPt,
    yPt,
    bounds: Object.freeze({
      xPt,
      yPt,
      widthPt: placement.bounds.widthPt,
      heightPt: placement.bounds.heightPt,
    }),
    exclusionBounds: Object.freeze({
      xPt: rect.xLeft / sc,
      yPt: rect.yTop / sc,
      widthPt: (rect.xRight - rect.xLeft) / sc,
      heightPt: (rect.yBottom - rect.yTop) / sc,
    }),
  });
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
  /** When true, skip the observed Office vertAnchor=page/margin bottom-clamp so
   *  the raw absolute box is returned. The paginator
   *  uses this to find where a page/margin-anchored table that overflows the text
   *  region must be row-split: the raw tblpY top drives slice 1's
   *  position, and clamping (which pins a too-tall box to the container top) would
   *  hide the overflow the split is meant to resolve. Placement (paint) keeps the
   *  clamp. Ignored for vertAnchor="text" (never clamped). */
  skipVClamp = false,
): FloatTableBox {
  const sc = state.scale;
  const textBand = frameXContainer('text', state);
  return resolvePointSpaceFloatBox({
    leftFromTextPt: tp.leftFromText * sc,
    rightFromTextPt: tp.rightFromText * sc,
    topFromTextPt: tp.topFromText * sc,
    bottomFromTextPt: tp.bottomFromText * sc,
    horzAnchor: tp.horzAnchor,
    horzSpecified: tp.horzSpecified,
    vertAnchor: tp.vertAnchor,
    xPt: tp.tblpX * sc,
    yPt: tp.tblpY * sc,
    ...(tp.tblpXSpec == null ? {} : { xAlign: tp.tblpXSpec }),
    ...(tp.tblpYSpec == null ? {} : { yAlign: tp.tblpYSpec }),
  }, {
    page: {
      xPt: 0,
      yPt: 0,
      widthPt: state.pageWidth * sc,
      heightPt: state.pageH,
    },
    margin: {
      xPt: state.marginLeft * sc,
      yPt: state.marginTop * sc,
      widthPt: Math.max(0, (state.pageWidth - state.marginLeft - state.marginRight) * sc),
      heightPt: Math.max(0, state.pageH - (state.marginTop + state.marginBottom) * sc),
    },
    text: {
      xPt: textBand.left,
      yPt: paraTop,
      widthPt: Math.max(0, textBand.right - textBand.left),
      heightPt: tableH,
    },
  }, tableW, tableH, skipVClamp);
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
 * `tblpPr` supplies an exclusion rectangle but no left/right wrap-side choice.
 * Preserve both candidate flanks so `resolveLineFloatWindow` can apply the same
 * widest-free-gap geometry used for other square exclusions. Choosing a side
 * from the column midpoint discarded usable space and had no OOXML basis.
 *
 * Exported for unit tests only — not package API.
 */
export function floatTableWrapSide(_box: FloatTableBox, _state: RenderState): string {
  return 'bothSides';
}
