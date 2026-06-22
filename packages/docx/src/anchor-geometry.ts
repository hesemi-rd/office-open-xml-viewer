// DrawingML anchor placement geometry (ECMA-376 §20.4.3.x).
//
// Pure placement math for `<wp:positionH>` / `<wp:positionV>` (relativeFrom +
// posOffset / align / pctPos): given the container indicated by `relativeFrom`
// and the section margins on `RenderState`, it answers "where does this anchor /
// anchor-group child sit?". Extracted from renderer.ts so the resolve logic can
// be unit-reasoned in isolation (see anchor-align.test.ts). Only `RenderState`
// is imported as a type (erased at runtime), so there is no import cycle with
// renderer.ts.

import type { RenderState } from './renderer.js';

/** Resolve a shape's page X by combining the explicit `anchorXPt` offset with
 *  any `anchorXAlign` (ECMA-376 §20.4.3.1 wp:align). When align is set we
 *  position the shape inside the container indicated by `relativeFrom` (or
 *  `anchorXFromMargin` for the legacy two-state hint). When `pctPos` is set
 *  we ignore the explicit offset and place the shape at `pct` of the
 *  container's width / height (ECMA-376 §20.4.2.7 wp14:pctPosH/VOffset).
 *
 *  relativeFrom containers (ECMA-376 §20.4.3.4):
 *    - "page"          → full page rect
 *    - "margin"        → printable area between margins
 *    - "leftMargin"    → strip from x=0 to x=marginLeft
 *    - "rightMargin"   → strip from x=pageW-marginRight to x=pageW
 *    - "insideMargin"  → on odd pages = leftMargin, even = rightMargin
 *                        (we approximate as leftMargin)
 *    - "outsideMargin" → on odd pages = rightMargin, even = leftMargin
 *                        (we approximate as rightMargin)
 *    - "character"     → degrade to "margin" (no run-relative anchor data)
 *    - "topMargin"     → strip from y=0 to y=marginTop
 *    - "bottomMargin"  → strip from y=pageH-marginBottom to y=pageH
 *    - "paragraph"/"line" → relative to paragraph top (V only) */
export function xContainer(
  relativeFrom: string | null | undefined,
  fromMarginHint: boolean,
  state: RenderState,
): { start: number; end: number } {
  const { scale } = state;
  const pageW = state.pageWidth * scale;
  const ml = state.marginLeft * scale;
  const mr = state.marginRight * scale;
  const rf = relativeFrom ?? (fromMarginHint ? 'margin' : 'page');
  switch (rf) {
    case 'page':          return { start: 0, end: pageW };
    case 'leftMargin':    return { start: 0, end: ml };
    case 'rightMargin':   return { start: pageW - mr, end: pageW };
    case 'insideMargin':  return { start: 0, end: ml };
    case 'outsideMargin': return { start: pageW - mr, end: pageW };
    case 'margin':
    case 'character':
    case 'column':
    default:              return { start: ml, end: pageW - mr };
  }
}

export function yContainer(
  relativeFrom: string | null | undefined,
  fromParaHint: boolean,
  paragraphTopPx: number,
  state: RenderState,
): { start: number; end: number } {
  const { scale } = state;
  const mt = state.marginTop * scale;
  const mb = state.marginBottom * scale;
  const rf = relativeFrom ?? (fromParaHint ? 'paragraph' : 'page');
  switch (rf) {
    case 'page':         return { start: 0, end: state.pageH };
    case 'topMargin':    return { start: 0, end: mt };
    case 'bottomMargin': return { start: state.pageH - mb, end: state.pageH };
    case 'paragraph':
    case 'line':         return { start: paragraphTopPx, end: state.pageH };
    case 'margin':
    default:             return { start: mt, end: state.pageH - mb };
  }
}

/** Resolve the page X for an anchor or anchor-group child. `offsetPx` carries
 *  the shape's offset (within the group for wgp children, 0 for standalone
 *  anchors). `alignWidthPx` is the width used when aligning — the GROUP's
 *  width for wgp children, the shape's own width for standalone anchors. */
export function resolveAnchorX(
  align: string | null | undefined,
  fromMargin: boolean,
  offsetPt: number,
  widthPx: number,
  state: RenderState,
  relativeFrom?: string | null,
  pctPos?: number | null,
  alignWidthPt?: number | null,
): number {
  const { scale } = state;
  const c = xContainer(relativeFrom, fromMargin, state);
  const offsetPx = offsetPt * scale;
  if (pctPos != null) {
    return c.start + (c.end - c.start) * pctPos + offsetPx;
  }
  if (!align) {
    return c.start + offsetPx;
  }
  const containerW = c.end - c.start;
  const aw = alignWidthPt != null ? alignWidthPt * scale : widthPx;
  switch (align) {
    case 'center': return c.start + (containerW - aw) / 2 + offsetPx;
    case 'right':
    case 'outside': return c.end - aw + offsetPx;
    case 'inside':
    case 'left':
    default:        return c.start + offsetPx;
  }
}

export function resolveAnchorY(
  align: string | null | undefined,
  fromPara: boolean,
  offsetPt: number,
  heightPx: number,
  paragraphTopPx: number,
  state: RenderState,
  relativeFrom?: string | null,
  pctPos?: number | null,
  alignHeightPt?: number | null,
): number {
  const { scale } = state;
  const c = yContainer(relativeFrom, fromPara, paragraphTopPx, state);
  const offsetPx = offsetPt * scale;
  if (pctPos != null) {
    return c.start + (c.end - c.start) * pctPos + offsetPx;
  }
  if (!align) {
    return c.start + offsetPx;
  }
  const containerH = c.end - c.start;
  const ah = alignHeightPt != null ? alignHeightPt * scale : heightPx;
  switch (align) {
    case 'center': return c.start + (containerH - ah) / 2 + offsetPx;
    // ECMA-376 §20.4.3.1 ST_AlignV: "inside"/"outside" are page-binding-
    // relative. Mirroring resolveAnchorX (and the insideMargin/outsideMargin
    // approximation in yContainer/xContainer): on an odd page the binding edge
    // is the top, so inside→top edge and outside→bottom edge. This is an
    // odd-page approximation; the true §20.4.3.1 page-parity behavior (even
    // pages mirror the binding edge) is not implemented. Update this when that
    // approximation is removed.
    case 'bottom':
    case 'outside': return c.end - ah + offsetPx;
    case 'top':
    case 'inside':
    default:        return c.start + offsetPx;
  }
}
