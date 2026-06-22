// Text-frame / drop-cap placement geometry (ECMA-376 §17.3.1.11 `<w:framePr>`).
//
// Pure placement math: given a `<w:framePr>` and the section geometry on
// `RenderState`, resolve the frame box (canvas px) and the wrap-exclusion
// FloatRect it pushes onto `state.floats`. Extracted from renderer.ts so the
// resolve logic can be unit-reasoned in isolation (see frame-geometry.test.ts /
// measure-column-geometry.test.ts).
//
// This module is the shared base for floating-table placement
// (float-table-geometry.ts), which reuses frameXContainer / frameYContainer /
// resolveAlignedPosH / resolveAlignedPosV (the anchor/alignment semantics line
// up 1:1 between a text frame and a floating table) and pushFloatRect (the
// single source of exclusion-rect construction). Only `RenderState` is imported
// as a type (erased at runtime), so there is no import cycle with renderer.ts.

import type { FramePr } from './types.js';
import type { RenderState } from './renderer.js';
import { type FloatRect, resolveFloatOverlap } from './float-layout.js';

/** Resolved geometry (canvas px) of a `<w:framePr>` text frame. Exported for
 *  unit tests only (the table-driven frame-geometry assertions) — not part of
 *  the package API. */
export interface FrameBox {
  /** Drawing origin of the frame content (text area top-left). */
  x: number;
  y: number;
  /** Frame content width / height. */
  w: number;
  h: number;
  /** Padded exclusion rect for the wrap FloatRect (frame + hSpace/vSpace). */
  exLeft: number;
  exRight: number;
  exTop: number;
  exBottom: number;
}

/**
 * Horizontal container band for a frame's hAnchor (ECMA-376 §17.3.1.11 /
 * §17.18.35). This is a SEPARATE relativeFrom set from DrawingML's
 * §20.4.3 (so `xContainer` in anchor-geometry is intentionally not reused):
 *   - "text"   → the COLUMN text margin the anchor paragraph sits in
 *                (state.contentX..contentX+contentW). This keeps a drop cap
 *                inside its own newspaper column (#513 per-section columns).
 *   - "margin" → the page content margin (marginLeft..pageWidth-marginRight).
 *   - "page"   → the physical page edges (0..pageWidth).
 * All values in canvas px.
 */
export function frameXContainer(hAnchor: string, state: RenderState): { left: number; right: number } {
  const sc = state.scale;
  switch (hAnchor) {
    case 'margin':
      return { left: state.marginLeft * sc, right: (state.pageWidth - state.marginRight) * sc };
    case 'page':
      return { left: 0, right: state.pageWidth * sc };
    case 'text':
    case 'column':
    default:
      // "text" anchors against the current COLUMN band so a frame in a multi-
      // column section stays inside its column.
      return { left: state.contentX, right: state.contentX + state.contentW };
  }
}

/**
 * Vertical container origin for a frame's vAnchor (ECMA-376 §17.3.1.11 /
 * §17.18.100). `paraTop` is the anchor paragraph's text-area top (canvas px).
 *   - "text"   → the paragraph top (y offsets/relative positions are measured
 *                from where the frame paragraph sits in the flow).
 *   - "margin" → the page top content margin.
 *   - "page"   → the physical page top.
 */
export function frameYContainer(vAnchor: string, paraTop: number, state: RenderState): number {
  const sc = state.scale;
  switch (vAnchor) {
    case 'margin':
      return state.marginTop * sc;
    case 'page':
      return 0;
    case 'text':
    default:
      return paraTop;
  }
}

/**
 * Resolve a horizontal aligned position (canvas px) for a frame (xAlign,
 * §17.3.1.11) or a floating table (tblpXSpec, §17.4.57). Both use the same
 * ST_XAlign vocabulary against a container band [containerLeft, containerRight]:
 *   center          → box centred in the band
 *   right / outside  → box flush to the band's right edge
 *   left / inside / * → box flush to the band's left edge (the default)
 * Shared by {@link computeFrameBox} and computeFloatTableBox so the two
 * stay byte-identical.
 */
export function resolveAlignedPosH(
  spec: string,
  containerLeft: number,
  containerRight: number,
  size: number,
): number {
  switch (spec) {
    case 'center':
      return containerLeft + (containerRight - containerLeft - size) / 2;
    case 'right':
    case 'outside':
      return containerRight - size;
    case 'left':
    case 'inside':
    default:
      return containerLeft;
  }
}

/**
 * Resolve a vertical aligned position (canvas px) for a frame (yAlign,
 * §17.3.1.11) or a floating table (tblpYSpec, §17.4.57). Both use the same
 * ST_YAlign vocabulary, measured against the page box (not the band) per spec:
 *   center           → box centred between the top/bottom content margins
 *   bottom / outside  → box flush to the bottom content margin
 *   top / inside / inline / * → box at the vAnchor origin `vy` (the default)
 * Callers gate this on vAnchor!=='text' (relative vertical positioning is not
 * allowed there). Shared by {@link computeFrameBox} and computeFloatTableBox.
 */
export function resolveAlignedPosV(
  spec: string,
  vy: number,
  size: number,
  state: RenderState,
): number {
  const sc = state.scale;
  switch (spec) {
    case 'center':
      return vy + (state.pageH - size) / 2 - state.marginTop * sc;
    case 'bottom':
    case 'outside':
      return state.pageH - state.marginBottom * sc - size;
    case 'top':
    case 'inside':
    case 'inline':
    default:
      return vy;
  }
}

/**
 * Resolve a frame's box in canvas px. `paraTop` is the in-flow top of the frame
 * paragraph (post-spaceBefore). `contentW`/`contentH` are the frame content's
 * measured natural size (px); `anchorLineHpx` is one line height of the
 * following non-frame (anchor) paragraph, used to size a drop cap by `lines`.
 *
 * Exported for unit tests only (frame-geometry table) — not package API.
 */
export function computeFrameBox(
  fp: FramePr,
  state: RenderState,
  paraTop: number,
  contentW: number,
  contentH: number,
  anchorLineHpx: number,
): FrameBox {
  const sc = state.scale;
  const isDropCap = fp.dropCap === 'drop' || fp.dropCap === 'margin';

  const hx = frameXContainer(fp.hAnchor, state);
  const vy = frameYContainer(fp.vAnchor, paraTop, state);

  // Frame width: explicit `w` (exact) else natural content width (§17.3.1.11 w).
  const frameW = fp.w != null ? fp.w * sc : contentW;

  // Frame height. For a drop cap the height is `lines` × the anchor paragraph's
  // line height (§17.3.1.11 lines: "the height of the drop cap is the first N
  // lines of the anchor paragraph"; y/yAlign are ignored). For a generic frame
  // hRule gates h: exact = h, atLeast = max(h, content), auto = content.
  let frameH: number;
  if (isDropCap) {
    frameH = Math.max(1, fp.lines) * anchorLineHpx;
  } else {
    const hPx = fp.h != null ? fp.h * sc : 0;
    frameH =
      fp.hRule === 'exact'
        ? hPx
        : fp.hRule === 'atLeast'
          ? Math.max(hPx, contentH)
          : contentH;
  }

  // Horizontal placement.
  //   dropCap="drop"   → inside the column/text margin (frame at band left).
  //   dropCap="margin" → outside the margin (frame left = band left − frameW).
  //   generic frame    → xAlign (left/center/right/inside/outside) supersedes x;
  //                      else absolute x offset from the hAnchor's left edge.
  let frameX: number;
  if (fp.dropCap === 'drop') {
    frameX = hx.left;
  } else if (fp.dropCap === 'margin') {
    frameX = hx.left - frameW;
  } else if (fp.xAlign) {
    frameX = resolveAlignedPosH(fp.xAlign, hx.left, hx.right, frameW);
  } else {
    // §17.3.1.11 x: absolute signed offset from the hAnchor left edge.
    frameX = hx.left + (fp.x != null ? fp.x * sc : 0);
  }

  // Vertical placement. For a drop cap, y/yAlign are ignored: the cap sits at
  // the anchor paragraph top (§17.3.1.11 lines). Otherwise yAlign supersedes y
  // (ignored when vAnchor="text" — relative positioning is not allowed there,
  // §17.3.1.11 yAlign), else absolute y offset from the vAnchor edge.
  let frameY: number;
  if (isDropCap) {
    frameY = vy;
  } else if (fp.yAlign && fp.vAnchor !== 'text') {
    frameY = resolveAlignedPosV(fp.yAlign, vy, frameH, state);
  } else {
    frameY = vy + (fp.y != null ? fp.y * sc : 0);
  }

  // Exclusion padding: hSpace L/R applies only with wrap="around" (§17.3.1.11
  // hSpace); vSpace top/bottom always.
  const hSpacePx = fp.wrap === 'around' || fp.wrap === 'auto' ? fp.hSpace * sc : 0;
  const vSpacePx = fp.vSpace * sc;

  return {
    x: frameX,
    y: frameY,
    w: frameW,
    h: frameH,
    exLeft: frameX - hSpacePx,
    exRight: frameX + frameW + hSpacePx,
    exTop: frameY - vSpacePx,
    exBottom: frameY + frameH + vSpacePx,
  };
}

/** Options for {@link pushFloatRect}: the resolved image/float box (x,y,w,h),
 *  its dist* padding (dl,dr,dt,db, all px), and the FloatRect descriptors. */
export interface PushFloatOpts {
  x: number;
  y: number;
  w: number;
  h: number;
  dl: number;
  dr: number;
  dt: number;
  db: number;
  /** What reserved this float — scopes overlap avoidance (§17.4.56). See
   *  {@link FloatRect.kind}. */
  kind: FloatRect['kind'];
  mode: 'square' | 'topAndBottom';
  side: string;
  imageKey: string;
  drawn: boolean;
  paraId: number;
  /** Run §20.4.2.3 / §17.4.56 overlap avoidance before fixing the rect. When
   *  false (frame floats) the box is used as-is. */
  avoidOverlap: boolean;
  /** allowOverlap arg passed to resolveFloatOverlap when avoidOverlap is true
   *  (true ⇒ only avoid OTHER paragraphs' floats; false ⇒ spec-mandated
   *  avoidance, scoped by `kind`: a table avoids only other tables, §17.4.56).
   *  Ignored when avoidOverlap is false. */
  allowOverlap?: boolean;
}

/**
 * Build a wrap-exclusion {@link FloatRect} from a resolved box + dist padding,
 * optionally running overlap avoidance first, push it onto `state.floats`, and
 * return it. Single source of the `xLeft = x − dl, xRight = x + w + dr,
 * yTop = y − dt, yBottom = y + h + db` exclusion-rect construction shared by
 * registerFrameFloat / registerTableFloat / registerImageFloat /
 * registerShapeFloat (the `dist*` fields carry dl/dr/dt/db verbatim for
 * re-seating). The returned ref lets the image path flip `drawn` after painting.
 */
export function pushFloatRect(state: RenderState, o: PushFloatOpts): FloatRect {
  let px = o.x;
  let py = o.y;
  if (o.avoidOverlap) {
    const resolved = resolveFloatOverlap(
      px, py, o.w, o.h, o.dl, o.dr, o.dt, o.db, o.paraId, o.allowOverlap ?? true,
      o.kind, state.pageWidth * state.scale, state.floats,
    );
    px = resolved.x;
    py = resolved.y;
  }
  const rect: FloatRect = {
    kind: o.kind,
    mode: o.mode,
    imageKey: o.imageKey,
    imageX: px,
    imageY: py,
    imageW: o.w,
    imageH: o.h,
    xLeft: px - o.dl,
    xRight: px + o.w + o.dr,
    yTop: py - o.dt,
    yBottom: py + o.h + o.db,
    side: o.side,
    distLeft: o.dl,
    distRight: o.dr,
    distTop: o.dt,
    distBottom: o.db,
    paraId: o.paraId,
    drawn: o.drawn,
  };
  state.floats.push(rect);
  return rect;
}

/**
 * Push the wrap-exclusion FloatRect for a resolved frame box onto
 * `state.floats` so following body text flows around the frame. No-op for
 * wrap="none" or a degenerate (zero-area) box. Shared by the renderer (after
 * drawing) and the paginator (so the anchor paragraph's measured height
 * accounts for the wrap). The exclusion x-range is COLUMN-relative (built in
 * frameXContainer from state.contentX/contentW for hAnchor="text"), so
 * resolveLineFloatWindow only constrains the matching column (#513).
 *
 * Wrap-mode → FloatRect mapping (ECMA-376 §17.18.104):
 *   none      → no exclusion (text may overlap; the frame is drawn absolutely
 *               and following text starts at its normal Y).
 *   notBeside → topAndBottom (text never sits beside the frame).
 *   around / auto → square side wrap. "auto" is Word's application-defined
 *               default, effectively "around" in Word, so treated as around.
 *   tight / through → a frame is a rectangle, so contour wrapping collapses to
 *               a square wrap (no contour follow for a rectangular frame).
 *
 * Exported for unit tests only (frame-geometry table) — not package API.
 */
export function registerFrameFloat(box: FrameBox, fp: FramePr, state: RenderState): void {
  if (fp.wrap === 'none') return;
  if (box.w <= 0 || box.h <= 0) return;

  const paraId = state.floatParaSeq++;
  const mode: 'square' | 'topAndBottom' = fp.wrap === 'notBeside' ? 'topAndBottom' : 'square';
  // dist padding recovered from the box's pre-computed exclusion edges so the
  // unified builder reproduces xLeft=box.exLeft etc. exactly.
  pushFloatRect(state, {
    x: box.x,
    y: box.y,
    w: box.w,
    h: box.h,
    dl: box.x - box.exLeft,
    dr: box.exRight - (box.x + box.w),
    dt: box.y - box.exTop,
    db: box.exBottom - (box.y + box.h),
    kind: 'frame',
    mode,
    // A drop cap sits at the column's left edge, so text wraps only to its
    // RIGHT. A generic frame may sit anywhere, so text wraps on both sides
    // (resolveLineFloatWindow then takes the widest free gap around it).
    side: fp.dropCap === 'drop' || fp.dropCap === 'margin' ? 'right' : 'bothSides',
    imageKey: '', // non-image float: the frame is painted above, not deferred.
    drawn: true, // painted by renderFrameParagraph; deferred path must skip it.
    paraId,
    avoidOverlap: false, // frames opt out of overlap re-seating.
  });
}
