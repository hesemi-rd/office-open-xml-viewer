import {
  crispOffset,
  doubleRailGeometry,
} from '@silurus/ooxml-core';
import type { BorderSegment, TextDecorationLayout } from '../layout/types.js';
import {
  inverseMapAffinePoint,
  inverseMapAffineVector,
  mapAffinePoint,
  scaleAffine,
} from './affine.js';
import type { CanvasPaintContext } from './types.js';

/** Paint one already-resolved point-space rule; layout owns every conflict and path. */
export function paintStrokeSegment(
  segment: BorderSegment | TextDecorationLayout,
  context: CanvasPaintContext,
): void {
  const { ctx } = context;
  ctx.strokeStyle = segment.color;
  ctx.lineWidth = segment.widthPt;
  ctx.setLineDash('dashPatternPt' in segment && segment.dashPatternPt
    ? [...segment.dashPatternPt]
    : []);
  ctx.beginPath();
  const path = 'path' in segment && segment.path?.length ? segment.path : [segment.from, segment.to];
  const axisAligned = path.length === 2
    && (path[0]!.xPt === path[1]!.xPt || path[0]!.yPt === path[1]!.yPt);
  const horizontal = axisAligned && path[0]!.yPt === path[1]!.yPt;
  const vertical = axisAligned && path[0]!.xPt === path[1]!.xPt;
  const pointToCss = context.pointToCss ?? scaleAffine(context.scale);
  const finalPath = path.map((point) => mapAffinePoint(pointToCss, point));
  const localDx = axisAligned ? path[1]!.xPt - path[0]!.xPt : 0;
  const localDy = axisAligned ? path[1]!.yPt - path[0]!.yPt : 0;
  const finalDx = pointToCss.a * localDx + pointToCss.c * localDy;
  const finalDy = pointToCss.b * localDx + pointToCss.d * localDy;
  const finalHorizontal = axisAligned && finalDy === 0;
  const finalVertical = axisAligned && finalDx === 0;
  const normalScale = horizontal
    ? Math.hypot(pointToCss.c, pointToCss.d)
    : vertical ? Math.hypot(pointToCss.a, pointToCss.b) : 0;
  if (segment.style === 'double' && axisAligned && normalScale > 0) {
    ctx.fillStyle = segment.color;
    if (finalHorizontal || finalVertical) {
      const fillFinalRect = (x: number, y: number, width: number, height: number): void => {
        const corners = [
          { xPt: x, yPt: y },
          { xPt: x + width, yPt: y },
          { xPt: x, yPt: y + height },
          { xPt: x + width, yPt: y + height },
        ].map((point) => inverseMapAffinePoint(pointToCss, point));
        if (corners.some((point) => point === null)) return;
        const local = corners.filter((point): point is { xPt: number; yPt: number } => point !== null);
        const xs = local.map((point) => point.xPt);
        const ys = local.map((point) => point.yPt);
        ctx.fillRect(
          Math.min(...xs), Math.min(...ys),
          Math.max(...xs) - Math.min(...xs),
          Math.max(...ys) - Math.min(...ys),
        );
      };
      const { railDev, gapDev, spanDev } = doubleRailGeometry(
        segment.widthPt * normalScale,
        context.dpr,
      );
      const railCss = railDev / context.dpr;
      // Snapping must happen after the full affine transform. Mapping a snapped
      // device-space rectangle back to local coordinates preserves that result
      // without presenting a partial object as a Canvas context.
      if (finalHorizontal) {
        const startDev = Math.round(finalPath[0]!.yPt * context.dpr - spanDev / 2);
        const x = Math.min(finalPath[0]!.xPt, finalPath[1]!.xPt);
        const width = Math.abs(finalPath[1]!.xPt - finalPath[0]!.xPt);
        fillFinalRect(x, startDev / context.dpr, width, railCss);
        fillFinalRect(x, (startDev + railDev + gapDev) / context.dpr, width, railCss);
      } else {
        const startDev = Math.round(finalPath[0]!.xPt * context.dpr - spanDev / 2);
        const y = Math.min(finalPath[0]!.yPt, finalPath[1]!.yPt);
        const height = Math.abs(finalPath[1]!.yPt - finalPath[0]!.yPt);
        fillFinalRect(startDev / context.dpr, y, railCss, height);
        fillFinalRect((startDev + railDev + gapDev) / context.dpr, y, railCss, height);
      }
    } else {
      // A general affine has no device row/column to snap against, but the two
      // authored rails still retain their point-space separation.
      const { railDev, gapDev, spanDev } = doubleRailGeometry(
        segment.widthPt * normalScale,
        context.dpr,
      );
      const railPt = railDev / context.dpr / normalScale;
      const gapPt = gapDev / context.dpr / normalScale;
      const spanPt = spanDev / context.dpr / normalScale;
      if (horizontal) {
        const x = Math.min(path[0]!.xPt, path[1]!.xPt);
        const width = Math.abs(path[1]!.xPt - path[0]!.xPt);
        ctx.fillRect(x, path[0]!.yPt - spanPt / 2, width, railPt);
        ctx.fillRect(x, path[0]!.yPt - spanPt / 2 + railPt + gapPt, width, railPt);
      } else {
        const y = Math.min(path[0]!.yPt, path[1]!.yPt);
        const height = Math.abs(path[1]!.yPt - path[0]!.yPt);
        ctx.fillRect(path[0]!.xPt - spanPt / 2, y, railPt, height);
        ctx.fillRect(path[0]!.xPt - spanPt / 2 + railPt + gapPt, y, railPt, height);
      }
    }
    ctx.setLineDash([]);
    return;
  }
  const cssOffset = finalVertical && normalScale > 0
    ? { xPt: crispOffset(finalPath[0]!.xPt, segment.widthPt * normalScale, context.dpr), yPt: 0 }
    : finalHorizontal && normalScale > 0
      ? { xPt: 0, yPt: crispOffset(finalPath[0]!.yPt, segment.widthPt * normalScale, context.dpr) }
      : { xPt: 0, yPt: 0 };
  const localOffset = inverseMapAffineVector(pointToCss, cssOffset) ?? { xPt: 0, yPt: 0 };
  const first = path[0]!;
  ctx.moveTo(first.xPt + localOffset.xPt, first.yPt + localOffset.yPt);
  for (const point of path.slice(1)) {
    ctx.lineTo(point.xPt + localOffset.xPt, point.yPt + localOffset.yPt);
  }
  ctx.stroke();
  ctx.setLineDash([]);
}
