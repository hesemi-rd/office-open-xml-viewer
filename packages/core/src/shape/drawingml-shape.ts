import type { Fill, PathCmd, Stroke } from '../types/common';
import { drawArrowHead, lineEndRetract, retractLineEndpoint } from './arrow';
import { buildCustomPath } from './custGeom';
import { getCustGeomEndpoints } from './custgeom-endpoints';
import { applyStroke, resolveFill } from './paint';
import { buildShapePath } from './preset';
import { getConnectorAnchors, hasPreset, renderPresetShape } from './preset-geometry';

type DeepReadonly<T> =
  T extends (...args: never[]) => unknown ? T
  : T extends readonly (infer U)[] ? readonly DeepReadonly<U>[]
  : T extends object ? { readonly [K in keyof T]: DeepReadonly<T[K]> }
  : T;

export type DrawingMLShapeFill = DeepReadonly<Exclude<Fill, { fillType: 'image' }>>;

export type DrawingMLShapeGeometry =
  | Readonly<{
      kind: 'preset';
      name: string;
      adjustments: readonly (number | null)[];
    }>
  | Readonly<{
      kind: 'custom';
      subpaths: readonly (readonly PathCmd[])[];
    }>;

export interface DrawingMLShapePaintPlan {
  readonly rect: Readonly<{ x: number; y: number; w: number; h: number }>;
  readonly geometry: DrawingMLShapeGeometry;
  readonly fill: DrawingMLShapeFill | null;
  readonly stroke: Readonly<Stroke> | null;
  readonly transform: Readonly<{
    rotationDeg: number;
    flipH: boolean;
    flipV: boolean;
  }>;
}

const CONNECTOR_GEOMETRIES = new Set([
  'line', 'straightconnector1',
  'bentconnector2', 'bentconnector3', 'bentconnector4', 'bentconnector5',
  'curvedconnector2', 'curvedconnector3', 'curvedconnector4', 'curvedconnector5',
]);

const CALLOUT_GEOMETRIES = new Set([
  'callout1', 'callout2', 'callout3',
  'bordercallout1', 'bordercallout2', 'bordercallout3',
  'accentcallout1', 'accentcallout2', 'accentcallout3',
  'accentbordercallout1', 'accentbordercallout2', 'accentbordercallout3',
]);

function retractableLeader(geometry: string): boolean {
  return CALLOUT_GEOMETRIES.has(geometry)
    || geometry === 'line'
    || geometry === 'straightconnector1'
    || geometry.startsWith('bentconnector');
}

function paintConnectorEnds(
  ctx: CanvasRenderingContext2D,
  plan: DrawingMLShapePaintPlan,
  geometry: string,
  unitToDevice: number,
): void {
  const stroke = plan.stroke as Stroke | null;
  if (!stroke || (!CONNECTOR_GEOMETRIES.has(geometry) && !CALLOUT_GEOMETRIES.has(geometry))) {
    return;
  }
  const { x, y, w, h } = plan.rect;
  const adjustments = plan.geometry.kind === 'preset' ? plan.geometry.adjustments : [];
  const anchors = getConnectorAnchors(geometry, x, y, w, h, [...adjustments]);
  if (!anchors) return;
  if (retractableLeader(geometry)
    && anchors.vertices.length >= 2
    && (stroke.headEnd || stroke.tailEnd)) {
    const points = anchors.vertices.map((vertex) => ({ x: vertex.x, y: vertex.y }));
    if (stroke.tailEnd) {
      points[points.length - 1] = retractLineEndpoint(
        points[points.length - 1],
        points[points.length - 2],
        lineEndRetract(stroke.tailEnd, stroke, unitToDevice),
      );
    }
    if (stroke.headEnd) {
      points[0] = retractLineEndpoint(
        points[0],
        points[1],
        lineEndRetract(stroke.headEnd, stroke, unitToDevice),
      );
    }
    applyStroke(ctx, stroke, unitToDevice);
    ctx.beginPath();
    ctx.moveTo(points[0].x, points[0].y);
    for (let index = 1; index < points.length; index++) {
      ctx.lineTo(points[index].x, points[index].y);
    }
    ctx.stroke();
  }
  if (stroke.tailEnd) {
    drawArrowHead(
      ctx, anchors.end.x, anchors.end.y, anchors.end.angle,
      stroke.tailEnd, stroke, unitToDevice,
    );
  }
  if (stroke.headEnd) {
    drawArrowHead(
      ctx, anchors.start.x, anchors.start.y, anchors.start.angle,
      stroke.headEnd, stroke, unitToDevice,
    );
  }
}

function paintCustomEnds(
  ctx: CanvasRenderingContext2D,
  plan: DrawingMLShapePaintPlan,
  unitToDevice: number,
): void {
  if (plan.geometry.kind !== 'custom') return;
  const stroke = plan.stroke as Stroke | null;
  if (!stroke || (!stroke.headEnd && !stroke.tailEnd)) return;
  const endpoints = getCustGeomEndpoints(plan.geometry.subpaths as PathCmd[][]);
  const { x, y, w, h } = plan.rect;
  if (endpoints.start && stroke.headEnd) {
    drawArrowHead(
      ctx,
      x + endpoints.start.x * w,
      y + endpoints.start.y * h,
      Math.atan2(endpoints.start.dy * h, endpoints.start.dx * w),
      stroke.headEnd,
      stroke,
      unitToDevice,
    );
  }
  if (endpoints.end && stroke.tailEnd) {
    drawArrowHead(
      ctx,
      x + endpoints.end.x * w,
      y + endpoints.end.y * h,
      Math.atan2(endpoints.end.dy * h, endpoints.end.dx * w),
      stroke.tailEnd,
      stroke,
      unitToDevice,
    );
  }
}

export function paintDrawingMLShape(
  ctx: CanvasRenderingContext2D,
  plan: DrawingMLShapePaintPlan,
  unitToDevice: number,
): void {
  const { x, y, w, h } = plan.rect;
  const { rotationDeg, flipH, flipV } = plan.transform;
  ctx.save();
  try {
    if (rotationDeg !== 0 || flipH || flipV) {
      ctx.translate(x + w / 2, y + h / 2);
      if (rotationDeg !== 0) ctx.rotate(rotationDeg * Math.PI / 180);
      ctx.scale(flipH ? -1 : 1, flipV ? -1 : 1);
      ctx.translate(-(x + w / 2), -(y + h / 2));
    }
    // Shared fill resolution is observational; retained plans keep gradient
    // stops readonly so layout snapshots cannot be mutated by a painter.
    const fillStyle = resolveFill(plan.fill as Fill | null, ctx, x, y, w, h);
    const stroke = plan.stroke as Stroke | null;
    const applyAndStroke = stroke
      ? () => {
          applyStroke(ctx, stroke, unitToDevice);
          ctx.stroke();
        }
      : null;
    if (plan.geometry.kind === 'preset') {
      const geometry = plan.geometry.name.toLowerCase();
      const adjustments = [...plan.geometry.adjustments];
      const hasDecoratedRetractableLeader = retractableLeader(geometry)
        && !!(stroke?.headEnd || stroke?.tailEnd);
      const painted = hasPreset(geometry) && renderPresetShape(
        ctx,
        geometry,
        x,
        y,
        w,
        h,
        adjustments,
        fillStyle,
        applyAndStroke,
        () => {},
        hasDecoratedRetractableLeader ? { skipTrailingStroke: true } : undefined,
      );
      if (!painted) {
        ctx.beginPath();
        buildShapePath(
          ctx, geometry, x, y, w, h,
          adjustments[0], adjustments[1], adjustments[2], adjustments[3],
        );
        if (fillStyle && geometry !== 'arc') {
          ctx.fillStyle = fillStyle;
          if (geometry === 'donut' || geometry === 'smileyface' || geometry === 'frame') {
            ctx.fill('evenodd');
          } else {
            ctx.fill();
          }
        }
        if (applyAndStroke) applyAndStroke();
      }
      paintConnectorEnds(ctx, plan, geometry, unitToDevice);
    } else {
      ctx.beginPath();
      buildCustomPath(ctx, plan.geometry.subpaths as PathCmd[][], x, y, w, h);
      if (fillStyle) {
        ctx.fillStyle = fillStyle;
        ctx.fill();
      }
      if (applyAndStroke) applyAndStroke();
      paintCustomEnds(ctx, plan, unitToDevice);
    }
  } finally {
    ctx.restore();
  }
}
