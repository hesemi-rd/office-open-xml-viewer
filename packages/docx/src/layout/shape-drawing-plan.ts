import type { ArrowEnd, DrawingMLShapePaintPlan, Stroke } from '@silurus/ooxml-core';
import type { ShapeRun } from '../types.js';
import type {
  DrawingPaintCommand,
  LayoutRect,
  VmlTextPathAcquisitionInput,
} from './types.js';
import { snapshotPlainData } from './plain-data.js';
import type { TextLayoutService } from './text.js';

export type ShapeDrawingPlanResult = Readonly<{
  status: 'planned';
  command: Extract<DrawingPaintCommand, { kind: 'noop' | 'drawingml-shape' | 'watermark-text' }>;
}>;

// When fitshape is active, every shaped dimension is multiplied by the same
// target/source ratio, so a 1pt internal reference cancels out exactly. This is
// used only when VML omitted font-size; an authored size is never replaced.
const SCALE_NEUTRAL_REFERENCE_FONT_SIZE_PT = 1;

function arrowEnd(end: ShapeRun['headEnd']): ArrowEnd | undefined {
  return end ? { type: end.type, w: end.w, len: end.len } : undefined;
}

function shapeStroke(shape: Readonly<ShapeRun>): Stroke | null {
  if (!shape.stroke || !shape.strokeWidth || shape.strokeWidth <= 0) return null;
  return {
    color: shape.stroke,
    width: shape.strokeWidth,
    ...(shape.strokeDash ? { dashStyle: shape.strokeDash } : {}),
    ...(shape.strokeCap ? { lineCap: shape.strokeCap } : {}),
    ...(arrowEnd(shape.headEnd) ? { headEnd: arrowEnd(shape.headEnd) } : {}),
    ...(arrowEnd(shape.tailEnd) ? { tailEnd: arrowEnd(shape.tailEnd) } : {}),
  };
}

export function planShapeDrawing(
  shape: Readonly<ShapeRun>,
  bounds: LayoutRect,
  text?: TextLayoutService,
  textPath?: Readonly<VmlTextPathAcquisitionInput>,
): ShapeDrawingPlanResult {
  const parserControlled = textPath !== undefined && (
    textPath.textPathOk !== undefined
    || textPath.on !== undefined
    || textPath.fitShape !== undefined
    || textPath.fitPath !== undefined
    || textPath.trim !== undefined
    || textPath.xScale !== undefined
  );
  // The stable public ShapeRun predates the private VML control projection.
  // Absence therefore keeps its historical visible+fitshape interpretation;
  // parser-created values always carry explicit false defaults and follow VML.
  const textPathEnabled = textPath !== undefined && (
    parserControlled
      ? textPath.textPathOk === true && textPath.on === true
      : true
  );
  if (textPathEnabled) {
    if (textPath.fitPath === true) {
      throw new Error('Unsupported VML textPath fitPath=true');
    }
    if (textPath.xScale === true) {
      throw new Error('Unsupported VML textPath xScale=true');
    }
    if (textPath.string.trim().length === 0) {
      return Object.freeze({ status: 'planned', command: Object.freeze({ kind: 'noop' }) });
    }
    if (!text) throw new Error('Shape textPath acquisition requires TextLayoutService');
    const fitShape = parserControlled ? textPath.fitShape === true : true;
    if (textPath.fontSizePt !== undefined
      && (!Number.isFinite(textPath.fontSizePt) || textPath.fontSizePt < 0)) {
      throw new RangeError('VML textPath fontSizePt must be finite and non-negative');
    }
    if (!fitShape && textPath.fontSizePt === undefined) {
      throw new Error('VML textPath fitShape=false requires an authored font-size');
    }
    if (textPath.fontSizePt === 0) {
      return Object.freeze({ status: 'planned', command: Object.freeze({ kind: 'noop' }) });
    }
    const fontSizePt = textPath.fontSizePt ?? SCALE_NEUTRAL_REFERENCE_FONT_SIZE_PT;
    const family = textPath.fontFamily ?? undefined;
    const shaped = text.shape({
      text: textPath.string,
      fontSizePt,
      fonts: {
        ascii: family,
        highAnsi: family,
        eastAsia: family,
        complexScript: family,
      },
      weight: textPath.bold ? 700 : 400,
      style: textPath.italic ? 'italic' : 'normal',
      measure: true,
    });
    if (textPath.trim === true && !shaped.inkBounds) {
      throw new Error('VML textPath trim=true requires glyph ink bounds');
    }
    // The retained source rectangle must choose one coherent metric domain:
    // trim removes the font's reserved box and uses tight actual ink, while
    // non-trim keeps the typographic advance/ascent/descent rectangle.
    const xMinPt = textPath.trim === true ? shaped.inkBounds?.xMinPt ?? 0 : 0;
    const xMaxPt = textPath.trim === true
      ? shaped.inkBounds?.xMaxPt ?? 0
      : shaped.advancePt;
    const ascentPt = textPath.trim === true ? shaped.inkBounds?.ascentPt ?? 0 : shaped.ascentPt;
    const descentPt = textPath.trim === true ? shaped.inkBounds?.descentPt ?? 0 : shaped.descentPt;
    const sourceBounds = {
      xPt: xMinPt,
      yPt: -ascentPt,
      widthPt: xMaxPt - xMinPt,
      heightPt: ascentPt + descentPt,
    };
    if (
      !Number.isFinite(shaped.advancePt)
      || !Number.isFinite(sourceBounds.widthPt)
      || sourceBounds.widthPt <= 0
      || !Number.isFinite(sourceBounds.heightPt)
      || sourceBounds.heightPt <= 0
      || shaped.spans.length === 0
    ) {
      throw new Error('Shape textPath acquisition produced degenerate metrics');
    }
    return Object.freeze({
      status: 'planned',
      command: snapshotPlainData({
        kind: 'watermark-text' as const,
        rect: { ...bounds },
        text: textPath.string,
        fill: shape.fill ? { ...shape.fill, ...(shape.fill.fillType === 'gradient'
          ? { stops: shape.fill.stops.map((stop) => ({ ...stop })) }
          : {}) } : null,
        opacity: Math.max(0, Math.min(1, shape.fillOpacity ?? 1)),
        rotationDeg: shape.rotation ?? 0,
        fitShape,
        fontSizePt,
        sourceBounds,
        spans: shaped.spans.map((span) => ({
          text: span.text,
          advancePt: span.advancePt,
          fontRoute: span.fontRoute,
          fontWeight: span.font.weight,
          fontStyle: span.font.style,
        })),
      }, 'VML textPath command'),
    });
  }
  const plan: DrawingMLShapePaintPlan = {
    rect: { x: bounds.xPt, y: bounds.yPt, w: bounds.widthPt, h: bounds.heightPt },
    geometry: shape.presetGeometry
      ? {
          kind: 'preset',
          name: shape.presetGeometry,
          adjustments: [...(shape.adjValues ?? [])],
        }
      : {
          kind: 'custom',
          subpaths: shape.subpaths.map((subpath) => subpath.map((command) => ({ ...command }))),
        },
    fill: shape.fill ? { ...shape.fill, ...(shape.fill.fillType === 'gradient'
      ? { stops: shape.fill.stops.map((stop) => ({ ...stop })) }
      : {}) } : null,
    stroke: shapeStroke(shape),
    transform: {
      rotationDeg: shape.rotation ?? 0,
      flipH: shape.flipH ?? false,
      flipV: shape.flipV ?? false,
    },
  };
  return Object.freeze({
    status: 'planned',
    command: snapshotPlainData(
      { kind: 'drawingml-shape', plan } as const,
      'DrawingML shape command',
    ),
  });
}
