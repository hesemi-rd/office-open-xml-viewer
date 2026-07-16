import type {
  LayoutRect,
  Matrix2DData,
  PointPt,
  SectionRegionCoordinateSpace,
  WritingMode,
} from './types.js';

export type PhysicalPageExtent = Readonly<{
  widthPt: number;
  heightPt: number;
}>;

export function writingModeFromTextDirection(textDirection: string): WritingMode {
  switch (textDirection) {
    case 'tb':
    case 'tbV':
    case 'lrTb':
    case 'lrTbV':
      return 'horizontal-tb';
    case 'rl':
    case 'rlV':
    case 'tbRl':
    case 'tbRlV':
      return 'vertical-rl';
    case 'btLr':
      // Compatibility rule `word-section-btlr-tbrl-page-frame`; glyph orientation is paint-owned.
      return 'vertical-rl';
    case 'lr':
    case 'lrV':
    case 'tbLrV':
      return 'vertical-lr';
    default:
      throw new RangeError(`Unsupported Transitional text direction ${JSON.stringify(textDirection)}`);
  }
}

function requirePage(page: PhysicalPageExtent): void {
  if (!Number.isFinite(page.widthPt) || !Number.isFinite(page.heightPt)
    || page.widthPt <= 0 || page.heightPt <= 0) {
    throw new RangeError('Physical page extents must be positive and finite');
  }
}

function requirePoint(point: PointPt): void {
  if (!Number.isFinite(point.xPt) || !Number.isFinite(point.yPt)) {
    throw new RangeError('Point coordinates must be finite');
  }
}

function requireMatrix(matrix: Matrix2DData): void {
  if (![matrix.a, matrix.b, matrix.c, matrix.d, matrix.e, matrix.f].every(Number.isFinite)) {
    throw new RangeError('Matrix coefficients must be finite');
  }
}

function requireRect(rect: LayoutRect): void {
  requirePoint(rect);
  if (!Number.isFinite(rect.widthPt) || !Number.isFinite(rect.heightPt)
    || rect.widthPt < 0 || rect.heightPt < 0) {
    throw new RangeError('Rectangle extents must be finite and non-negative');
  }
}

export function logicalPageExtent(
  physicalPage: PhysicalPageExtent,
  writingMode: WritingMode,
): PhysicalPageExtent {
  requirePage(physicalPage);
  switch (writingMode) {
    case 'horizontal-tb':
      return { widthPt: physicalPage.widthPt, heightPt: physicalPage.heightPt };
    case 'vertical-rl':
    case 'vertical-lr':
      return { widthPt: physicalPage.heightPt, heightPt: physicalPage.widthPt };
    default:
      throw new RangeError(`Unsupported writing mode ${String(writingMode)}`);
  }
}

export function uprightPhysicalExtent(
  logicalSectionExtent: PhysicalPageExtent,
  writingMode: WritingMode,
): PhysicalPageExtent {
  requirePage(logicalSectionExtent);
  switch (writingMode) {
    case 'horizontal-tb':
      return {
        widthPt: logicalSectionExtent.widthPt,
        heightPt: logicalSectionExtent.heightPt,
      };
    case 'vertical-rl':
    case 'vertical-lr':
      return {
        widthPt: logicalSectionExtent.heightPt,
        heightPt: logicalSectionExtent.widthPt,
      };
    default:
      throw new RangeError(`Unsupported writing mode ${String(writingMode)}`);
  }
}

export function logicalToPhysicalMatrix(
  writingMode: WritingMode,
  page: PhysicalPageExtent,
): Matrix2DData {
  requirePage(page);
  switch (writingMode) {
    case 'horizontal-tb':
      return { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 };
    case 'vertical-rl':
      return { a: 0, b: 1, c: -1, d: 0, e: page.widthPt, f: 0 };
    case 'vertical-lr':
      return { a: 0, b: 1, c: 1, d: 0, e: 0, f: 0 };
    default:
      throw new RangeError(`Unsupported writing mode ${String(writingMode)}`);
  }
}

export function physicalToLogicalMatrix(
  writingMode: WritingMode,
  page: PhysicalPageExtent,
): Matrix2DData {
  requirePage(page);
  switch (writingMode) {
    case 'horizontal-tb':
      return { a: 1, b: 0, c: 0, d: 1, e: 0, f: 0 };
    case 'vertical-rl':
      return { a: 0, b: -1, c: 1, d: 0, e: 0, f: page.widthPt };
    case 'vertical-lr':
      return { a: 0, b: 1, c: 1, d: 0, e: 0, f: 0 };
    default:
      throw new RangeError(`Unsupported writing mode ${String(writingMode)}`);
  }
}

export function transformPoint(matrix: Matrix2DData, point: PointPt): PointPt {
  requireMatrix(matrix);
  requirePoint(point);
  return {
    xPt: matrix.a * point.xPt + matrix.c * point.yPt + matrix.e,
    yPt: matrix.b * point.xPt + matrix.d * point.yPt + matrix.f,
  };
}

export function transformRect(matrix: Matrix2DData, rect: LayoutRect): LayoutRect {
  requireRect(rect);
  const corners = [
    transformPoint(matrix, rect),
    transformPoint(matrix, { xPt: rect.xPt + rect.widthPt, yPt: rect.yPt }),
    transformPoint(matrix, { xPt: rect.xPt, yPt: rect.yPt + rect.heightPt }),
    transformPoint(matrix, {
      xPt: rect.xPt + rect.widthPt,
      yPt: rect.yPt + rect.heightPt,
    }),
  ];
  const xs = corners.map(({ xPt }) => xPt);
  const ys = corners.map(({ yPt }) => yPt);
  const xPt = Math.min(...xs);
  const yPt = Math.min(...ys);
  return {
    xPt,
    yPt,
    widthPt: Math.max(...xs) - xPt,
    heightPt: Math.max(...ys) - yPt,
  };
}

export function createSectionRegionCoordinateSpace(
  writingMode: WritingMode,
  page: PhysicalPageExtent,
): SectionRegionCoordinateSpace {
  return {
    writingMode,
    logicalToPhysical: logicalToPhysicalMatrix(writingMode, page),
    physicalToLogical: physicalToLogicalMatrix(writingMode, page),
  };
}
