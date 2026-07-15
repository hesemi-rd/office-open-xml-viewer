import { resolveFloatOverlap, type FloatRect } from '../float-layout.js';
import type {
  FloatRegistryEntryPt,
  FloatRegistryDeltaPt,
  FloatRegistrySnapshotPt,
  FloatingTablePlacementLayout,
  FloatingTablePositionInput,
  FloatingTableReferenceFramesPt,
  ResolvedFloatingTablePlacementLayout,
} from './types.js';

export interface FloatingTablePlacementTransaction {
  readonly coordinateSpace: FloatRegistrySnapshotPt['coordinateSpace'];
  readonly flowDomainId: string;
  readonly base: readonly FloatRegistryEntryPt[];
  readonly delta: readonly FloatRegistryEntryPt[];
  readonly nextParagraphId: number;
}

function endX(frame: FloatingTableReferenceFramesPt['page']): number {
  return frame.xPt + frame.widthPt;
}

function endY(frame: FloatingTableReferenceFramesPt['page']): number {
  return frame.yPt + frame.heightPt;
}

function alignedHorizontal(spec: string, start: number, end: number, size: number): number {
  if (spec === 'center') return start + (end - start - size) / 2;
  if (spec === 'right' || spec === 'outside') return end - size;
  return start;
}

function alignedVertical(spec: string, start: number, end: number, size: number): number {
  if (spec === 'center') return start + (end - start - size) / 2;
  if (spec === 'bottom' || spec === 'outside') return end - size;
  return start;
}

export function resolvePointSpaceFloatingTableBoxPt(
  positioning: FloatingTablePositionInput,
  frames: FloatingTableReferenceFramesPt,
  widthPt: number,
  heightPt: number,
  skipVerticalClamp = false,
): Readonly<{ x: number; y: number; w: number; h: number }> {
  const horizontalFrame = positioning.horzSpecified
    ? positioning.horzAnchor === 'page'
      ? frames.page
      : positioning.horzAnchor === 'margin' ? frames.margin : frames.text
    : frames.text;
  const verticalFrame = positioning.vertAnchor === 'page'
    ? frames.page
    : positioning.vertAnchor === 'margin' ? frames.margin : frames.text;
  const x = positioning.xAlign
    ? alignedHorizontal(positioning.xAlign, horizontalFrame.xPt, endX(horizontalFrame), widthPt)
    : horizontalFrame.xPt + positioning.xPt;
  let y = positioning.yAlign && positioning.vertAnchor !== 'text'
    ? alignedVertical(positioning.yAlign, verticalFrame.yPt, endY(verticalFrame), heightPt)
    : verticalFrame.yPt + positioning.yPt;
  if (!skipVerticalClamp
    && (positioning.vertAnchor === 'page' || positioning.vertAnchor === 'margin')
    && y + heightPt > endY(verticalFrame)) {
    // ECMA-376 leaves mid-cell/page overflow policy open; retain the explicit
    // bottom-clamp already used by this renderer without claiming Word normativity.
    y = Math.max(verticalFrame.yPt, endY(verticalFrame) - heightPt);
  }
  return Object.freeze({ x, y, w: widthPt, h: heightPt });
}

export function resolveFloatingTableBoxPt(
  positioning: FloatingTablePositionInput,
  frames: FloatingTableReferenceFramesPt,
  widthPt: number,
  heightPt: number,
): Readonly<{ x: number; y: number; w: number; h: number }> {
  return resolvePointSpaceFloatingTableBoxPt(positioning, frames, widthPt, heightPt);
}

function resolvedPlacement(
  placement: FloatingTablePlacementLayout,
  xPt: number,
  yPt: number,
): ResolvedFloatingTablePlacementLayout {
  const widthPt = placement.child.columnWidthsPt.reduce((sum, width) => sum + width, 0);
  const heightPt = placement.child.advancePt;
  const positioning = placement.positioning;
  const bounds = Object.freeze({ xPt, yPt, widthPt, heightPt });
  const exclusionBounds = Object.freeze({
    xPt: xPt - positioning.leftFromTextPt,
    yPt: yPt - positioning.topFromTextPt,
    widthPt: widthPt + positioning.leftFromTextPt + positioning.rightFromTextPt,
    heightPt: heightPt + positioning.topFromTextPt + positioning.bottomFromTextPt,
  });
  return Object.freeze({
    kind: 'resolved-floating-table-placement',
    occurrenceId: placement.occurrenceId,
    xPt,
    yPt,
    bounds,
    exclusionBounds,
    overlap: placement.overlap,
    child: placement.child,
    source: placement,
  });
}

export function resolveFloatingTablePlacement(
  placement: FloatingTablePlacementLayout,
  frames: FloatingTableReferenceFramesPt,
): ResolvedFloatingTablePlacementLayout {
  const widthPt = placement.child.columnWidthsPt.reduce((sum, width) => sum + width, 0);
  const heightPt = placement.child.advancePt;
  const raw = resolveFloatingTableBoxPt(placement.positioning, frames, widthPt, heightPt);
  const usesTextX = !placement.positioning.horzSpecified
    || (placement.positioning.horzAnchor !== 'page'
      && placement.positioning.horzAnchor !== 'margin');
  const usesTextY = placement.positioning.vertAnchor !== 'page'
    && placement.positioning.vertAnchor !== 'margin';
  return resolvedPlacement(
    placement,
    usesTextX && placement.acquiredTextOffsetPt
      ? frames.text.xPt + placement.acquiredTextOffsetPt.xPt : raw.x,
    usesTextY && placement.acquiredTextOffsetPt
      ? frames.text.yPt + placement.acquiredTextOffsetPt.yPt : raw.y,
  );
}

function registryFloat(entry: FloatRegistryEntryPt): FloatRect {
  return {
    kind: entry.kind,
    mode: 'square',
    imageKey: entry.occurrenceId,
    imageX: entry.bounds.xPt,
    imageY: entry.bounds.yPt,
    imageW: entry.bounds.widthPt,
    imageH: entry.bounds.heightPt,
    xLeft: entry.exclusionBounds.xPt,
    xRight: entry.exclusionBounds.xPt + entry.exclusionBounds.widthPt,
    yTop: entry.exclusionBounds.yPt,
    yBottom: entry.exclusionBounds.yPt + entry.exclusionBounds.heightPt,
    side: 'bothSides',
    distLeft: entry.bounds.xPt - entry.exclusionBounds.xPt,
    distRight: endX(entry.exclusionBounds) - endX(entry.bounds),
    distTop: entry.bounds.yPt - entry.exclusionBounds.yPt,
    distBottom: endY(entry.exclusionBounds) - endY(entry.bounds),
    paraId: entry.paragraphId,
    drawn: true,
  };
}

export function floatingTableRegistryDelta(
  snapshot: FloatRegistrySnapshotPt,
  entries: readonly FloatRegistryEntryPt[],
  nextParagraphId: number,
): FloatRegistryDeltaPt {
  return Object.freeze({
    coordinateSpace: snapshot.coordinateSpace,
    flowDomainId: snapshot.flowDomainId,
    baseNextParagraphId: snapshot.nextParagraphId,
    nextParagraphId,
    entries: Object.freeze([...entries]),
  });
}

/** Convert the exact probed point-space delta into the renderer's scaled registry. */
export function commitFloatingTableRegistryDelta(
  delta: FloatRegistryDeltaPt,
  current: Readonly<{
    coordinateSpace: FloatRegistrySnapshotPt['coordinateSpace'];
    flowDomainId: string;
    nextParagraphId: number;
    occurrenceIds: readonly string[];
  }>,
  scale: number,
): Readonly<{ entries: readonly FloatRect[]; nextParagraphId: number }> {
  if (current.coordinateSpace !== delta.coordinateSpace
    || current.flowDomainId !== delta.flowDomainId
    || current.nextParagraphId !== delta.baseNextParagraphId) {
    throw new Error('Floating table registry delta base/domain mismatch');
  }
  const existingIds = new Set(current.occurrenceIds);
  if (delta.entries.some((entry) => existingIds.has(entry.occurrenceId))) {
    throw new Error('Floating table registry delta was already committed');
  }
  if (delta.nextParagraphId !== delta.baseNextParagraphId + delta.entries.length) {
    throw new Error('Floating table registry delta sequence mismatch');
  }
  return Object.freeze({
    entries: Object.freeze(delta.entries.map((entry) => {
      const pointRect = registryFloat(entry);
      return Object.freeze({
        ...pointRect,
        imageX: pointRect.imageX * scale,
        imageY: pointRect.imageY * scale,
        imageW: pointRect.imageW * scale,
        imageH: pointRect.imageH * scale,
        xLeft: pointRect.xLeft * scale,
        xRight: pointRect.xRight * scale,
        yTop: pointRect.yTop * scale,
        yBottom: pointRect.yBottom * scale,
        distLeft: pointRect.distLeft * scale,
        distRight: pointRect.distRight * scale,
        distTop: pointRect.distTop * scale,
        distBottom: pointRect.distBottom * scale,
      });
    })),
    nextParagraphId: delta.nextParagraphId,
  });
}

export function beginFloatingTablePlacementTransaction(
  base: readonly FloatRegistryEntryPt[],
  nextParagraphId: number,
  coordinateSpace: FloatRegistrySnapshotPt['coordinateSpace'] = 'logical-page-points',
  flowDomainId = 'logical-page',
): FloatingTablePlacementTransaction {
  const ids = new Set<string>();
  for (const entry of base) {
    if (ids.has(entry.occurrenceId)) {
      throw new Error(`Duplicate float registry occurrence: ${entry.occurrenceId}`);
    }
    ids.add(entry.occurrenceId);
  }
  return Object.freeze({
    coordinateSpace,
    flowDomainId,
    base: Object.freeze([...base]),
    delta: Object.freeze([]),
    nextParagraphId,
  });
}

export function resolveFloatingTablePlacementInTransaction(
  placement: FloatingTablePlacementLayout,
  frames: FloatingTableReferenceFramesPt,
  transaction: FloatingTablePlacementTransaction,
): Readonly<{
  placement: ResolvedFloatingTablePlacementLayout;
  transaction: FloatingTablePlacementTransaction;
}> {
  const registry = [...transaction.base, ...transaction.delta];
  const existing = registry.find((entry) => entry.occurrenceId === placement.occurrenceId);
  if (existing) {
    return Object.freeze({
      placement: Object.freeze({
        ...resolvedPlacement(placement, existing.bounds.xPt, existing.bounds.yPt),
        bounds: existing.bounds,
        exclusionBounds: existing.exclusionBounds,
      }),
      transaction,
    });
  }
  const initial = resolveFloatingTablePlacement(placement, frames);
  const positioning = placement.positioning;
  const position = resolveFloatOverlap(
    initial.xPt, initial.yPt, initial.bounds.widthPt, initial.bounds.heightPt,
    positioning.leftFromTextPt, positioning.rightFromTextPt,
    positioning.topFromTextPt, positioning.bottomFromTextPt,
    transaction.nextParagraphId, placement.overlap !== 'never', 'table',
    endX(frames.page), registry.map(registryFloat),
  );
  const finalPlacement = resolvedPlacement(placement, position.x, position.y);
  const entry = Object.freeze({
    kind: 'table' as const,
    occurrenceId: placement.occurrenceId,
    paragraphId: transaction.nextParagraphId,
    bounds: finalPlacement.bounds,
    exclusionBounds: finalPlacement.exclusionBounds,
  });
  return Object.freeze({
    placement: finalPlacement,
    transaction: Object.freeze({
      coordinateSpace: transaction.coordinateSpace,
      flowDomainId: transaction.flowDomainId,
      base: transaction.base,
      delta: Object.freeze([...transaction.delta, entry]),
      nextParagraphId: transaction.nextParagraphId + 1,
    }),
  });
}

export function resolveFloatingTablePlacementsInFreshRegistry(
  placements: readonly FloatingTablePlacementLayout[],
  framesFor: (placement: FloatingTablePlacementLayout) => FloatingTableReferenceFramesPt,
  coordinateSpace: FloatRegistrySnapshotPt['coordinateSpace'],
  flowDomainId: string,
): Readonly<{
  placements: readonly ResolvedFloatingTablePlacementLayout[];
  transaction: FloatingTablePlacementTransaction;
}> {
  let transaction = beginFloatingTablePlacementTransaction(
    [], 0, coordinateSpace, flowDomainId,
  );
  const resolved: ResolvedFloatingTablePlacementLayout[] = [];
  for (const placement of placements) {
    const resolution = resolveFloatingTablePlacementInTransaction(
      placement, framesFor(placement), transaction,
    );
    resolved.push(resolution.placement);
    transaction = resolution.transaction;
  }
  return Object.freeze({ placements: Object.freeze(resolved), transaction });
}
