import { LayoutInvariantError } from './diagnostics.js';
import {
  createSectionRegionCoordinateSpace,
  logicalPageExtent,
  transformRect,
  uprightPhysicalExtent,
  writingModeFromTextDirection,
} from './coordinate-space.js';
import { orderedPagePaintNodes, pageLayerNodes, PageGraphError } from './page-graph.js';
import {
  derivePageBookmarkStarts,
  sectionLayoutContextsEqual,
} from './page-factory.js';
import type {
  DeepReadonly,
  DocumentLayout,
  DrawingPaintCommand,
  DrawingLayout,
  FlowDomain,
  LayoutRect,
  LayoutPage,
  PageSectionRegion,
  PaintNode,
  PointPt,
  SectionRegionCoordinateSpace,
  WritingMode,
} from './types.js';

function assertPlainData(value: unknown, path: string, ancestors = new WeakSet<object>()): void {
  if (value === null || typeof value === 'string' || typeof value === 'boolean') return;
  if (typeof value === 'number') {
    if (!Number.isFinite(value)) {
      throw new LayoutInvariantError('INVALID_GEOMETRY', `${path} is not finite`);
    }
    return;
  }
  if (typeof value !== 'object') {
    throw new LayoutInvariantError('INVALID_GEOMETRY', `${path} contains ${typeof value}`);
  }
  if (ancestors.has(value)) {
    throw new LayoutInvariantError('INVALID_GEOMETRY', `${path} contains a cycle`);
  }

  ancestors.add(value);
  try {
    if (Array.isArray(value)) {
      let indexCount = 0;
      for (const key of Reflect.ownKeys(value)) {
        if (key === 'length') continue;
        if (typeof key !== 'string') {
          throw new LayoutInvariantError('INVALID_GEOMETRY', `${path} has a symbol key`);
        }
        const index = Number(key);
        if (!Number.isInteger(index) || index < 0 || String(index) !== key || index >= value.length) {
          throw new LayoutInvariantError('INVALID_GEOMETRY', `${path}.${key} is not an array index`);
        }
        const descriptor = Object.getOwnPropertyDescriptor(value, key);
        if (!descriptor?.enumerable || !('value' in descriptor)) {
          throw new LayoutInvariantError('INVALID_GEOMETRY', `${path}[${key}] is not plain data`);
        }
        assertPlainData(descriptor.value, `${path}[${key}]`, ancestors);
        indexCount += 1;
      }
      if (indexCount !== value.length) {
        throw new LayoutInvariantError('INVALID_GEOMETRY', `${path} is sparse`);
      }
      return;
    }

    const prototype = Object.getPrototypeOf(value);
    if (prototype !== Object.prototype && prototype !== null) {
      throw new LayoutInvariantError('INVALID_GEOMETRY', `${path} is not a plain record`);
    }
    for (const key of Reflect.ownKeys(value)) {
      if (typeof key !== 'string') {
        throw new LayoutInvariantError('INVALID_GEOMETRY', `${path} has a symbol key`);
      }
      const descriptor = Object.getOwnPropertyDescriptor(value, key);
      if (!descriptor?.enumerable || !('value' in descriptor)) {
        throw new LayoutInvariantError('INVALID_GEOMETRY', `${path}.${key} is not plain data`);
      }
      assertPlainData(descriptor.value, `${path}.${key}`, ancestors);
    }
  } finally {
    ancestors.delete(value);
  }
}

function requireFinite(value: number, path: string): void {
  if (!Number.isFinite(value)) {
    throw new LayoutInvariantError('INVALID_GEOMETRY', `${path} is not finite`);
  }
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return value !== null && typeof value === 'object' && !Array.isArray(value);
}

function requirePoint(point: PointPt, path: string): void {
  if (!isRecord(point)) {
    throw new LayoutInvariantError('INVALID_GEOMETRY', `${path} is not a point`);
  }
  requireFinite(point.xPt, `${path}.xPt`);
  requireFinite(point.yPt, `${path}.yPt`);
}

function requireRect(rect: LayoutRect, path: string): void {
  requirePoint(rect, path);
  requireFinite(rect.widthPt, `${path}.widthPt`);
  requireFinite(rect.heightPt, `${path}.heightPt`);
  if (rect.widthPt < 0 || rect.heightPt < 0) {
    throw new LayoutInvariantError('INVALID_GEOMETRY', `${path} has a negative extent`);
  }
}

function requireMatrix(matrix: unknown, path: string): asserts matrix is SectionRegionCoordinateSpace['logicalToPhysical'] {
  if (!isRecord(matrix)) {
    throw new LayoutInvariantError('INVALID_GEOMETRY', `${path} is not a matrix`);
  }
  for (const coefficient of ['a', 'b', 'c', 'd', 'e', 'f'] as const) {
    requireFinite(matrix[coefficient] as number, `${path}.${coefficient}`);
  }
}

function requireWritingMode(value: unknown, path: string): asserts value is WritingMode {
  if (value !== 'horizontal-tb' && value !== 'vertical-rl' && value !== 'vertical-lr') {
    throw new LayoutInvariantError('INVALID_GEOMETRY', `${path} is unsupported`);
  }
}

function requireCoordinateSpace(
  value: unknown,
  path: string,
): asserts value is SectionRegionCoordinateSpace {
  if (!isRecord(value)) {
    throw new LayoutInvariantError('INVALID_GEOMETRY', `${path} is not a coordinate space`);
  }
  requireWritingMode(value.writingMode, `${path}.writingMode`);
  requireMatrix(value.logicalToPhysical, `${path}.logicalToPhysical`);
  requireMatrix(value.physicalToLogical, `${path}.physicalToLogical`);
}

function requireDrawingMLShapePlan(
  command: Extract<DrawingPaintCommand, { kind: 'drawingml-shape' }>,
  path: string,
): void {
  const { plan } = command;
  assertPlainData(plan, `${path}.plan`);
  requireFinite(plan.rect.x, `${path}.plan.rect.x`);
  requireFinite(plan.rect.y, `${path}.plan.rect.y`);
  requireFinite(plan.rect.w, `${path}.plan.rect.w`);
  requireFinite(plan.rect.h, `${path}.plan.rect.h`);
  if (plan.rect.w < 0 || plan.rect.h < 0) {
    throw new LayoutInvariantError('INVALID_GEOMETRY', `${path}.plan.rect has a negative extent`);
  }
  requireFinite(plan.transform.rotationDeg, `${path}.plan.transform.rotationDeg`);
  if (plan.geometry.kind === 'preset') {
    if (plan.geometry.name.length === 0) {
      throw new LayoutInvariantError('INVALID_GEOMETRY', `${path}.plan.geometry.name is empty`);
    }
    plan.geometry.adjustments.forEach((adjustment, index) => {
      if (adjustment !== null) {
        requireFinite(adjustment, `${path}.plan.geometry.adjustments[${index}]`);
      }
    });
  } else {
    plan.geometry.subpaths.forEach((subpath, subpathIndex) => {
      subpath.forEach((pathCommand, commandIndex) => {
        if (pathCommand.cmd.length === 0) {
          throw new LayoutInvariantError(
            'INVALID_GEOMETRY',
            `${path}.plan.geometry.subpaths[${subpathIndex}][${commandIndex}].cmd is empty`,
          );
        }
      });
    });
  }
  if (plan.stroke) {
    requireFinite(plan.stroke.width, `${path}.plan.stroke.width`);
    if (plan.stroke.width < 0) {
      throw new LayoutInvariantError('INVALID_GEOMETRY', `${path}.plan.stroke.width is negative`);
    }
  }
}

function overlaps(a: LayoutRect, b: LayoutRect): boolean {
  return a.xPt < b.xPt + b.widthPt
    && b.xPt < a.xPt + a.widthPt
    && a.yPt < b.yPt + b.heightPt
    && b.yPt < a.yPt + a.heightPt;
}

function contains(outer: LayoutRect, inner: LayoutRect): boolean {
  return inner.xPt >= outer.xPt
    && inner.yPt >= outer.yPt
    && inner.xPt + inner.widthPt <= outer.xPt + outer.widthPt
    && inner.yPt + inner.heightPt <= outer.yPt + outer.heightPt;
}

function equalRect(left: LayoutRect, right: LayoutRect): boolean {
  return left.xPt === right.xPt
    && left.yPt === right.yPt
    && left.widthPt === right.widthPt
    && left.heightPt === right.heightPt;
}

function equalMatrix(
  left: PageSectionRegion['coordinateSpace']['logicalToPhysical'],
  right: PageSectionRegion['coordinateSpace']['logicalToPhysical'],
): boolean {
  return left.a === right.a && left.b === right.b && left.c === right.c
    && left.d === right.d && left.e === right.e && left.f === right.f;
}

function retainUniqueNodeId(
  id: string,
  pageIds: Set<string>,
  documentIds: Set<string>,
): void {
  if (documentIds.has(id)) {
    throw new LayoutInvariantError('INVALID_REFERENCE', `duplicate retained node id ${id}`);
  }
  documentIds.add(id);
  pageIds.add(id);
}

function collectRetainedNodeIds(
  node: PaintNode,
  pageIds: Set<string>,
  documentIds: Set<string>,
): void {
  retainUniqueNodeId(node.id, pageIds, documentIds);
  if (node.kind === 'paragraph') {
    node.drawings.forEach((drawing) =>
      collectRetainedNodeIds(drawing, pageIds, documentIds));
    node.textBoxes.forEach((textBox) =>
      collectRetainedNodeIds(textBox, pageIds, documentIds));
    return;
  }
  if (node.kind === 'table') {
    node.rows.forEach((row) => {
      retainUniqueNodeId(row.id, pageIds, documentIds);
      row.cells.forEach((cell) => {
        retainUniqueNodeId(cell.id, pageIds, documentIds);
        cell.blocks.forEach((block) =>
          collectRetainedNodeIds(block.layout, pageIds, documentIds));
      });
    });
    return;
  }
  if (node.kind === 'textbox') {
    node.paragraphs.forEach((paragraph) =>
      collectRetainedNodeIds(paragraph, pageIds, documentIds));
  }
}

function requireDrawingGeometry(node: DrawingLayout, path: string): void {
  if (node.transform) {
    for (const key of ['a', 'b', 'c', 'd', 'e', 'f'] as const) {
      requireFinite(node.transform[key], `${path}.transform.${key}`);
    }
  }
  if (node.clip?.kind === 'rect') requireRect(node.clip.rect, `${path}.clip.rect`);
  if (node.clip?.kind === 'polygon') {
    node.clip.points.forEach((point, index) => requirePoint(point, `${path}.clip.points[${index}]`));
  }
  node.commands.forEach((command, index) => {
    const commandPath = `${path}.commands[${index}]`;
    if (command.kind === 'noop') return;
    if (command.kind === 'drawingml-shape') {
      requireDrawingMLShapePlan(command, commandPath);
      return;
    }
    requireRect(command.rect, `${commandPath}.rect`);
    if (command.kind === 'stroke-rect') {
      requireFinite(command.lineWidthPt, `${commandPath}.lineWidthPt`);
      command.dashPt.forEach((dash, dashIndex) =>
        requireFinite(dash, `${commandPath}.dashPt[${dashIndex}]`));
    }
    if (command.kind === 'text') {
      requireFinite(command.fontSizePt, `${commandPath}.fontSizePt`);
      requireFinite(command.fontWeight, `${commandPath}.fontWeight`);
    }
    if (command.kind === 'watermark-text') {
      requireRect(command.sourceBounds, `${commandPath}.sourceBounds`);
      if (command.sourceBounds.widthPt <= 0 || command.sourceBounds.heightPt <= 0) {
        throw new LayoutInvariantError(
          'INVALID_GEOMETRY',
          `${commandPath}.sourceBounds must have positive extents`,
        );
      }
      requireFinite(command.opacity, `${commandPath}.opacity`);
      requireFinite(command.rotationDeg, `${commandPath}.rotationDeg`);
      requireFinite(command.fontSizePt, `${commandPath}.fontSizePt`);
      if (command.opacity < 0 || command.opacity > 1 || command.fontSizePt <= 0) {
        throw new LayoutInvariantError('INVALID_GEOMETRY', `${commandPath} has invalid textPath paint metrics`);
      }
      command.spans.forEach((span, spanIndex) => {
        requireFinite(span.advancePt, `${commandPath}.spans[${spanIndex}].advancePt`);
        requireFinite(span.fontWeight, `${commandPath}.spans[${spanIndex}].fontWeight`);
      });
    }
  });
}

function assertDocumentLayoutUnchecked(layout: DocumentLayout): void {
  assertPlainData(layout, 'layout');
  const documentRetainedNodeIds = new Set<string>();
  layout.pages.forEach((page, pageIndex) => {
    if (!Number.isInteger(page.pageIndex) || page.pageIndex !== pageIndex) {
      throw new LayoutInvariantError(
        'INVALID_REFERENCE',
        `pages[${pageIndex}] has invalid page index ${page.pageIndex}`,
      );
    }
    requireRect(page.geometry, `pages[${pageIndex}].geometry`);
    requireFinite(page.geometry.contentTopPt, `pages[${pageIndex}].geometry.contentTopPt`);
    requireFinite(page.geometry.contentBottomPt, `pages[${pageIndex}].geometry.contentBottomPt`);
    if (
      page.geometry.widthPt <= 0
      || page.geometry.heightPt <= 0
      || page.geometry.contentTopPt < 0
      || page.geometry.contentTopPt > page.geometry.contentBottomPt
      || page.geometry.contentBottomPt > page.geometry.heightPt
    ) {
      throw new LayoutInvariantError(
        'INVALID_GEOMETRY',
        `pages[${pageIndex}] has invalid effective page edges`,
      );
    }

    const domains = new Map<string, FlowDomain>();
    page.flowDomains.forEach((domain, domainIndex) => {
      requireRect(
        domain.logicalBounds,
        `pages[${pageIndex}].flowDomains[${domainIndex}].logicalBounds`,
      );
      requireRect(
        domain.physicalBounds,
        `pages[${pageIndex}].flowDomains[${domainIndex}].physicalBounds`,
      );
      if (domains.has(domain.id)) {
        throw new LayoutInvariantError('INVALID_REFERENCE', `duplicate flow domain ${domain.id}`);
      }
      domains.set(domain.id, domain);
    });

    if (page.parityBlank && (
      page.flowDomains.length > 0
      || (page.sectionRegions?.length ?? 0) > 0
      || pageLayerNodes(page).length > 0
      || page.layers.paintOrder.length > 0
      || page.readingOrder.length > 0
      || (page.bookmarkStarts?.length ?? 0) > 0
    )) {
      throw new LayoutInvariantError(
        'INVALID_REFERENCE',
        `pages[${pageIndex}] parity blank retains page content`,
      );
    }

    const sectionOccurrenceIds = new Set<string>();
    if (page.sectionOccurrenceId !== undefined) {
      if (page.sectionOccurrenceId.length === 0) {
        throw new LayoutInvariantError(
          'INVALID_REFERENCE',
          `pages[${pageIndex}] has an empty section occurrence id`,
        );
      }
      sectionOccurrenceIds.add(page.sectionOccurrenceId);
    }

    const regionByDomain = new Map<string, PageSectionRegion>();
    if (page.sectionRegions) {
      const regionIds = new Set<string>();
      const occurrenceIds = new Set<string>();
      const bodyOwnership = new Map<string, number>();
      let priorBlockEndPt = 0;
      let pageWritingMode: WritingMode | undefined;
      page.sectionRegions.forEach((region, regionIndex) => {
        const path = `pages[${pageIndex}].sectionRegions[${regionIndex}]`;
        if (region.id.length === 0 || regionIds.has(region.id)) {
          throw new LayoutInvariantError('INVALID_REFERENCE', `${path} has an invalid region id`);
        }
        regionIds.add(region.id);
        if (region.sectionOccurrenceId.length === 0) {
          throw new LayoutInvariantError(
            'INVALID_REFERENCE',
            `${path} has an empty section occurrence id`,
          );
        }
        if (occurrenceIds.has(region.sectionOccurrenceId)) {
          throw new LayoutInvariantError(
            'INVALID_REFERENCE',
            `${path} has a duplicate section occurrence id`,
          );
        }
        occurrenceIds.add(region.sectionOccurrenceId);
        sectionOccurrenceIds.add(region.sectionOccurrenceId);
        requireCoordinateSpace(region.coordinateSpace, `${path}.coordinateSpace`);
        const writingMode = region.coordinateSpace.writingMode;
        if (pageWritingMode !== undefined && pageWritingMode !== writingMode) {
          throw new LayoutInvariantError(
            'INVALID_GEOMETRY',
            `${path} mixes coordinate systems on one physical page`,
          );
        }
        pageWritingMode = writingMode;
        let sectionWritingMode: WritingMode;
        try {
          sectionWritingMode = writingModeFromTextDirection(region.section.textDirection);
        } catch (error) {
          throw new LayoutInvariantError(
            'INVALID_GEOMETRY',
            `${path}.section.textDirection is unsupported: ${(error as Error).message}`,
          );
        }
        if (writingMode !== sectionWritingMode) {
          throw new LayoutInvariantError(
            'INVALID_GEOMETRY',
            `${path} writing mode contradicts its section text direction`,
          );
        }
        const logicalExtent = logicalPageExtent(page.geometry, writingMode);
        const physicalExtent = uprightPhysicalExtent({
          widthPt: region.section.geometry.pageWidth,
          heightPt: region.section.geometry.pageHeight,
        }, writingMode);
        if (physicalExtent.widthPt !== page.geometry.widthPt
          || physicalExtent.heightPt !== page.geometry.heightPt) {
          throw new LayoutInvariantError(
            'INVALID_GEOMETRY',
            `${path} section geometry does not match the upright physical page`,
          );
        }
        requireFinite(region.blockStartPt, `${path}.blockStartPt`);
        requireFinite(region.blockEndPt, `${path}.blockEndPt`);
        if (region.blockStartPt < 0
          || region.blockStartPt < priorBlockEndPt
          || region.blockEndPt <= region.blockStartPt
          || region.blockEndPt > logicalExtent.heightPt) {
          throw new LayoutInvariantError('INVALID_GEOMETRY', `${path} has an invalid block interval`);
        }
        priorBlockEndPt = region.blockEndPt;
        const expectedCoordinateSpace = createSectionRegionCoordinateSpace(
          region.coordinateSpace.writingMode,
          page.geometry,
        );
        if (!equalMatrix(
          region.coordinateSpace.logicalToPhysical,
          expectedCoordinateSpace.logicalToPhysical,
        ) || !equalMatrix(
          region.coordinateSpace.physicalToLogical,
          expectedCoordinateSpace.physicalToLogical,
        )) {
          throw new LayoutInvariantError('INVALID_GEOMETRY', `${path} has an invalid coordinate transform`);
        }
        if (region.flowDomainIds.length !== region.section.columns.length) {
          throw new LayoutInvariantError('INVALID_GEOMETRY', `${path} columns contradict its section`);
        }
        let priorInlineEndPt = 0;
        region.flowDomainIds.forEach((domainId, columnIndex) => {
          const domain = domains.get(domainId);
          if (!domain) {
            throw new LayoutInvariantError('INVALID_REFERENCE', `${path} references missing flow domain ${domainId}`);
          }
          if (domain.kind !== 'body') {
            throw new LayoutInvariantError('INVALID_REFERENCE', `${path} owns non-body flow domain ${domainId}`);
          }
          bodyOwnership.set(domainId, (bodyOwnership.get(domainId) ?? 0) + 1);
          regionByDomain.set(domainId, region);
          const bounds = domain.logicalBounds;
          const sectionColumn = region.section.columns[columnIndex];
          if (bounds.widthPt <= 0 || bounds.heightPt <= 0
            || bounds.yPt !== region.blockStartPt
            || bounds.yPt + bounds.heightPt !== region.blockEndPt
            || bounds.xPt < 0
            || bounds.xPt < priorInlineEndPt
            || bounds.xPt + bounds.widthPt > logicalExtent.widthPt
            || sectionColumn === undefined
            || bounds.xPt !== sectionColumn.xPt
            || bounds.widthPt !== sectionColumn.wPt) {
            throw new LayoutInvariantError(
              'INVALID_GEOMETRY',
              `${domainId} is not the section column's positive logical region`,
            );
          }
          priorInlineEndPt = bounds.xPt + bounds.widthPt;
          const expectedPhysicalBounds = transformRect(
            region.coordinateSpace.logicalToPhysical,
            domain.logicalBounds,
          );
          if (!equalRect(expectedPhysicalBounds, domain.physicalBounds)) {
            throw new LayoutInvariantError(
              'INVALID_GEOMETRY',
              `${domainId} physical bounds do not match its section region transform`,
            );
          }
          if (!contains(page.geometry, domain.physicalBounds)) {
            throw new LayoutInvariantError(
              'INVALID_GEOMETRY',
              `${domainId} physical bounds leave the upright physical page`,
            );
          }
        });
      });
      page.flowDomains.filter((domain) => domain.kind === 'body').forEach((domain) => {
        if (bodyOwnership.get(domain.id) !== 1) {
          throw new LayoutInvariantError(
            'INVALID_REFERENCE',
            `${domain.id} has invalid section region ownership`,
          );
        }
      });
      if (!page.parityBlank && page.sectionRegions.length > 0
      ) {
        const firstRegion = page.sectionRegions[0]!;
        if (page.sectionOccurrenceId !== firstRegion.sectionOccurrenceId) {
          throw new LayoutInvariantError(
            'INVALID_REFERENCE',
            `pages[${pageIndex}] page-start section occurrence does not match its first region`,
          );
        }
        if (!sectionLayoutContextsEqual(page.section, firstRegion.section)) {
          throw new LayoutInvariantError(
            'INVALID_GEOMETRY',
            `pages[${pageIndex}] page-start section facts do not match its first region`,
          );
        }
      }
    }
    for (const domain of page.flowDomains) {
      if (!regionByDomain.has(domain.id) && !equalRect(domain.logicalBounds, domain.physicalBounds)) {
        throw new LayoutInvariantError(
          'INVALID_GEOMETRY',
          `${domain.id} has unequal logical and physical bounds without a section region`,
        );
      }
    }

    if (page.pageNumber) {
      requireFinite(page.pageNumber.displayNumber, `pages[${pageIndex}].pageNumber.displayNumber`);
      if (!Number.isInteger(page.pageNumber.displayNumber)) {
        throw new LayoutInvariantError(
          'INVALID_GEOMETRY',
          `pages[${pageIndex}] page number is not an integer`,
        );
      }
      if (
        page.pageNumber.format.length === 0
        || !sectionOccurrenceIds.has(page.pageNumber.sectionOccurrenceId)
      ) {
        throw new LayoutInvariantError(
          'INVALID_REFERENCE',
          `pages[${pageIndex}] has an invalid page number section owner`,
        );
      }
    }

    const ordinary: PaintNode[] = [];
    try {
      orderedPagePaintNodes(page);
    } catch (error) {
      if (error instanceof PageGraphError) {
        throw new LayoutInvariantError('INVALID_REFERENCE', error.message);
      }
      throw error;
    }
    const nodes = new Map<string, PaintNode>();
    const retainedNodeIds = new Set<string>();
    pageLayerNodes(page).forEach(({ node }, nodeIndex) => {
      const path = `pages[${pageIndex}].nodes[${nodeIndex}]`;
      nodes.set(node.id, node);
      collectRetainedNodeIds(node, retainedNodeIds, documentRetainedNodeIds);
      requireRect(node.flowBounds, `${path}.flowBounds`);
      requireRect(node.inkBounds, `${path}.inkBounds`);
      if (node.clipBounds) requireRect(node.clipBounds, `${path}.clipBounds`);
      requireFinite(node.advancePt, `${path}.advancePt`);
      if (node.kind === 'drawing') requireDrawingGeometry(node, path);
      const domain = domains.get(node.flowDomainId);
      if (!domain) {
        throw new LayoutInvariantError('INVALID_REFERENCE', `${node.id} references missing flow domain ${node.flowDomainId}`);
      }
      if (!node.ordinaryFlow) return;
      if (domain.kind === 'body') {
        const region = regionByDomain.get(domain.id);
        const blockEndPt = region?.blockEndPt ?? page.geometry.contentBottomPt;
        if (node.flowBounds.yPt + node.flowBounds.heightPt > blockEndPt) {
          throw new LayoutInvariantError('BOTTOM_MARGIN_INVASION', `${node.id} crosses logical block end`);
        }
      }
      if (!contains(domain.logicalBounds, node.flowBounds)) {
        throw new LayoutInvariantError('FLOW_DOMAIN_INVASION', `${node.id} crosses flow domain ${domain.id}`);
      }
      ordinary.push(node);
    });

    const read = new Set<string>();
    page.readingOrder.forEach((nodeId) => {
      if (!nodes.has(nodeId) || read.has(nodeId)) {
        throw new LayoutInvariantError('INVALID_REFERENCE', `invalid reading-order reference ${nodeId}`);
      }
      read.add(nodeId);
    });

    if (page.bookmarkStarts !== undefined) {
      const sectionByDomain = new Map(
        [...regionByDomain].map(([domainId, region]) => [
          domainId,
          region.sectionOccurrenceId,
        ]),
      );
      const expectedBookmarkStarts = derivePageBookmarkStarts(
        orderedPagePaintNodes(page),
        page.sectionOccurrenceId ?? '',
        sectionByDomain,
      );
      const bookmarkOwnersAreValid = expectedBookmarkStarts.every((bookmark) =>
        bookmark.sectionOccurrenceId.length > 0
        && sectionOccurrenceIds.has(bookmark.sectionOccurrenceId));
      const bookmarkMetadataMatches = page.bookmarkStarts.length === expectedBookmarkStarts.length
        && page.bookmarkStarts.every((bookmark, index) => {
          const expected = expectedBookmarkStarts[index];
          return expected !== undefined
            && bookmark.name === expected.name
            && bookmark.nodeId === expected.nodeId
            && bookmark.sectionOccurrenceId === expected.sectionOccurrenceId;
        });
      if (!bookmarkOwnersAreValid || !bookmarkMetadataMatches) {
        throw new LayoutInvariantError(
          'INVALID_REFERENCE',
          `pages[${pageIndex}] bookmark metadata does not match its retained graph (invalid bookmark node or ownership)`,
        );
      }
    }

    for (let index = 0; index < ordinary.length; index += 1) {
      for (let other = index + 1; other < ordinary.length; other += 1) {
        const first = ordinary[index];
        const second = ordinary[other];
        if (first && second
          && first.flowDomainId === second.flowDomainId
          && overlaps(first.flowBounds, second.flowBounds)) {
          throw new LayoutInvariantError('FLOW_OVERLAP', `${first.id} overlaps ${second.id}`);
        }
      }
    }
  });
}

export function assertDocumentLayout(layout: DocumentLayout): void {
  try {
    assertDocumentLayoutUnchecked(layout);
  } catch (error) {
    if (error instanceof LayoutInvariantError) throw error;
    if (error instanceof TypeError || error instanceof RangeError) {
      throw new LayoutInvariantError('INVALID_GEOMETRY', error.message);
    }
    throw error;
  }
}

function canonicalize(value: unknown): unknown {
  if (typeof value === 'number') {
    if (!Number.isFinite(value)) throw new LayoutInvariantError('INVALID_GEOMETRY', 'fingerprint input is not finite');
    const normalized = Number(value.toFixed(6));
    return Object.is(normalized, -0) ? 0 : normalized;
  }
  if (value === null || typeof value === 'string' || typeof value === 'boolean') return value;
  if (Array.isArray(value)) return value.map((entry) => canonicalize(entry));
  if (typeof value === 'object') {
    return Object.fromEntries(Object.entries(value as Record<string, unknown>)
      .sort(([left], [right]) => left.localeCompare(right))
      .map(([entryKey, entry]) => [entryKey, canonicalize(entry)]));
  }
  throw new LayoutInvariantError('INVALID_GEOMETRY', `fingerprint contains ${typeof value}`);
}

export function layoutFingerprint(layout: DocumentLayout): string {
  assertPlainData(layout, 'layout');
  const value = {
    pages: layout.pages,
    diagnostics: layout.diagnostics.map(({ message: _message, ...identity }) => identity),
  };
  return JSON.stringify(canonicalize(value));
}

function deepFreeze<T>(value: T, seen: WeakSet<object>): DeepReadonly<T> {
  if (value === null || typeof value !== 'object') return value as DeepReadonly<T>;
  if (seen.has(value)) return value as DeepReadonly<T>;
  seen.add(value);
  for (const child of Object.values(value as Record<string, unknown>)) deepFreeze(child, seen);
  return Object.freeze(value) as DeepReadonly<T>;
}

export function deepFreezeDocumentLayout(layout: DocumentLayout): DeepReadonly<DocumentLayout> {
  assertPlainData(layout, 'layout');
  return deepFreeze(layout, new WeakSet<object>());
}
