import type { SectionLayoutContext } from '../layout-context.js';
import {
  createSectionRegionCoordinateSpace,
  logicalPageExtent,
  transformRect,
  uprightPhysicalExtent,
  writingModeFromTextDirection,
  type PhysicalPageExtent,
} from './coordinate-space.js';
import { PAGE_LAYER_IDS, type PageLayerNode } from './page-graph.js';
import type { BodyOccurrenceDestination } from './occurrence-projection.js';
import type {
  DeepReadonly,
  FlowDomain,
  LayoutPage,
  LayoutRect,
  PageBookmarkStart,
  PageLayers,
  PageNumberMetadata,
  PageSectionRegion,
  PaintNode,
  ParagraphLayout,
  TableLayout,
  TextBoxLayout,
  WritingMode,
} from './types.js';

export interface PhysicalPageInput extends PhysicalPageExtent {
  /** Effective main-story edges after §17.6.11 header/footer interaction. */
  readonly contentTopPt: number;
  readonly contentBottomPt: number;
}

export interface LogicalColumnInput {
  readonly inlineStartPt: number;
  readonly inlineExtentPt: number;
}

export interface PageSectionRegionInput {
  readonly id: string;
  readonly sectionOccurrenceId: string;
  readonly section: DeepReadonly<SectionLayoutContext>;
  readonly writingMode: WritingMode;
  readonly blockStartPt: number;
  readonly blockEndPt: number;
  readonly columns: readonly LogicalColumnInput[];
}

export interface LayoutPageAccumulatorInput {
  readonly pageIndex: number;
  readonly physicalPage: PhysicalPageInput;
  readonly sectionOccurrenceId: string;
  readonly section: DeepReadonly<SectionLayoutContext>;
}

export interface LayoutPageAccumulator extends LayoutPageAccumulatorInput {
  readonly sectionRegions: readonly PageSectionRegionInput[];
  readonly paint: readonly PageLayerNode[];
  readonly readingOrder: readonly PaintNode[];
}

export interface LayoutPageFactoryInput extends LayoutPageAccumulatorInput {
  readonly sectionRegions: readonly PageSectionRegionInput[];
  readonly paint: readonly PageLayerNode[];
  readonly readingOrder: readonly PaintNode[];
  readonly pageNumber: PageNumberMetadata;
}

export interface ParityBlankLayoutPageInput {
  readonly pageIndex: number;
  readonly physicalPage: PhysicalPageInput;
  readonly sectionOccurrenceId: string;
  readonly section: DeepReadonly<SectionLayoutContext>;
  readonly pageNumber: PageNumberMetadata;
}

export function bodyFlowDomainId(
  pageIndex: number,
  regionId: string,
  columnIndex: number,
): string {
  return `page:${pageIndex}:region:${encodeURIComponent(regionId)}:column:${columnIndex}`;
}

function pageGeometry(page: PhysicalPageInput): LayoutPage['geometry'] {
  requireEffectivePageEdges(page);
  return {
    xPt: 0,
    yPt: 0,
    widthPt: page.widthPt,
    heightPt: page.heightPt,
    // §17.6.11 makes positive top/bottom edges depend on header/footer extent,
    // while negative values ignore that extent. The page owner resolves those
    // facts; this factory only retains the resulting effective coordinates.
    contentTopPt: page.contentTopPt,
    contentBottomPt: page.contentBottomPt,
  };
}

function requireEffectivePageEdges(page: PhysicalPageInput): void {
  if (
    !Number.isFinite(page.widthPt)
    || !Number.isFinite(page.heightPt)
    || !Number.isFinite(page.contentTopPt)
    || !Number.isFinite(page.contentBottomPt)
    || page.widthPt <= 0
    || page.heightPt <= 0
    || page.contentTopPt < 0
    || page.contentTopPt > page.contentBottomPt
    || page.contentBottomPt > page.heightPt
  ) {
    throw new RangeError(
      'Effective page edges must satisfy 0 <= contentTopPt <= contentBottomPt <= heightPt',
    );
  }
  // Equal edges are valid and represent an empty main-story interval.
}

function requireIdentity(value: string, name: string): void {
  if (value.length === 0) throw new RangeError(`${name} must not be empty`);
}

function equalColumns(
  left: SectionLayoutContext['columns'],
  right: SectionLayoutContext['columns'],
): boolean {
  return left.length === right.length && left.every((column, index) => {
    const other = right[index];
    return other !== undefined && column.xPt === other.xPt && column.wPt === other.wPt;
  });
}

function equalLineNumbering(
  left: SectionLayoutContext['lineNumbering'],
  right: SectionLayoutContext['lineNumbering'],
): boolean {
  return left === right || (left !== undefined && right !== undefined
    && left.start === right.start
    && left.countBy === right.countBy
    && left.distance === right.distance
    && left.restart === right.restart);
}

export function sectionLayoutContextsEqual(
  left: DeepReadonly<SectionLayoutContext>,
  right: DeepReadonly<SectionLayoutContext>,
): boolean {
  return left.geometry.pageWidth === right.geometry.pageWidth
    && left.geometry.pageHeight === right.geometry.pageHeight
    && left.geometry.marginTop === right.geometry.marginTop
    && left.geometry.marginRight === right.geometry.marginRight
    && left.geometry.marginBottom === right.geometry.marginBottom
    && left.geometry.marginLeft === right.geometry.marginLeft
    && left.geometry.headerDistance === right.geometry.headerDistance
    && left.geometry.footerDistance === right.geometry.footerDistance
    && equalColumns(left.columns, right.columns)
    && left.textDirection === right.textDirection
    && left.grid.kind === right.grid.kind
    && left.grid.linePitchPt === right.grid.linePitchPt
    && left.grid.charSpacePt === right.grid.charSpacePt
    && left.verticalAlignment === right.verticalAlignment
    && equalLineNumbering(left.lineNumbering, right.lineNumbering);
}

function requireRegionSectionAgreement(input: PageSectionRegionInput): void {
  const writingMode = writingModeFromTextDirection(input.section.textDirection);
  if (writingMode !== input.writingMode) {
    throw new RangeError('Section region writing mode must agree with its section text direction');
  }
  if (input.columns.length !== input.section.columns.length
    || input.columns.some((column, index) => {
      const sectionColumn = input.section.columns[index];
      return sectionColumn === undefined
        || column.inlineStartPt !== sectionColumn.xPt
        || column.inlineExtentPt !== sectionColumn.wPt;
    })) {
    throw new RangeError('Section region columns must equal its normalized section columns');
  }
}

function requireRect(rect: LayoutRect, name: string): void {
  if (!Number.isFinite(rect.xPt) || !Number.isFinite(rect.yPt)
    || !Number.isFinite(rect.widthPt) || !Number.isFinite(rect.heightPt)
    || rect.widthPt < 0 || rect.heightPt < 0) {
    throw new RangeError(`${name} must be a finite rectangle with non-negative extents`);
  }
}

function requirePageIndex(pageIndex: number): void {
  if (!Number.isInteger(pageIndex) || pageIndex < 0) {
    throw new RangeError('Layout page index must be a non-negative integer');
  }
}

function buildRegions(
  pageIndex: number,
  physicalPage: PhysicalPageInput,
  inputs: readonly PageSectionRegionInput[],
): Readonly<{
  regions: readonly PageSectionRegion[];
  domains: readonly FlowDomain[];
  sectionByDomain: ReadonlyMap<string, string>;
}> {
  const regions: PageSectionRegion[] = [];
  const domains: FlowDomain[] = [];
  const sectionByDomain = new Map<string, string>();
  const regionIds = new Set<string>();
  const occurrenceIds = new Set<string>();
  let priorBlockEndPt = 0;
  let pageWritingMode: WritingMode | undefined;

  for (const input of inputs) {
    requireIdentity(input.id, 'Section region id');
    requireIdentity(input.sectionOccurrenceId, 'Section occurrence id');
    if (regionIds.has(input.id) || occurrenceIds.has(input.sectionOccurrenceId)) {
      throw new RangeError('Section region and occurrence identities must be unique');
    }
    regionIds.add(input.id);
    occurrenceIds.add(input.sectionOccurrenceId);
    if (pageWritingMode !== undefined && pageWritingMode !== input.writingMode) {
      throw new RangeError('One physical page cannot mix writing modes');
    }
    pageWritingMode = input.writingMode;
    requireRegionSectionAgreement(input);
    const expectedPhysicalExtent = uprightPhysicalExtent({
      widthPt: input.section.geometry.pageWidth,
      heightPt: input.section.geometry.pageHeight,
    }, input.writingMode);
    if (expectedPhysicalExtent.widthPt !== physicalPage.widthPt
      || expectedPhysicalExtent.heightPt !== physicalPage.heightPt) {
      throw new RangeError('Section regions on one physical page must use the same page box');
    }
    const logicalExtent = logicalPageExtent(physicalPage, input.writingMode);
    const logicalInlineExtent = logicalExtent.widthPt;
    const logicalBlockExtent = logicalExtent.heightPt;
    if (!Number.isFinite(input.blockStartPt) || !Number.isFinite(input.blockEndPt)
      || input.blockStartPt < 0 || input.blockEndPt <= input.blockStartPt
      || input.blockEndPt > logicalBlockExtent || input.blockStartPt < priorBlockEndPt) {
      throw new RangeError('Section regions must be ordered, disjoint, and inside the logical page');
    }
    priorBlockEndPt = input.blockEndPt;
    if (input.columns.length === 0) throw new RangeError('Section region must contain a column');
    let priorInlineEndPt = 0;
    const coordinateSpace = createSectionRegionCoordinateSpace(input.writingMode, physicalPage);
    const flowDomainIds = input.columns.map((column, columnIndex) => {
      if (!Number.isFinite(column.inlineStartPt) || !Number.isFinite(column.inlineExtentPt)
        || column.inlineStartPt < 0 || column.inlineExtentPt <= 0
        || column.inlineStartPt + column.inlineExtentPt > logicalInlineExtent
        || column.inlineStartPt < priorInlineEndPt) {
        throw new RangeError('Columns must be ordered, disjoint, and inside the logical page');
      }
      priorInlineEndPt = column.inlineStartPt + column.inlineExtentPt;
      const id = bodyFlowDomainId(pageIndex, input.id, columnIndex);
      if (sectionByDomain.has(id)) throw new RangeError(`Duplicate flow domain ${id}`);
      const logicalBounds = {
        xPt: column.inlineStartPt,
        yPt: input.blockStartPt,
        widthPt: column.inlineExtentPt,
        heightPt: input.blockEndPt - input.blockStartPt,
      };
      domains.push({
        id,
        kind: 'body',
        logicalBounds,
        physicalBounds: transformRect(coordinateSpace.logicalToPhysical, logicalBounds),
      });
      sectionByDomain.set(id, input.sectionOccurrenceId);
      return id;
    });
    regions.push({
      id: input.id,
      sectionOccurrenceId: input.sectionOccurrenceId,
      coordinateSpace,
      blockStartPt: input.blockStartPt,
      blockEndPt: input.blockEndPt,
      flowDomainIds,
      section: input.section,
    });
  }

  return { regions, domains, sectionByDomain };
}

export function bodyOccurrenceDestinationFor(
  pageIndex: number,
  region: PageSectionRegionInput,
  columnIndex: number,
  blockStartPt: number,
  retainedFlowBounds: LayoutRect,
): BodyOccurrenceDestination {
  requirePageIndex(pageIndex);
  requireIdentity(region.id, 'Section region id');
  if (!Number.isInteger(columnIndex) || columnIndex < 0 || columnIndex >= region.columns.length) {
    throw new RangeError('Column index must identify a section region column');
  }
  if (!Number.isFinite(blockStartPt)) throw new RangeError('Block start must be finite');
  requireRect(retainedFlowBounds, 'Retained flow bounds');
  const column = region.columns[columnIndex]!;
  if (!Number.isFinite(column.inlineStartPt)) throw new RangeError('Column inline start must be finite');
  return {
    coordinateSpace: 'logical-page-points',
    flowDomainId: bodyFlowDomainId(pageIndex, region.id, columnIndex),
    translation: {
      xPt: column.inlineStartPt - retainedFlowBounds.xPt,
      yPt: blockStartPt - retainedFlowBounds.yPt,
    },
  };
}

function buildLayers(entries: readonly PageLayerNode[]): PageLayers {
  const nodes = new Map(PAGE_LAYER_IDS.map((layer) => [layer, [] as PaintNode[]]));
  for (const entry of entries) nodes.get(entry.layer)!.push(entry.node);
  return {
    paintOrder: entries.map(({ layer, node }) => ({ layer, nodeId: node.id })),
    background: nodes.get('background')!,
    behindText: nodes.get('behindText')!,
    header: nodes.get('header')!,
    body: nodes.get('body')!,
    notes: nodes.get('notes')!,
    front: nodes.get('front')!,
    footer: nodes.get('footer')!,
  };
}

function visitBookmarkParagraphs(
  node: PaintNode,
  visit: (paragraph: ParagraphLayout) => void,
): void {
  if (node.kind === 'paragraph') {
    visit(node);
    node.drawings.forEach((drawing) => visitBookmarkParagraphs(drawing, visit));
    node.textBoxes.forEach((textBox) => visitBookmarkParagraphs(textBox, visit));
    return;
  }
  if (node.kind === 'table') {
    visitTableBookmarks(node, visit);
    return;
  }
  if (node.kind === 'textbox') {
    visitTextBoxBookmarks(node, visit);
  }
}

function visitTableBookmarks(
  table: TableLayout,
  visit: (paragraph: ParagraphLayout) => void,
): void {
  for (const row of table.rows) {
    for (const cell of row.cells) {
      for (const block of cell.blocks) visitBookmarkParagraphs(block.layout, visit);
    }
  }
}

function visitTextBoxBookmarks(
  textBox: TextBoxLayout,
  visit: (paragraph: ParagraphLayout) => void,
): void {
  textBox.paragraphs.forEach((paragraph) => visitBookmarkParagraphs(paragraph, visit));
}

export function derivePageBookmarkStarts(
  paint: readonly PaintNode[],
  defaultSectionOccurrenceId: string,
  sectionByDomain: ReadonlyMap<string, string>,
): readonly PageBookmarkStart[] {
  const starts: PageBookmarkStart[] = [];
  const seen = new Set<string>();
  for (const node of paint) {
    const sectionOccurrenceId = sectionByDomain.get(node.flowDomainId)
      ?? defaultSectionOccurrenceId;
    visitBookmarkParagraphs(node, (paragraph) => {
      for (const line of paragraph.lines) {
        for (const placement of line.placements) {
          if (placement.kind !== 'text' || !placement.bookmark || seen.has(placement.bookmark)) {
            continue;
          }
          seen.add(placement.bookmark);
          starts.push({
            name: placement.bookmark,
            nodeId: paragraph.id,
            sectionOccurrenceId,
          });
        }
      }
    });
  }
  return starts;
}

export function createLayoutPage(input: LayoutPageFactoryInput): LayoutPage {
  requirePageIndex(input.pageIndex);
  requireIdentity(input.sectionOccurrenceId, 'Page-start section occurrence id');
  const { regions, domains, sectionByDomain } = buildRegions(
    input.pageIndex,
    input.physicalPage,
    input.sectionRegions,
  );
  const firstRegion = input.sectionRegions[0];
  if (firstRegion !== undefined && (
    input.sectionOccurrenceId !== firstRegion.sectionOccurrenceId
    || !sectionLayoutContextsEqual(input.section, firstRegion.section)
  )) {
    throw new RangeError('Page-start section context must equal the first section region');
  }
  return {
    pageIndex: input.pageIndex,
    geometry: pageGeometry(input.physicalPage),
    flowDomains: domains,
    section: input.section,
    sectionOccurrenceId: input.sectionOccurrenceId,
    parityBlank: false,
    bookmarkStarts: derivePageBookmarkStarts(
      input.paint.map(({ node }) => node),
      input.sectionOccurrenceId,
      sectionByDomain,
    ),
    pageNumber: input.pageNumber,
    sectionRegions: regions,
    layers: buildLayers(input.paint),
    readingOrder: input.readingOrder.map((node) => node.id),
  };
}

export function createLayoutPageAccumulator(
  input: LayoutPageAccumulatorInput,
): LayoutPageAccumulator {
  requirePageIndex(input.pageIndex);
  requireIdentity(input.sectionOccurrenceId, 'Page-start section occurrence id');
  requireEffectivePageEdges(input.physicalPage);
  return Object.freeze({
    ...input,
    sectionRegions: Object.freeze([]),
    paint: Object.freeze([]),
    readingOrder: Object.freeze([]),
  });
}

export function accumulatePageSectionRegion(
  accumulator: LayoutPageAccumulator,
  sectionRegion: PageSectionRegionInput,
): LayoutPageAccumulator {
  return Object.freeze({
    ...accumulator,
    sectionRegions: Object.freeze([...accumulator.sectionRegions, sectionRegion]),
  });
}

export function accumulatePagePaintNode(
  accumulator: LayoutPageAccumulator,
  entry: PageLayerNode,
  inReadingOrder: boolean,
): LayoutPageAccumulator {
  return Object.freeze({
    ...accumulator,
    paint: Object.freeze([...accumulator.paint, entry]),
    readingOrder: inReadingOrder
      ? Object.freeze([...accumulator.readingOrder, entry.node])
      : accumulator.readingOrder,
  });
}

export function finalizeLayoutPage(
  accumulator: LayoutPageAccumulator,
  pageNumber: PageNumberMetadata,
): LayoutPage {
  return createLayoutPage({ ...accumulator, pageNumber });
}

export function createParityBlankLayoutPage(
  input: ParityBlankLayoutPageInput,
): LayoutPage {
  requirePageIndex(input.pageIndex);
  requireIdentity(input.sectionOccurrenceId, 'Page-start section occurrence id');
  return {
    pageIndex: input.pageIndex,
    geometry: pageGeometry(input.physicalPage),
    flowDomains: [],
    section: input.section,
    sectionOccurrenceId: input.sectionOccurrenceId,
    parityBlank: true,
    bookmarkStarts: [],
    pageNumber: input.pageNumber,
    sectionRegions: [],
    layers: buildLayers([]),
    readingOrder: [],
  };
}
