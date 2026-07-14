import type { DeepReadonly, LayoutDiagnostic, SourceRef } from './types.js';
import { stableFingerprint } from './fingerprint.js';
import type { BodyElement, DocRun, DocTable, DocxDocumentModel, ShapeRun } from '../types.js';
import { rasterExceedsBudget, sniffRasterDimensions } from '@silurus/ooxml-core';

export interface ImageLayoutResource {
  readonly widthPt: number;
  readonly heightPt: number;
  readonly mimeType: string;
}

export interface ImageMetadataRecord extends ImageLayoutResource {
  readonly resourceKey: string;
}

export interface MathLayoutResource {
  readonly resourceKey: string;
  /** Deterministic content/display alias used by the transitional line adapter. */
  readonly lookupKey?: string;
  readonly widthEm: number;
  readonly ascentEm: number;
  readonly descentEm: number;
  readonly diagnostics: readonly LayoutDiagnostic[];
  readonly available?: boolean;
}

export interface ImageMetadataService {
  readonly fingerprint: string;
  resolve(resourceKey: string): Readonly<ImageLayoutResource>;
}

export interface MathMetadataService {
  readonly fingerprint: string;
  resolve(resourceKey: string): DeepReadonly<MathLayoutResource>;
}

function sourceKey(source: SourceRef): string {
  return `${source.story}:${encodeURIComponent(source.storyInstance)}:${source.path.join('.')}`;
}

export function imageResourceKey(source: SourceRef, partPath: string): string {
  return `image:${sourceKey(source)}:${encodeURIComponent(partPath)}`;
}

export function mathResourceKey(source: SourceRef, localName: string): string {
  return `math:${sourceKey(source)}:${encodeURIComponent(localName)}`;
}

export function mathAstResourceKey(nodes: unknown): string {
  return stableFingerprint('math-resource', nodes);
}

function finiteNonNegative(value: number, name: string): number {
  if (!Number.isFinite(value) || value < 0) throw new RangeError(`${name} must be finite and non-negative`);
  return value;
}

export function createImageMetadataService(records: readonly ImageMetadataRecord[]): ImageMetadataService {
  const snapshot = [...records]
    .map((record) => Object.freeze({
      resourceKey: record.resourceKey,
      widthPt: finiteNonNegative(record.widthPt, 'widthPt'),
      heightPt: finiteNonNegative(record.heightPt, 'heightPt'),
      mimeType: record.mimeType,
    }))
    .sort((a, b) => a.resourceKey.localeCompare(b.resourceKey));
  const byKey = new Map(snapshot.map(({ resourceKey, ...metadata }) => [resourceKey, Object.freeze(metadata)]));
  if (byKey.size !== snapshot.length) throw new Error('Duplicate image resource key');
  return Object.freeze({
    fingerprint: stableFingerprint('images', snapshot),
    resolve(resourceKey: string): Readonly<ImageLayoutResource> {
      const resource = byKey.get(resourceKey);
      if (!resource) throw new Error(`Unknown image resource: ${resourceKey}`);
      return resource;
    },
  });
}

export function rasterImageMetadataRecord(
  resourceKey: string,
  bytes: Uint8Array,
  mimeType: string,
  dpi: number,
): ImageMetadataRecord {
  const dimensions = sniffRasterDimensions(bytes);
  if (!dimensions || rasterExceedsBudget(dimensions)) {
    throw new Error(`Raster dimensions are unavailable or unsafe for ${resourceKey}`);
  }
  if (!Number.isFinite(dpi) || dpi <= 0) throw new RangeError('dpi must be positive');
  return {
    resourceKey,
    widthPt: dimensions.width * 72 / dpi,
    heightPt: dimensions.height * 72 / dpi,
    mimeType,
  };
}

export function createMathMetadataService(records: readonly MathLayoutResource[]): MathMetadataService {
  const snapshot = [...records]
    .map((record) => Object.freeze({
      resourceKey: record.resourceKey,
      ...(record.lookupKey ? { lookupKey: record.lookupKey } : {}),
      widthEm: finiteNonNegative(record.widthEm, 'widthEm'),
      ascentEm: finiteNonNegative(record.ascentEm, 'ascentEm'),
      descentEm: finiteNonNegative(record.descentEm, 'descentEm'),
      diagnostics: Object.freeze(record.diagnostics.map((diagnostic) => Object.freeze({ ...diagnostic }))),
      ...(record.available === false ? { available: false } : {}),
    }))
    .sort((a, b) => a.resourceKey.localeCompare(b.resourceKey));
  const byKey = new Map(snapshot.map((resource) => [resource.resourceKey, resource]));
  if (byKey.size !== snapshot.length) throw new Error('Duplicate math resource key');
  for (const resource of snapshot) {
    if (resource.lookupKey && !byKey.has(resource.lookupKey)) byKey.set(resource.lookupKey, resource);
  }
  return Object.freeze({
    fingerprint: stableFingerprint('math', snapshot),
    resolve(resourceKey: string): DeepReadonly<MathLayoutResource> {
      const resource = byKey.get(resourceKey);
      if (!resource) throw new Error(`Unknown math resource: ${resourceKey}`);
      return resource;
    },
  });
}

export function documentImageMetadataRecords(doc: DocxDocumentModel): ImageMetadataRecord[] {
  const records: ImageMetadataRecord[] = [];
  const add = (source: SourceRef, imagePath: string, mimeType: string, widthPt: number, heightPt: number): void => {
    records.push({ resourceKey: imageResourceKey(source, imagePath), widthPt, heightPt, mimeType });
  };
  const visitRun = (run: DocRun, source: SourceRef): void => {
    if (run.type === 'image') {
      add(source, run.imagePath, run.mimeType, run.widthPt, run.heightPt);
      return;
    }
    if (run.type !== 'shape') return;
    const shape = run as { type: 'shape' } & ShapeRun;
    shape.textBlocks?.forEach((block, index) => {
      if (!block.imagePath || !block.mimeType || block.imageWidthPt == null || block.imageHeightPt == null) return;
      const textBoxSource: SourceRef = {
        story: 'textbox',
        storyInstance: `${source.story}:${source.storyInstance}:${source.path.join('.')}`,
        path: [index],
      };
      add(textBoxSource, block.imagePath, block.mimeType, block.imageWidthPt, block.imageHeightPt);
    });
  };
  const visitTable = (table: DocTable, story: SourceRef['story'], storyInstance: string, prefix: number[]): void => {
    table.rows.forEach((row, rowIndex) => row.cells.forEach((cell, cellIndex) => {
      visitBody(cell.content as BodyElement[], story, storyInstance, [...prefix, rowIndex, cellIndex]);
    }));
  };
  const visitBody = (body: BodyElement[], story: SourceRef['story'], storyInstance: string, prefix: number[] = []): void => {
    body.forEach((element, elementIndex) => {
      const path = [...prefix, elementIndex];
      if (element.type === 'paragraph') {
        element.runs.forEach((run, runIndex) => visitRun(run, { story, storyInstance, path: [...path, runIndex] }));
      } else if (element.type === 'table') {
        visitTable(element, story, storyInstance, path);
      }
      if (element.type === 'sectionBreak') {
        for (const kind of ['default', 'first', 'even'] as const) {
          const header = element.headers?.[kind];
          const footer = element.footers?.[kind];
          if (header) visitBody(header.body, 'header', `section:${elementIndex}:${kind}`);
          if (footer) visitBody(footer.body, 'footer', `section:${elementIndex}:${kind}`);
        }
      }
    });
  };
  visitBody(doc.body, 'body', 'body');
  for (const kind of ['default', 'first', 'even'] as const) {
    const header = doc.headers[kind];
    const footer = doc.footers[kind];
    if (header) visitBody(header.body, 'header', kind);
    if (footer) visitBody(footer.body, 'footer', kind);
  }
  return records;
}
