import type { DeepReadonly, LayoutDiagnostic, SourceRef } from './types.js';
import { stableFingerprint } from './fingerprint.js';
import type { DocxDocumentModel } from '../types.js';
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
      widthEm: finiteNonNegative(record.widthEm, 'widthEm'),
      ascentEm: finiteNonNegative(record.ascentEm, 'ascentEm'),
      descentEm: finiteNonNegative(record.descentEm, 'descentEm'),
      diagnostics: Object.freeze([...record.diagnostics]),
      ...(record.available === false ? { available: false } : {}),
    }))
    .sort((a, b) => a.resourceKey.localeCompare(b.resourceKey));
  const byKey = new Map(snapshot.map((resource) => [resource.resourceKey, resource]));
  if (byKey.size !== snapshot.length) throw new Error('Duplicate math resource key');
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
  const visit = (value: unknown, path: readonly (string | number)[]): void => {
    if (Array.isArray(value)) {
      value.forEach((child, index) => visit(child, [...path, index]));
      return;
    }
    if (!value || typeof value !== 'object') return;
    const record = value as Record<string, unknown>;
    if (
      typeof record.imagePath === 'string'
      && typeof record.widthPt === 'number'
      && typeof record.heightPt === 'number'
      && typeof record.mimeType === 'string'
    ) {
      records.push({
        resourceKey: `image:document:${path.map((part) => encodeURIComponent(String(part))).join('/')}:${encodeURIComponent(record.imagePath)}`,
        widthPt: record.widthPt,
        heightPt: record.heightPt,
        mimeType: record.mimeType,
      });
    }
    for (const key of Object.keys(record).sort()) visit(record[key], [...path, key]);
  };
  visit(doc, []);
  return records;
}
