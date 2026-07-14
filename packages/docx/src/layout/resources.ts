import type { DeepReadonly, LayoutDiagnostic, SourceRef } from './types.js';
import { stableFingerprint } from './fingerprint.js';
import type { BodyElement, DocRun, DocTable, DocxDocumentModel, ShapeRun } from '../types.js';
import type { MathNode } from '@silurus/ooxml-core';
import { rasterExceedsBudget, sniffRasterDimensions } from '@silurus/ooxml-core';
import { imageResourceKey, mathResourceKey } from './source-key.js';
import { normalizeInternalDocumentModel } from '../parser-model.js';

export { imageResourceKey, mathResourceKey } from './source-key.js';

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

export interface MathOccurrence {
  readonly nodes: MathNode[];
  readonly display: boolean;
  readonly source: SourceRef;
  readonly resourceKey: string;
}

export interface ImageMetadataService {
  readonly fingerprint: string;
  resolve(resourceKey: string): Readonly<ImageLayoutResource>;
}

export interface MathMetadataService {
  readonly fingerprint: string;
  resolve(resourceKey: string): DeepReadonly<MathLayoutResource>;
}

export function bodyMathOccurrences(
  body: BodyElement[],
  story: SourceRef['story'] = 'body',
  storyInstance = 'body',
): MathOccurrence[] {
  const found: MathOccurrence[] = [];
  const visit = (elements: BodyElement[], prefix: number[] = []): void => {
    elements.forEach((element, elementIndex) => {
      const path = [...prefix, elementIndex];
      if (element.type === 'paragraph') {
        element.runs.forEach((run, runIndex) => {
          if (run.type === 'math') found.push({
            nodes: run.nodes,
            display: run.display,
            source: { story, storyInstance, path: [...path, runIndex] },
            resourceKey: mathResourceKey(
              { story, storyInstance, path: [...path, runIndex] },
              run.display ? 'display' : 'inline',
            ),
          });
        });
      } else if (element.type === 'table') {
        element.rows.forEach((row, rowIndex) => row.cells.forEach((cell, cellIndex) => {
          visit(cell.content as BodyElement[], [...path, rowIndex, cellIndex]);
        }));
      }
    });
  };
  visit(body);
  return found;
}

/** Collect every story whose current model retains OMML MathNode runs. Shape
 * txbxContent is intentionally excluded: the parser currently flattens its
 * equations into ShapeTextRun.text. Planned rich-textbox work must extend this
 * traversal when that model begins retaining MathNode rather than silently
 * resolving an absent resource. */
export function documentMathOccurrences(doc: DocxDocumentModel): MathOccurrence[] {
  return [...normalizeInternalDocumentModel(doc).mathOccurrences];
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
      diagnostics: Object.freeze(record.diagnostics.map((diagnostic) => Object.freeze({ ...diagnostic }))),
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
        const numbering = element.numbering;
        if (numbering?.picBulletImagePath && numbering.picBulletMimeType
          && numbering.picBulletWidthPt != null && numbering.picBulletHeightPt != null) {
          add(
            { story, storyInstance, path },
            numbering.picBulletImagePath,
            numbering.picBulletMimeType,
            numbering.picBulletWidthPt,
            numbering.picBulletHeightPt,
          );
        }
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
  for (const note of doc.footnotes ?? []) visitBody(note.content, 'footnote', note.id);
  for (const note of doc.endnotes ?? []) visitBody(note.content, 'endnote', note.id);
  return records;
}
