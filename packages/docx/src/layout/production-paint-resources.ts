import type {
  BodyElement,
  DocParagraph,
  DocRun,
  DocTable,
  DocxDocumentModel,
  HeadersFooters,
  ImageRun,
  ShapeRun,
} from '../types.js';
import { stableFingerprint } from './fingerprint.js';
import { documentImageMetadataRecords, documentMathOccurrences } from './resources.js';
import type { ImageMetadataRecord } from './resources.js';
import { imageResourceKey } from './source-key.js';
import {
  createPaintResourceRegistry,
} from './paint-resources.js';
import type {
  ImagePaintResourceDescriptor,
  PaintResourceDescriptor,
  PaintResourceRegistry,
} from './types.js';

export function chartPaintResourceKey(source: import('./types.js').SourceRef): string {
  return stableFingerprint('chart-resource', source);
}

type ImageDescriptorCandidate = Omit<ImagePaintResourceDescriptor, 'intrinsicSize' | 'mimeType'>;

function imageCandidate(
  kind: 'image' | 'picture-bullet',
  resourceKey: string,
  partPath: string,
  run: Partial<ImageRun> = {},
): ImageDescriptorCandidate {
  return {
    kind,
    resourceKey,
    partPath,
    ...(run.svgImagePath === undefined ? {} : { svgImagePath: run.svgImagePath }),
    ...(run.srcRect == null ? {} : { srcRect: { ...run.srcRect } }),
    ...(run.rotation === undefined ? {} : { rotation: run.rotation }),
    ...(run.flipH === undefined ? {} : { flipH: run.flipH }),
    ...(run.flipV === undefined ? {} : { flipV: run.flipV }),
    ...(run.alpha === undefined ? {} : { alpha: run.alpha }),
    ...(run.colorReplaceFrom === undefined ? {} : { colorReplaceFrom: run.colorReplaceFrom }),
    ...(run.duotone === undefined ? {} : { duotone: { ...run.duotone } }),
  };
}

function collectDescriptorCandidates(
  doc: DocxDocumentModel,
  retainedImageMetadata?: readonly ImageMetadataRecord[],
): PaintResourceDescriptor[] {
  const imageCandidates: ImageDescriptorCandidate[] = [];
  const descriptors: PaintResourceDescriptor[] = [];
  const visitRun = (run: DocRun, source: import('./types.js').SourceRef): void => {
    if (run.type === 'image') {
      imageCandidates.push(imageCandidate(
        'image', imageResourceKey(source, run.imagePath), run.imagePath, run,
      ));
      return;
    }
    if (run.type === 'chart') {
      descriptors.push({
        kind: 'chart',
        resourceKey: chartPaintResourceKey(source),
        intrinsicSize: { widthPt: run.widthPt, heightPt: run.heightPt },
        model: run.chart,
      });
      return;
    }
    if (run.type !== 'shape') return;
    const storyInstance = `${source.story}:${source.storyInstance}:${source.path.join('.')}`;
    (run as ShapeRun).textBlocks?.forEach((block, blockIndex) => {
      if (!block.imagePath) return;
      const textBoxSource = {
        story: 'textbox' as const,
        storyInstance,
        // normalizeTextBoxInput projects each compatibility block to a retained
        // paragraph and its optional image to run 0.
        path: [blockIndex, 0],
      };
      imageCandidates.push(imageCandidate(
        'image', imageResourceKey(textBoxSource, block.imagePath), block.imagePath,
        { svgImagePath: block.svgImagePath },
      ));
    });
  };
  const visitTable = (
    table: DocTable,
    story: import('./types.js').SourceRef['story'],
    storyInstance: string,
    prefix: number[],
  ): void => {
    table.rows.forEach((row, rowIndex) => row.cells.forEach((cell, cellIndex) => {
      visitBody(
        cell.content as BodyElement[],
        story,
        storyInstance,
        [...prefix, rowIndex, cellIndex],
      );
    }));
  };
  const visitParts = (
    parts: HeadersFooters | undefined,
    story: 'header' | 'footer',
    instancePrefix?: string,
  ): void => {
    if (!parts) return;
    for (const kind of ['default', 'first', 'even'] as const) {
      const part = parts[kind];
      if (part) visitBody(part.body, story, instancePrefix ? `${instancePrefix}:${kind}` : kind);
    }
  };
  const visitParagraph = (
    paragraph: DocParagraph,
    source: import('./types.js').SourceRef,
  ): void => {
    const numbering = paragraph.numbering;
    if (numbering?.picBulletImagePath) {
      imageCandidates.push(imageCandidate(
        'picture-bullet',
        imageResourceKey(source, numbering.picBulletImagePath),
        numbering.picBulletImagePath,
      ));
    }
    paragraph.runs.forEach((run, runIndex) => {
      visitRun(run, { ...source, path: [...source.path, runIndex] });
    });
  };
  const visitBody = (
    body: BodyElement[],
    story: import('./types.js').SourceRef['story'],
    storyInstance: string,
    prefix: number[] = [],
  ): void => {
    body.forEach((element, elementIndex) => {
      const path = [...prefix, elementIndex];
      if (element.type === 'paragraph') {
        visitParagraph(element, { story, storyInstance, path });
      } else if (element.type === 'table') {
        visitTable(element, story, storyInstance, path);
      } else if (element.type === 'sectionBreak') {
        visitParts(element.headers, 'header', `section:${elementIndex}`);
        visitParts(element.footers, 'footer', `section:${elementIndex}`);
      }
    });
  };

  visitBody(doc.body, 'body', 'body');
  visitParts(doc.headers, 'header');
  visitParts(doc.footers, 'footer');
  for (const note of doc.footnotes ?? []) visitBody(note.content, 'footnote', note.id);
  for (const note of doc.endnotes ?? []) visitBody(note.content, 'endnote', note.id);

  const metadata = retainedImageMetadata ?? documentImageMetadataRecords(doc);
  const metadataByKey = new Map(metadata.map((record) => [record.resourceKey, record]));
  const candidateKeys = imageCandidates.map((candidate) => candidate.resourceKey).sort();
  const metadataKeys = metadata.map((record) => record.resourceKey).sort();
  if (candidateKeys.length !== metadataKeys.length
    || candidateKeys.some((key, index) => key !== metadataKeys[index])) {
    throw new Error('Paint image descriptor membership differs from layout image metadata');
  }
  for (const candidate of imageCandidates) {
    const record = metadataByKey.get(candidate.resourceKey);
    if (!record) throw new Error(`Missing layout image metadata: ${candidate.resourceKey}`);
    descriptors.push({
      ...candidate,
      mimeType: record.mimeType,
      intrinsicSize: { widthPt: record.widthPt, heightPt: record.heightPt },
    });
  }
  for (const occurrence of documentMathOccurrences(doc)) {
    descriptors.push({ kind: 'math', resourceKey: occurrence.resourceKey });
  }
  return descriptors;
}

export function createDocumentPaintResourceRegistry(
  doc: DocxDocumentModel,
  retainedImageMetadata?: readonly ImageMetadataRecord[],
): PaintResourceRegistry {
  return createPaintResourceRegistry(collectDescriptorCandidates(doc, retainedImageMetadata));
}
