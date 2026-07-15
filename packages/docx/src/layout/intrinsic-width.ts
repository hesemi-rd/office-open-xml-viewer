import { graphemeClusterOffsets } from '@silurus/ooxml-core';
import type { DocParagraph, DocTable, DocTableCell } from '../types.js';
import type { ParagraphLayoutContext } from '../layout-context.js';
import {
  buildFont,
  buildSegments,
  hasCJKBreakOpportunity,
  layoutLines,
  segAdvanceWidth,
  splitTextForLayout,
  type LayoutLine,
  type LayoutSeg,
  type LayoutTextSeg,
} from '../line-layout.js';
import type {
  ParagraphMeasurementEnvironment,
  TextMeasurer,
} from '../paragraph-measure.js';
import { calcEffectiveFontPx } from './text.js';
import type { TextFontSlots } from './text.js';
import { stableFingerprint } from './fingerprint.js';
import {
  numberingMarkerLogicalInterval,
  type NumberingMarkerGeometry,
} from './numbering-marker.js';

export interface ParagraphIntrinsicWidths {
  readonly minWidthPt: number;
  readonly maxWidthPt: number;
}

export interface TableCellIntrinsicWidths {
  readonly minWidthPt: number;
  readonly maxWidthPt: number;
}

export interface TableCellIntrinsicWidthDependencies {
  paragraph(paragraph: DocParagraph): TableCellIntrinsicWidths;
  nestedTable(table: DocTable): TableCellIntrinsicWidths;
}

/** Fold public cell content into one intrinsic interval. OOXML width/style
 * precedence is deliberately absent: parser/model projection and the column
 * solver own those separate responsibilities. */
export function measureTableCellIntrinsicWidths(
  cell: Readonly<DocTableCell>,
  margins: Readonly<{ left: number; right: number }>,
  dependencies: TableCellIntrinsicWidthDependencies,
): TableCellIntrinsicWidths {
  let minContentWidthPt = 0;
  let maxContentWidthPt = 0;
  for (const element of cell.content) {
    const intrinsic = element.type === 'paragraph'
      ? dependencies.paragraph(element)
      : dependencies.nestedTable(element);
    minContentWidthPt = Math.max(minContentWidthPt, intrinsic.minWidthPt);
    maxContentWidthPt = Math.max(maxContentWidthPt, intrinsic.maxWidthPt);
  }
  const horizontalMarginsPt = Math.max(0, margins.left) + Math.max(0, margins.right);
  return {
    minWidthPt: minContentWidthPt + horizontalMarginsPt,
    maxWidthPt: Math.max(minContentWidthPt, maxContentWidthPt) + horizontalMarginsPt,
  };
}

function compatibleTextKey(segment: LayoutTextSeg): string {
  const request = segment.textShapeRequest;
  const slots = (value: TextFontSlots | undefined) => value
    ? [
        value.ascii ?? null,
        value.highAnsi ?? null,
        value.eastAsia ?? null,
        value.complexScript ?? null,
      ]
    : null;
  return stableFingerprint('paragraph-intrinsic-text', [
    segment.textLayoutService?.fingerprint ?? null,
    request ? [
      slots(request.fonts),
      slots(request.themeFonts),
      request.themeFontPresence ? [
        request.themeFontPresence.ascii ?? false,
        request.themeFontPresence.highAnsi ?? false,
        request.themeFontPresence.eastAsia ?? false,
        request.themeFontPresence.complexScript ?? false,
      ] : null,
      request.fontHint ?? null,
      request.fontSizePt,
      request.weight ?? null,
      request.style ?? null,
      request.complexScript ?? false,
      request.eastAsiaLanguage ?? null,
      request.eastAsiaFontCharset ?? null,
      request.genericFamily ?? null,
      request.letterSpacingPt ?? null,
      request.kerning ?? null,
    ] : null,
    segment.bold,
    segment.italic,
    calcEffectiveFontPx(segment, 1),
    segment.fontFamily,
    segment.fontRoute ?? null,
    segment.charScale ?? 1,
    segment.charSpacing ?? 0,
    segment.fitTextPerGapPx ?? null,
    segment.fitTextTrailingPadPx ?? null,
    segment.fitTextRegionIndex ?? null,
    segment.snapToCharacterGrid !== false,
    segment.tateChuYoko ?? false,
    // A tate-chu-yoko run is one authored one-em cell (§17.3.2.10). Two
    // adjacent runs with identical fonts remain two cells, so their source-run
    // boundary is semantic rather than a shaping-only seam.
    segment.tateChuYoko ? (segment.sourceRunIndex ?? null) : null,
    // Ruby belongs to its authored base run. Extending that base across a run
    // seam would change the annotation's ownership during the intrinsic probe.
    segment.ruby ? [
      segment.sourceRunIndex ?? null,
      segment.ruby.text,
      segment.ruby.fontSizePt,
      segment.ruby.hpsRaisePt ?? null,
    ] : null,
    segment.verticalRun ?? false,
  ]);
}

/** Run boundaries with identical effective metrics are not shaping boundaries.
 * Merge only for the intrinsic probe; retained source/run ownership stays intact. */
function mergeCompatibleTextSegments(segments: readonly LayoutSeg[]): LayoutSeg[] {
  const merged: LayoutSeg[] = [];
  for (const segment of segments) {
    const previous = merged.at(-1);
    if (
      previous
      && 'text' in previous
      && 'text' in segment
      && compatibleTextKey(previous) === compatibleTextKey(segment)
    ) {
      const text = previous.text + segment.text;
      merged[merged.length - 1] = {
        ...previous,
        text,
        textShapeRequest: previous.textShapeRequest
          ? { ...previous.textShapeRequest, text }
          : undefined,
      };
      continue;
    }
    merged.push({ ...segment });
  }
  return merged;
}

function measureTextRange(
  pieces: readonly Readonly<{ segment: LayoutTextSeg; start: number; end: number }>[],
  joinedText: string,
  start: number,
  end: number,
  measurer: TextMeasurer,
  gridDeltaPt: number,
): number {
  let widthPt = 0;
  for (const piece of pieces) {
    const overlapStart = Math.max(start, piece.start);
    const overlapEnd = Math.min(end, piece.end);
    if (overlapStart >= overlapEnd) continue;
    const text = joinedText.slice(overlapStart, overlapEnd);
    const candidate = { ...piece.segment, text };
    if (candidate.textLayoutService && candidate.textShapeRequest) {
      const shaped = candidate.textLayoutService.shape({
        ...candidate.textShapeRequest,
        text,
        fontSizePt: calcEffectiveFontPx(candidate, 1),
        measure: true,
        clusterGeometry: false,
      });
      widthPt += segAdvanceWidth(candidate, shaped.advancePt, gridDeltaPt, 1);
      continue;
    }
    measurer.context.font = buildFont(
      candidate.bold,
      candidate.italic,
      calcEffectiveFontPx(candidate, 1),
      candidate.fontFamily,
      measurer.fontFamilyClasses as Record<string, string>,
      candidate.fontRoute,
    );
    widthPt += segAdvanceWidth(
      candidate,
      measurer.context.measureText(text).width,
      gridDeltaPt,
      1,
    );
  }
  return widthPt;
}

function minimumTextAtomWidthPt(
  segments: readonly LayoutSeg[],
  context: ParagraphLayoutContext,
  measurer: TextMeasurer,
): number {
  const gridDeltaPt = context.characterGrid.active ? context.characterGrid.deltaPt : 0;
  let maximumPt = 0;
  for (let segmentIndex = 0; segmentIndex < segments.length; segmentIndex += 1) {
    const segment = segments[segmentIndex];
    if (!('text' in segment) || segment.text.length === 0) continue;
    const pieces: Array<{ segment: LayoutTextSeg; start: number; end: number }> = [];
    let joinedText = '';
    const append = (piece: LayoutTextSeg): void => {
      const start = joinedText.length;
      joinedText += piece.text;
      pieces.push({ segment: piece, start, end: joinedText.length });
    };
    append(segment);
    while (segmentIndex + 1 < segments.length) {
      const following = segments[segmentIndex + 1];
      if (!('text' in following) || following.joinPrev !== true) break;
      append(following);
      segmentIndex += 1;
    }

    let tokenStart = 0;
    for (const token of splitTextForLayout(joinedText)) {
      const trimmed = token.replace(/\s+$/u, '');
      const trimmedStart = tokenStart;
      const trimmedEnd = tokenStart + trimmed.length;
      tokenStart += token.length;
      if (!trimmed) continue;
      if (!hasCJKBreakOpportunity(trimmed)) {
        maximumPt = Math.max(
          maximumPt,
          measureTextRange(
            pieces,
            joinedText,
            trimmedStart,
            trimmedEnd,
            measurer,
            gridDeltaPt,
          ),
        );
        continue;
      }

      const boundaries = [0, ...graphemeClusterOffsets(trimmed), trimmed.length];
      const clusters: Array<{ text: string; start: number; end: number }> = [];
      for (let index = 1; index < boundaries.length; index += 1) {
        clusters.push({
          text: trimmed.slice(boundaries[index - 1], boundaries[index]),
          start: trimmedStart + boundaries[index - 1],
          end: trimmedStart + boundaries[index],
        });
      }
      const atoms: Array<{ text: string; start: number; end: number }> = [];
      let atom = clusters[0];
      for (let index = 1; index < clusters.length; index += 1) {
        const previous = [...atom.text].at(-1)?.codePointAt(0);
        const next = clusters[index].text.codePointAt(0);
        const breakAllowed = previous !== undefined
          && next !== undefined
          && !context.kinsoku.lineEndForbidden.has(previous)
          && !context.kinsoku.lineStartForbidden.has(next);
        if (breakAllowed) {
          atoms.push(atom);
          atom = clusters[index];
        } else {
          atom = {
            text: atom.text + clusters[index].text,
            start: atom.start,
            end: clusters[index].end,
          };
        }
      }
      if (atom) atoms.push(atom);
      for (const unbreakable of atoms) {
        maximumPt = Math.max(
          maximumPt,
          measureTextRange(
            pieces,
            joinedText,
            unbreakable.start,
            unbreakable.end,
            measurer,
            gridDeltaPt,
          ),
        );
      }
    }
  }
  return maximumPt;
}

function logicalLineInterval(
  line: LayoutLine,
  lineIndex: number,
  context: ParagraphLayoutContext,
): Readonly<{ startPt: number; endPt: number }> {
  const leadingIndentPt = context.baseRtl
    ? context.physicalIndentRightPt
    : context.physicalIndentLeftPt;
  const startPt = leadingIndentPt
    + (lineIndex === 0 ? context.firstIndentPt : 0)
    + line.xOffset;
  const widthPt = line.segments.reduce((sum, segment) => sum + segment.measuredWidth, 0);
  return { startPt, endPt: startPt + widthPt };
}

export function measureParagraphIntrinsicWidths(
  paragraph: DocParagraph,
  context: ParagraphLayoutContext,
  maximumWidthPt: number,
  measurer: TextMeasurer,
  environment: ParagraphMeasurementEnvironment,
  numbering?: NumberingMarkerGeometry,
): ParagraphIntrinsicWidths {
  if (!Number.isFinite(maximumWidthPt) || maximumWidthPt < 0) {
    throw new RangeError('maximumWidthPt must be finite and non-negative');
  }
  if (maximumWidthPt === 0) return { minWidthPt: 0, maxWidthPt: 0 };

  const segments = mergeCompatibleTextSegments(buildSegments(paragraph.runs, environment));
  const paragraphWidthPt = Math.max(
    1,
    maximumWidthPt - context.physicalIndentLeftPt - context.physicalIndentRightPt,
  );
  const lines = segments.length === 0 ? [] : layoutLines(
    measurer.context,
    segments,
    paragraphWidthPt,
    context.firstIndentPt,
    1,
    [...context.tabStops],
    undefined,
    measurer.fontFamilyClasses as Record<string, string>,
    context.physicalIndentLeftPt,
    context.kinsoku,
    context.characterGrid.active ? context.characterGrid.deltaPt : 0,
    context.defaultTabPt,
    paragraphWidthPt + context.physicalIndentRightPt,
    context.baseRtl,
    context.isJustified,
    context.stretchLastLine,
    undefined,
    'intrinsic',
  );
  const oppositeIndentPt = context.baseRtl
    ? context.physicalIndentLeftPt
    : context.physicalIndentRightPt;
  let minimumLeftPt = 0;
  let maximumRightPt = 0;
  lines.forEach((line, index) => {
    const interval = logicalLineInterval(line, index, context);
    minimumLeftPt = Math.min(minimumLeftPt, interval.startPt);
    maximumRightPt = Math.max(maximumRightPt, interval.endPt);
  });
  const markerInterval = numbering ? numberingMarkerLogicalInterval({
    leadingIndentPt: context.baseRtl
      ? context.physicalIndentRightPt
      : context.physicalIndentLeftPt,
    authoredFirstIndentPt: paragraph.indentFirst,
    markerShiftPt: numbering.markerShiftPt,
    markerWidthPt: numbering.markerWidthPt,
  }) : undefined;
  if (markerInterval) {
    minimumLeftPt = Math.min(minimumLeftPt, markerInterval.startPt);
    maximumRightPt = Math.max(maximumRightPt, markerInterval.endPt);
  }
  const maxWidthPt = Math.min(
    maximumWidthPt,
    Math.max(0, maximumRightPt - minimumLeftPt + oppositeIndentPt),
  );

  let minimumAtomPt = minimumTextAtomWidthPt(segments, context, measurer);
  for (const line of lines) {
    let penPt = 0;
    const lineWidthPt = line.segments.reduce((sum, segment) => sum + segment.measuredWidth, 0);
    for (const segment of line.segments) {
      penPt += segment.measuredWidth;
      if ('imagePath' in segment && !segment.anchor) {
        minimumAtomPt = Math.max(minimumAtomPt, segment.measuredWidth);
      } else if ('mathNodes' in segment) {
        minimumAtomPt = Math.max(minimumAtomPt, segment.measuredWidth);
      } else if ('isTab' in segment) {
        minimumAtomPt = Math.max(
          minimumAtomPt,
          segment.resolvedAlignment === 'left' ? penPt : lineWidthPt,
        );
      }
    }
  }
  const leadingIndentPt = context.baseRtl
    ? context.physicalIndentRightPt
    : context.physicalIndentLeftPt;
  const continuationStartPt = leadingIndentPt;
  let minLeftPt = Math.min(0, continuationStartPt);
  let minRightPt = Math.max(0, continuationStartPt + minimumAtomPt);
  const firstStartPt = leadingIndentPt + context.firstIndentPt;
  minLeftPt = Math.min(minLeftPt, firstStartPt);
  minRightPt = Math.max(minRightPt, firstStartPt + minimumAtomPt);
  if (markerInterval) {
    minLeftPt = Math.min(minLeftPt, markerInterval.startPt);
    minRightPt = Math.max(minRightPt, markerInterval.endPt);
  }
  const minWidthPt = Math.min(
    maximumWidthPt,
    Math.max(0, minRightPt - minLeftPt + oppositeIndentPt),
  );
  return { minWidthPt, maxWidthPt };
}
