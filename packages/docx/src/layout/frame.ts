import type { BodyElement, DocParagraph, FramePr } from '../types.js';
import type { ParagraphLayoutContext } from '../layout-context.js';
import { buildSegments, layoutLines } from '../line-layout.js';
import type {
  ParagraphMeasurementEnvironment,
  TextMeasurer,
} from '../paragraph-measure.js';
import { stableFingerprint } from './fingerprint.js';
import {
  numberingMarkerLogicalInterval,
  type NumberingMarkerGeometry,
} from './numbering-marker.js';
import {
  resolveParagraphBorderEdges,
  type ParagraphBorderEdges,
} from './paragraph-border-adjacency.js';

/** Parser-private effective CT_FramePr state required for grouping, not API. */
type EffectiveFramePr = FramePr & Readonly<{ __anchorLock?: boolean }>;

export interface BodyFrameGroup {
  readonly id: string;
  readonly owner: DocParagraph;
  readonly members: readonly DocParagraph[];
  readonly sourceIndices: readonly number[];
  readonly framePr: FramePr;
}

/**
 * Resolve a paragraph's non-wrapping natural width, capped by its real anchor
 * band. ECMA-376 §17.3.1.11 leaves an absent frame width to the application,
 * but the admissible width is still bounded by the frame's anchor container.
 * Keeping both the no-wrap policy and that bound explicit avoids the former
 * synthetic "very wide page" measurement.
 *
 * The returned width is only the intrinsic-width probe. Callers must acquire
 * the paragraph again at the selected final width; this probe is never retained
 * as final line geometry.
 */
export function measureParagraphIntrinsicWidth(
  paragraph: DocParagraph,
  context: ParagraphLayoutContext,
  maximumWidthPt: number,
  measurer: TextMeasurer,
  environment: ParagraphMeasurementEnvironment,
  numbering?: NumberingMarkerGeometry,
): number {
  if (!Number.isFinite(maximumWidthPt) || maximumWidthPt < 0) {
    throw new RangeError('maximumWidthPt must be finite and non-negative');
  }
  if (maximumWidthPt === 0) return 0;
  const paragraphWidthPt = Math.max(
    1,
    maximumWidthPt - context.physicalIndentLeftPt - context.physicalIndentRightPt,
  );
  const segments = buildSegments(paragraph.runs, environment);
  if (segments.length === 0 && !numbering) return 0;
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
  const leadingIndentPt = context.baseRtl
    ? context.physicalIndentRightPt
    : context.physicalIndentLeftPt;
  const oppositeIndentPt = context.baseRtl
    ? context.physicalIndentLeftPt
    : context.physicalIndentRightPt;
  let minimumLeftPt = 0;
  let maximumRightPt = 0;
  lines.forEach((line, index) => {
    const lineLeftPt = leadingIndentPt
      + (index === 0 ? context.firstIndentPt : 0)
      + line.xOffset;
    const lineWidthPt = line.segments.reduce(
      (widthPt, segment) => widthPt + segment.measuredWidth,
      0,
    );
    minimumLeftPt = Math.min(minimumLeftPt, lineLeftPt);
    maximumRightPt = Math.max(maximumRightPt, lineLeftPt + lineWidthPt);
  });
  const bodyNaturalWidthPt = maximumRightPt - minimumLeftPt + oppositeIndentPt;
  // The marker is not part of the run segments. Its final shaped/picture
  // geometry therefore has to participate explicitly in the same logical
  // leading-edge coordinate system used by retained marker placement.
  const markerInterval = numbering ? numberingMarkerLogicalInterval({
    leadingIndentPt,
    authoredFirstIndentPt: paragraph.indentFirst,
    markerShiftPt: numbering.markerShiftPt,
    markerWidthPt: numbering.markerWidthPt,
  }) : undefined;
  const naturalWidthPt = numbering
    ? Math.max(maximumRightPt, markerInterval!.endPt)
      - Math.min(minimumLeftPt, markerInterval!.startPt)
      + oppositeIndentPt
    : bodyNaturalWidthPt;
  return Math.min(maximumWidthPt, naturalWidthPt);
}

/**
 * ECMA-376 §17.3.1.11 groups adjacent paragraphs only when every effective
 * framePr value is identical. The explicit tuple is intentional: the parser's
 * private wire facts participate without making object property order or a
 * future unrelated runtime property part of document identity.
 */
export function effectiveFrameIdentity(framePr: FramePr): string {
  const frame = framePr as EffectiveFramePr;
  return stableFingerprint('w:framePr', [
    frame.dropCap, frame.lines, frame.wrap, frame.hAnchor, frame.vAnchor,
    frame.hRule, frame.hSpace, frame.vSpace,
    frame.w ?? null, frame.h ?? null, frame.x ?? null, frame.y ?? null,
    frame.xAlign ?? null, frame.yAlign ?? null,
    frame.__anchorLock === true,
  ]);
}

/** Build the body-local adjacency groups once, before pagination mutates pages. */
export function collectBodyFrameGroups(
  body: readonly BodyElement[],
): WeakMap<DocParagraph, BodyFrameGroup> {
  const result = new WeakMap<DocParagraph, BodyFrameGroup>();
  for (let index = 0; index < body.length;) {
    const element = body[index];
    if (element?.type !== 'paragraph' || !element.framePr) {
      index += 1;
      continue;
    }
    const identity = effectiveFrameIdentity(element.framePr);
    const members: DocParagraph[] = [element];
    const sourceIndices: number[] = [index];
    let next = index + 1;
    while (next < body.length) {
      const candidate = body[next];
      if (
        candidate?.type !== 'paragraph'
        || !candidate.framePr
        || effectiveFrameIdentity(candidate.framePr) !== identity
      ) break;
      members.push(candidate);
      sourceIndices.push(next);
      next += 1;
    }
    const group = Object.freeze({
      id: `${identity}:${index}`,
      owner: element,
      members: Object.freeze(members),
      sourceIndices: Object.freeze(sourceIndices),
      framePr: element.framePr,
    });
    for (const member of members) result.set(member, group);
    index = next;
  }
  return result;
}

const bodyFrameGroups = new WeakMap<DocParagraph, BodyFrameGroup>();
const bodyParagraphBorderEdges = new WeakMap<DocParagraph, ParagraphBorderEdges>();
const bodySourceIndices = new WeakMap<object, number>();

function registerBodySourceIdentity(value: unknown, sourceIndex: number, seen: WeakSet<object>): void {
  if (!value || typeof value !== 'object' || seen.has(value)) return;
  seen.add(value);
  const object = value as object;
  if (!bodySourceIndices.has(object)) bodySourceIndices.set(object, sourceIndex);
  if (Array.isArray(value)) {
    value.forEach((item) => registerBodySourceIdentity(item, sourceIndex, seen));
    return;
  }
  Object.values(value).forEach((item) => registerBodySourceIdentity(item, sourceIndex, seen));
}

function sourceIndexFromIdentity(value: unknown, seen: WeakSet<object>): number | undefined {
  if (!value || typeof value !== 'object' || seen.has(value)) return undefined;
  seen.add(value);
  const object = value as object;
  const direct = bodySourceIndices.get(object);
  if (direct !== undefined) return direct;
  const values = Array.isArray(value) ? value : Object.values(value);
  for (const item of values) {
    const nested = sourceIndexFromIdentity(item, seen);
    if (nested !== undefined) return nested;
  }
  return undefined;
}

/** Prepare body identity/adjacency metadata independently of layout services. */
export function prepareBodyFrameMetadata(body: readonly BodyElement[]): void {
  const groups = collectBodyFrameGroups(body);
  for (let index = 0; index < body.length; index += 1) {
    const element = body[index]!;
    // Pagination emits shallow/sliced table clones inside the frozen kernel.
    // Register their retained child identities at the document-session seam so
    // fragment acquisition can recover the source path without threading `i`
    // through computePages.
    registerBodySourceIdentity(element, index, new WeakSet());
    if (element.type !== 'paragraph') continue;
    const group = groups.get(element);
    if (group) bodyFrameGroups.set(element, group);
    const previous = body[index - 1];
    const next = body[index + 1];
    const groupedPrevious = previous?.type === 'paragraph' && groups.get(previous) === group
      ? previous : null;
    const groupedNext = next?.type === 'paragraph' && groups.get(next) === group
      ? next : null;
    bodyParagraphBorderEdges.set(element, resolveParagraphBorderEdges(
      group ? groupedPrevious : previous?.type === 'paragraph' ? previous : null,
      element,
      group ? groupedNext : next?.type === 'paragraph' ? next : null,
      group !== undefined,
    ));
  }
}

export const bodyFrameGroupFor = (paragraph: DocParagraph): BodyFrameGroup | undefined =>
  bodyFrameGroups.get(paragraph);

export const bodySourceIndexFor = (element: object): number | undefined =>
  sourceIndexFromIdentity(element, new WeakSet());

export const bodyParagraphBorderEdgesFor = (
  paragraph: DocParagraph,
): ParagraphBorderEdges | undefined => bodyParagraphBorderEdges.get(paragraph);
