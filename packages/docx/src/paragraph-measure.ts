import type { ParagraphLayoutContext } from './layout-context.js';
import {
  resolveLineFloatWindow,
  skipPastTopAndBottom,
  type FloatRect,
} from './float-layout.js';
import {
  buildSegments,
  getDefaultFontSize,
  isGridLineRule,
  layoutLines,
  lineBoxHeight,
  paragraphMarkLineHeight,
  type DocGridCtx,
  type LayoutLine,
  type LineLayoutEnvironment,
  type WrapLayoutCtx,
} from './line-layout.js';
import type { DocParagraph } from './types.js';

export type { LineLayoutEnvironment } from './line-layout.js';

export interface ParagraphMeasurementEnvironment extends LineLayoutEnvironment {
  readonly documentHasEastAsianText: boolean;
}

export interface WrapOracle {
  lineWindow(input: {
    readonly topYPt: number;
    readonly minimumStartWidthPt: number;
    readonly probeHeightPt: number;
    readonly paragraphXPt: number;
    readonly maximumWidthPt: number;
  }): {
    readonly topYPt: number;
    readonly xOffsetPt: number;
    readonly maximumWidthPt: number;
  };
  skipTopAndBottomBands(yPt: number): number;
}

/** Adapt registered scale-1 float rectangles to the placement-aware paragraph
 * measurement boundary. Float discovery, registration, and compatibility
 * behavior remain owned by the renderer. */
export function createFloatWrapOracle(floats: readonly FloatRect[]): WrapOracle {
  const activeFloats = [...floats];
  return {
    lineWindow: ({
      topYPt,
      minimumStartWidthPt,
      probeHeightPt,
      paragraphXPt,
      maximumWidthPt,
    }) => {
      const window = resolveLineFloatWindow(
        topYPt,
        minimumStartWidthPt,
        probeHeightPt,
        paragraphXPt,
        maximumWidthPt,
        activeFloats,
      );
      return {
        topYPt: window.topY,
        xOffsetPt: window.xOffset,
        maximumWidthPt: window.maxWidth,
      };
    },
    skipTopAndBottomBands: (yPt) => skipPastTopAndBottom(yPt, activeFloats),
  };
}

export interface TextMeasurer {
  readonly context:
    | CanvasRenderingContext2D
    | OffscreenCanvasRenderingContext2D;
  readonly fontFamilyClasses: Readonly<Record<string, string>>;
}

export interface ParagraphPlacement {
  readonly startYPt: number;
  readonly paragraphXPt: number;
  readonly availableWidthPt: number;
  readonly maximumYPt: number;
  readonly suppressSpaceBefore: boolean;
  readonly wrap?: WrapOracle;
}

export interface MeasuredLine {
  readonly layout: LayoutLine;
  readonly topYPt: number;
  readonly advancePt: number;
}

export interface MeasuredParagraph {
  readonly lines: readonly MeasuredLine[];
  readonly markOnly: boolean;
  readonly requestedSpaceBeforePt: number;
  readonly requestedSpaceAfterPt: number;
  readonly contentStartYPt: number;
  readonly contentEndYPt: number;
  readonly placement: Readonly<ParagraphPlacement>;
}

function paragraphGrid(context: ParagraphLayoutContext): DocGridCtx {
  return {
    type: context.lineGrid.active ? 'lines' : null,
    linePitchPt: context.lineGrid.active ? context.lineGrid.pitchPt : null,
    charSpacePt: context.characterGrid.active ? context.characterGrid.deltaPt : null,
  };
}

/** Preserve the renderer's paragraph-wide ruby/docGrid height calculation. */
function snapParagraphLineToGrid(heightPt: number, grid: DocGridCtx): number {
  if (!isGridLineRule(grid)) return heightPt;
  const pitchPt = grid.linePitchPt!;
  if (pitchPt <= 0) return heightPt;
  if (heightPt <= pitchPt) return pitchPt;
  return Math.ceil(heightPt / pitchPt) * pitchPt;
}

export function measureParagraph(
  paragraph: DocParagraph,
  context: ParagraphLayoutContext,
  placement: ParagraphPlacement,
  measurer: TextMeasurer,
  environment: ParagraphMeasurementEnvironment,
): MeasuredParagraph {
  const grid = paragraphGrid(context);
  const paragraphWidthPt = Math.max(
    1,
    placement.availableWidthPt
      - context.physicalIndentLeftPt
      - context.physicalIndentRightPt,
  );
  const paragraphXPt = placement.paragraphXPt + context.physicalIndentLeftPt;
  const requestedSpaceBeforePt = context.spaceBeforePt;
  const requestedSpaceAfterPt = context.spaceAfterPt;
  const recordedPlacement = Object.freeze({ ...placement });
  const fontFamilyClasses = measurer.fontFamilyClasses as Record<string, string>;

  let cursorPt = placement.startYPt
    + (placement.suppressSpaceBefore ? 0 : requestedSpaceBeforePt);
  if (placement.wrap) {
    cursorPt = placement.wrap.skipTopAndBottomBands(cursorPt);
  }

  const measureMarkOnly = (): MeasuredParagraph => {
    let markTopPt = cursorPt;
    if (placement.wrap) {
      markTopPt = placement.wrap.lineWindow({
        topYPt: markTopPt,
        minimumStartWidthPt: getDefaultFontSize(paragraph),
        probeHeightPt: 10,
        paragraphXPt,
        maximumWidthPt: paragraphWidthPt,
      }).topYPt;
    }
    const markAdvancePt = paragraphMarkLineHeight(
      paragraph,
      1,
      grid,
      context.hasRuby,
      environment.documentHasEastAsianText === true,
      measurer.context,
      fontFamilyClasses,
      context.lineSpacing,
    );
    return {
      lines: [],
      markOnly: true,
      requestedSpaceBeforePt,
      requestedSpaceAfterPt,
      contentStartYPt: markTopPt,
      contentEndYPt: markTopPt + markAdvancePt,
      placement: recordedPlacement,
    };
  };

  const segments = buildSegments(paragraph.runs, environment);
  if (segments.length === 0) return measureMarkOnly();

  const wrapContext: WrapLayoutCtx | undefined = placement.wrap
    ? {
        startPageY: cursorPt,
        paraX: paragraphXPt,
        floats: [],
        lineWindow: (input) => placement.wrap!.lineWindow(input),
        lineBoxH: (ascent, descent, _hasRuby, intendedSingle) => lineBoxHeight(
          context.lineSpacing,
          ascent,
          descent,
          1,
          grid,
          context.hasRuby,
          intendedSingle ?? 0,
          context.hasEastAsianText,
        ),
        pageH: placement.maximumYPt,
      }
    : undefined;
  const lines = layoutLines(
    measurer.context,
    segments,
    paragraphWidthPt,
    context.firstIndentPt,
    1,
    [...context.tabStops],
    wrapContext,
    fontFamilyClasses,
    context.physicalIndentLeftPt,
    context.kinsoku,
    context.characterGrid.active ? context.characterGrid.deltaPt : 0,
    context.defaultTabPt,
    paragraphWidthPt + context.physicalIndentRightPt,
    context.baseRtl,
  );
  if (lines.length === 0) return measureMarkOnly();

  const uniformRubyAdvancePt = context.hasRuby
    ? snapParagraphLineToGrid(
        Math.max(0, ...lines.map((line) => lineBoxHeight(
          context.lineSpacing,
          line.ascent,
          line.descent,
          1,
          grid,
          true,
          line.intendedSingle,
          context.hasEastAsianText,
        ))),
        grid,
      )
    : 0;
  const measuredLines: MeasuredLine[] = [];
  for (const line of lines) {
    const topYPt = line.topY !== undefined && line.topY > cursorPt
      ? line.topY
      : cursorPt;
    const advancePt = context.hasRuby
      ? uniformRubyAdvancePt
      : lineBoxHeight(
          context.lineSpacing,
          line.ascent,
          line.descent,
          1,
          grid,
          false,
          line.intendedSingle,
          context.hasEastAsianText,
        );
    measuredLines.push({ layout: line, topYPt, advancePt });
    cursorPt = topYPt + advancePt;
  }

  return {
    lines: measuredLines,
    markOnly: false,
    requestedSpaceBeforePt,
    requestedSpaceAfterPt,
    contentStartYPt: measuredLines[0].topYPt,
    contentEndYPt: cursorPt,
    placement: recordedPlacement,
  };
}
