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
  lineBelowBaselinePx,
  lineBoxHeight,
  paragraphMarkBelowBaselinePt,
  paragraphMarkLineHeight,
  type DocGridCtx,
  type LineBoundary,
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
    /** The paragraph's COLUMN band, scoping the topAndBottom gate (§20.4.2.20 /
     *  §17.6.4) to the column the float is anchored in — NOT the indented text
     *  band `paragraphXPt`/`maximumWidthPt` the square side-gap math uses. */
    readonly columnXPt: number;
    readonly columnWidthPt: number;
  }): {
    readonly topYPt: number;
    readonly xOffsetPt: number;
    readonly maximumWidthPt: number;
  };
  skipTopAndBottomBands(input: {
    readonly yPt: number;
    /** The paragraph's COLUMN band (colX()/colW()), used to scope a topAndBottom
     *  float to the column it is anchored in (§20.4.2.20 / §17.6.4) — NOT the
     *  indented text band `lineWindow` uses. */
    readonly columnXPt: number;
    readonly columnWidthPt: number;
  }): number;
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
      columnXPt,
      columnWidthPt,
    }) => {
      const window = resolveLineFloatWindow(
        topYPt,
        minimumStartWidthPt,
        probeHeightPt,
        paragraphXPt,
        maximumWidthPt,
        activeFloats,
        columnXPt,
        columnXPt + columnWidthPt,
      );
      return {
        topYPt: window.topY,
        xOffsetPt: window.xOffset,
        maximumWidthPt: window.maxWidth,
      };
    },
    skipTopAndBottomBands: ({ yPt, columnXPt, columnWidthPt }) =>
      skipPastTopAndBottom(yPt, activeFloats, columnXPt, columnXPt + columnWidthPt),
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
  /** ECMA-376 §17.3.3.25 paragraph-wide uniform line advance in points, snapped
   *  to the docGrid. Zero when the paragraph has no ruby. */
  readonly uniformRubyAdvancePt: number;
  readonly contentStartYPt: number;
  readonly contentEndYPt: number;
  /**
   * ECMA-376 §17.3.1.29 / §17.3.1.33 — the extent (pt) of the LAST line's box that
   * lies below its baseline (descent + half of any auto/atLeast leading). Word's
   * page fit is baseline-based: a line whose baseline sits within the text area may
   * let this below-baseline whitespace extend into the bottom margin. The paginator
   * uses it (for an empty paragraph, whose mark line paints no ink there) so a
   * trailing empty paragraph is not pushed to the next page merely because its
   * invisible mark box grazes past the bottom content edge.
   */
  readonly lastLineBelowBaselinePt: number;
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
  continuation?: {
    readonly boundary: LineBoundary;
    readonly uniformRubyAdvancePt?: number;
  },
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
    // §20.4.2.20 / §17.6.4 column scope: pass this paragraph's COLUMN band
    // (placement.paragraphXPt / availableWidthPt = colX()/colW()), NOT the
    // indented text band, so measure agrees bit-for-bit with the paint pass,
    // which scopes the same skip against state.contentX/contentW (the column
    // band). A topAndBottom float anchored in another newspaper column is
    // filtered out in both passes.
    cursorPt = placement.wrap.skipTopAndBottomBands({
      yPt: cursorPt,
      columnXPt: placement.paragraphXPt,
      columnWidthPt: placement.availableWidthPt,
    });
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
        // §20.4.2.20 / §17.6.4 column scope: the topAndBottom gate sees the raw
        // COLUMN band, not the indented mark band above.
        columnXPt: placement.paragraphXPt,
        columnWidthPt: placement.availableWidthPt,
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
      uniformRubyAdvancePt: 0,
      contentStartYPt: markTopPt,
      contentEndYPt: markTopPt + markAdvancePt,
      lastLineBelowBaselinePt: paragraphMarkBelowBaselinePt(
        paragraph,
        grid,
        context.hasRuby,
        environment.documentHasEastAsianText === true,
        measurer.context,
        fontFamilyClasses,
        context.lineSpacing,
      ),
      placement: recordedPlacement,
    };
  };

  const segments = buildSegments(paragraph.runs, environment);
  if (segments.length === 0) return measureMarkOnly();

  const wrapContext: WrapLayoutCtx | undefined = placement.wrap
    ? {
        startPageY: cursorPt,
        paraX: paragraphXPt,
        // Raw COLUMN band (placement) for the topAndBottom gate; paraX above is
        // the indented text band for the square side-gap math (§20.4.2.20 vs
        // §20.4.2.17). See WrapLayoutCtx.columnXPt.
        columnXPt: placement.paragraphXPt,
        columnWidthPt: placement.availableWidthPt,
        floats: [],
        lineWindow: (input) => placement.wrap!.lineWindow(input),
        lineBoxH: (ascent, descent, _hasRuby, intendedSingle, emPx, eastAsian) => lineBoxHeight(
          context.lineSpacing,
          ascent,
          descent,
          1,
          grid,
          context.hasRuby,
          intendedSingle ?? 0,
          // §17.6.5 cell rounding follows this line's script, matching text boxes;
          // ruby paragraphs retain their established uniform paragraph resolver.
          context.hasRuby ? context.hasEastAsianText : (eastAsian ?? false),
          emPx,
        ),
        pageH: placement.maximumYPt,
      }
    : undefined;
  const lines = layoutLines(
    measurer.context,
    segments,
    paragraphWidthPt,
    // ECMA-376 §17.3.1.12: first-line and hanging indents apply only to the
    // paragraph's first line, not to a continuation measured in another column.
    continuation ? 0 : context.firstIndentPt,
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
    context.isJustified,
    context.stretchLastLine,
    continuation?.boundary,
  );
  if (lines.length === 0) return measureMarkOnly();

  let uniformRubyAdvancePt = context.hasRuby
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
          line.height,
        ))),
        grid,
      )
    : 0;
  if (context.hasRuby && continuation?.uniformRubyAdvancePt !== undefined) {
    uniformRubyAdvancePt = Math.max(
      uniformRubyAdvancePt,
      continuation.uniformRubyAdvancePt,
    );
  }
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
          // §17.6.5 cell rounding is gated by the line's script; a Latin-only
          // line in a CJK paragraph keeps its natural height.
          line.eastAsian ?? false,
          line.height,
        );
    measuredLines.push({ layout: line, topYPt, advancePt });
    cursorPt = topYPt + advancePt;
  }

  const lastLine = measuredLines[measuredLines.length - 1];
  return {
    lines: measuredLines,
    markOnly: false,
    requestedSpaceBeforePt,
    requestedSpaceAfterPt,
    uniformRubyAdvancePt,
    contentStartYPt: measuredLines[0].topYPt,
    contentEndYPt: cursorPt,
    lastLineBelowBaselinePt: lineBelowBaselinePx(
      lastLine.advancePt,
      lastLine.layout.ascent,
      lastLine.layout.descent,
    ),
    placement: recordedPlacement,
  };
}
