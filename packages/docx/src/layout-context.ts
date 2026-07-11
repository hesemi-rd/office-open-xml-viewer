import {
  resolveKinsokuRules,
  type KinsokuRules,
} from '@silurus/ooxml-core';
import {
  EAST_ASIAN_RE,
  resolveDefaultTabPt,
  type DocGridCtx,
} from './line-layout.js';
import { jcIsFullyJustified, jcStretchesLastLine } from './bidi-line.js';
import type {
  BodyElement,
  ColumnGeom,
  DocParagraph,
  DocTable,
  DocxDocumentModel,
  DocxTextRun,
  LineNumbering,
  LineSpacing,
  SectionGeom,
  SectionProps,
  TabStop,
} from './types.js';

export interface DocumentLayoutSettings {
  readonly kinsoku: KinsokuRules;
  readonly defaultTabPt: number;
  readonly characterSpacingControl?: string;
  readonly mathDefJc?: string;
  readonly documentHasEastAsianText: boolean;
  readonly compat: {
    readonly adjustLineHeightInTable: boolean;
    readonly useFeLayout: boolean;
    readonly balanceSingleByteDoubleByteWidth: boolean;
  };
}

export interface SectionGridContext {
  readonly kind: 'none' | 'lines' | 'linesAndChars' | 'snapToChars';
  readonly linePitchPt: number | null;
  readonly charSpacePt: number | null;
}

export interface SectionLayoutContext {
  readonly geometry: SectionGeom;
  readonly columns: readonly ColumnGeom[];
  readonly grid: SectionGridContext;
  readonly textDirection: string;
  readonly verticalAlignment: string;
  readonly lineNumbering?: LineNumbering;
}

export type StoryKind =
  | 'body'
  | 'header'
  | 'footer'
  | 'footnote'
  | 'endnote'
  | 'textbox';

export type ContainerFrame = { readonly kind: 'tableCell' };

export interface StoryContext {
  readonly story: StoryKind;
  readonly containers: readonly ContainerFrame[];
  readonly lineNumberingEligible: boolean;
}

export function enterTableCellStoryContext(parent: StoryContext): StoryContext {
  return {
    story: parent.story,
    containers: [...parent.containers, { kind: 'tableCell' }],
    lineNumberingEligible: false,
  };
}

export interface LineGridPolicy {
  readonly active: boolean;
  readonly pitchPt: number | null;
}

export interface CharacterGridPolicy {
  readonly active: boolean;
  readonly deltaPt: number;
}

export interface ParagraphLayoutContext {
  readonly lineGrid: LineGridPolicy;
  readonly characterGrid: CharacterGridPolicy;
  readonly physicalIndentLeftPt: number;
  readonly physicalIndentRightPt: number;
  readonly firstIndentPt: number;
  readonly lineSpacing: LineSpacing | null;
  readonly spaceBeforePt: number;
  readonly spaceAfterPt: number;
  readonly baseRtl: boolean;
  readonly isJustified: boolean;
  readonly stretchLastLine: boolean;
  readonly tabStops: readonly TabStop[];
  readonly hasRuby: boolean;
  readonly hasEastAsianText: boolean;
  readonly kinsoku: KinsokuRules;
  readonly defaultTabPt: number;
}

export interface RunLayoutContext {
  readonly characterGrid: CharacterGridPolicy;
}

function paragraphHasRuby(paragraph: DocParagraph): boolean {
  return paragraph.runs.some(
    (run) => run.type === 'text' && Boolean((run as DocxTextRun).ruby),
  );
}

function paragraphHasEastAsianText(paragraph: DocParagraph): boolean {
  return paragraph.runs.some(
    (run) => run.type === 'text' && EAST_ASIAN_RE.test((run as DocxTextRun).text),
  );
}

export function documentHasEastAsianText(body: readonly BodyElement[]): boolean {
  for (const element of body) {
    if (element.type === 'paragraph') {
      if (paragraphHasEastAsianText(element as DocParagraph)) return true;
      continue;
    }
    if (element.type !== 'table') continue;
    for (const row of (element as DocTable).rows) {
      for (const cell of row.cells) {
        if (documentHasEastAsianText(cell.content)) return true;
      }
    }
  }
  return false;
}

export function resolveDocumentLayoutSettings(
  document: DocxDocumentModel,
): DocumentLayoutSettings {
  return {
    kinsoku: resolveKinsokuRules(document.settings),
    defaultTabPt: resolveDefaultTabPt(document.settings),
    characterSpacingControl: document.settings?.characterSpacingControl,
    mathDefJc: document.settings?.mathDefJc,
    documentHasEastAsianText: documentHasEastAsianText(document.body),
    compat: {
      adjustLineHeightInTable: document.settings?.adjustLineHeightInTable ?? false,
      useFeLayout: document.settings?.useFeLayout ?? false,
      balanceSingleByteDoubleByteWidth:
        document.settings?.balanceSingleByteDoubleByteWidth ?? false,
    },
  };
}

export function computeSectionColumns(section: SectionProps): ColumnGeom[] {
  const contentWidthPt = section.pageWidth - section.marginLeft - section.marginRight;
  const columns = section.columns;
  if (!columns || columns.count <= 1) {
    return [{ xPt: section.marginLeft, wPt: Math.max(1, contentWidthPt) }];
  }

  if (!columns.equalWidth && columns.cols.length > 0) {
    const result: ColumnGeom[] = [];
    let xPt = section.marginLeft;
    for (const column of columns.cols) {
      result.push({ xPt, wPt: Math.max(1, column.widthPt) });
      xPt += column.widthPt + column.spacePt;
    }
    return result;
  }

  const widthPt = Math.max(
    1,
    (contentWidthPt - (columns.count - 1) * columns.spacePt) / columns.count,
  );
  return Array.from({ length: columns.count }, (_, index) => ({
    xPt: section.marginLeft + index * (widthPt + columns.spacePt),
    wPt: widthPt,
  }));
}

function normalizeGridKind(type: string | null | undefined): SectionGridContext['kind'] {
  switch (type) {
    case 'lines':
    case 'linesAndChars':
    case 'snapToChars':
      return type;
    default:
      return 'none';
  }
}

export function isSectionLineGrid(kind: SectionGridContext['kind']): boolean {
  return kind === 'lines' || kind === 'linesAndChars' || kind === 'snapToChars';
}

export function isSectionCharacterGrid(kind: SectionGridContext['kind']): boolean {
  return kind === 'linesAndChars' || kind === 'snapToChars';
}

export function resolveSectionLayoutContext(
  _settings: DocumentLayoutSettings,
  section: SectionProps,
): SectionLayoutContext {
  return {
    geometry: {
      pageWidth: section.pageWidth,
      pageHeight: section.pageHeight,
      marginTop: section.marginTop,
      marginRight: section.marginRight,
      marginBottom: section.marginBottom,
      marginLeft: section.marginLeft,
      headerDistance: section.headerDistance,
      footerDistance: section.footerDistance,
    },
    columns: computeSectionColumns(section),
    grid: {
      kind: normalizeGridKind(section.docGridType),
      linePitchPt: section.docGridLinePitch ?? null,
      charSpacePt:
        section.docGridCharSpace == null ? null : section.docGridCharSpace / 4096,
    },
    textDirection: section.textDirection ?? 'lrTb',
    verticalAlignment: section.vAlign ?? 'top',
    lineNumbering: section.lineNumbering ?? undefined,
  };
}

/** Temporary bridge for call sites that still consume the legacy grid shape. */
export function toLegacyDocGridContext(
  section: SectionLayoutContext,
): DocGridCtx {
  return {
    type: section.grid.kind === 'none' ? null : section.grid.kind,
    linePitchPt: section.grid.linePitchPt,
    charSpacePt: section.grid.charSpacePt,
  };
}

function hasTableCellContainer(story: StoryContext): boolean {
  return story.containers.some((container) => container.kind === 'tableCell');
}

export function resolveParagraphLayoutContext(
  settings: DocumentLayoutSettings,
  section: SectionLayoutContext,
  story: StoryContext,
  paragraph: DocParagraph,
): ParagraphLayoutContext {
  const lineGridActive =
    isSectionLineGrid(section.grid.kind)
    && section.grid.linePitchPt != null
    && section.grid.linePitchPt > 0
    && paragraph.snapToGrid !== false
    && paragraph.lineSpacing?.rule !== 'exact'
    && (!hasTableCellContainer(story) || settings.compat.adjustLineHeightInTable);
  const characterGridActive =
    isSectionCharacterGrid(section.grid.kind)
    && section.grid.charSpacePt != null;
  const baseRtl = paragraph.bidi === true;

  // ECMA-376 §17.9.28 (`<w:suff>`, default "tab") + §17.3.1.6 (`<w:ind>` is logical
  // under `<w:bidi>`): a suff=tab numbering marker with a HANGING first-line indent
  // advances the first-line body to the indentLeft tab stop, so the first line's
  // effective indent is 0 — the marker occupies the hanging margin, mirrored to the
  // physical-right start edge in an RTL paragraph. The PAINT pass positions the RTL
  // suff=tab body with that marker-aware indent (renderer.ts `markerUsesBodyOffset` →
  // `numBodyOffset`, which is 0 whenever the marker fits the hanging indent), so the
  // MEASURE/paginate pass must use the same effective first-line width or the two
  // disagree on line count for a paragraph split across pages (the paginator's
  // `lineSlice` indices would then reference a different partition). Align them for
  // the RTL hanging suff=tab case here.
  //
  // Scope matches the paint gate exactly: RTL only, suff=tab only, and a genuine
  // hanging indent (`indentFirst < 0`). LTR keeps raw `indentFirst` (byte-identical
  // pagination; its long-standing measure/paint marker approximation is untouched);
  // a non-hanging or suff=space/nothing marker also keeps raw `indentFirst`, staying
  // consistent with paint. numBodyOffset needs marker font metrics (only resolved in
  // the renderer), so the rare suff=tab marker-OVERRUN sub-case — where a marker
  // wider than the hanging indent advances the body PAST the tab stop (§17.3.1.37) —
  // keeps a small bounded residual here; 0 still matches the common case exactly and
  // is far closer than the raw −hanging it replaces.
  const numbering = paragraph.numbering;
  const hasNumberingMarker =
    numbering != null && (numbering.text !== '' || numbering.picBulletImagePath != null);
  const rtlHangingSuffTabMarker =
    baseRtl
    && hasNumberingMarker
    && (numbering!.suff || 'tab') === 'tab'
    && paragraph.indentFirst < 0;

  return {
    lineGrid: {
      active: lineGridActive,
      pitchPt: lineGridActive ? section.grid.linePitchPt : null,
    },
    characterGrid: {
      active: characterGridActive,
      deltaPt: characterGridActive ? section.grid.charSpacePt ?? 0 : 0,
    },
    physicalIndentLeftPt: baseRtl ? paragraph.indentRight : paragraph.indentLeft,
    physicalIndentRightPt: baseRtl ? paragraph.indentLeft : paragraph.indentRight,
    firstIndentPt: rtlHangingSuffTabMarker ? 0 : paragraph.indentFirst,
    lineSpacing: paragraph.lineSpacing,
    spaceBeforePt: paragraph.spaceBefore,
    spaceAfterPt: paragraph.spaceAfter,
    baseRtl,
    isJustified: jcIsFullyJustified(paragraph.alignment),
    stretchLastLine: jcStretchesLastLine(paragraph.alignment),
    tabStops: [...paragraph.tabStops],
    hasRuby: paragraphHasRuby(paragraph),
    hasEastAsianText: paragraphHasEastAsianText(paragraph),
    kinsoku: settings.kinsoku,
    defaultTabPt: settings.defaultTabPt,
  };
}

export function resolveRunLayoutContext(
  paragraph: ParagraphLayoutContext,
  run: DocxTextRun,
): RunLayoutContext {
  const active = paragraph.characterGrid.active && run.snapToGrid !== false;
  return {
    characterGrid: {
      active,
      deltaPt: active ? paragraph.characterGrid.deltaPt : 0,
    },
  };
}
