import { describe, expect, it } from 'vitest';
import type {
  DocParagraph,
  DocxDocumentModel,
  DocxTextRun,
  SectionProps,
} from './types.js';
import {
  resolveDocumentLayoutSettings,
  resolveParagraphLayoutContext,
  resolveRunLayoutContext,
  resolveSectionLayoutContext,
  toLegacyDocGridContext,
  type StoryContext,
} from './layout-context.js';

const section = (overrides: Partial<SectionProps> = {}): SectionProps => ({
  pageWidth: 200,
  pageHeight: 300,
  marginTop: 20,
  marginRight: 20,
  marginBottom: 20,
  marginLeft: 20,
  headerDistance: 10,
  footerDistance: 10,
  titlePage: false,
  evenAndOddHeaders: false,
  ...overrides,
});

const paragraph = (overrides: Partial<DocParagraph> = {}): DocParagraph => ({
  alignment: 'left',
  indentLeft: 12,
  indentRight: 6,
  indentFirst: 2,
  spaceBefore: 3,
  spaceAfter: 4,
  lineSpacing: null,
  numbering: null,
  tabStops: [],
  runs: [],
  ...overrides,
});

const textRun = (overrides: Partial<DocxTextRun> = {}): DocxTextRun => ({
  text: 'x',
  bold: false,
  italic: false,
  underline: false,
  strikethrough: false,
  fontSize: 10,
  color: null,
  fontFamily: null,
  isLink: false,
  background: null,
  vertAlign: null,
  hyperlink: null,
  ...overrides,
});

const documentModel = (
  overrides: Partial<DocxDocumentModel> = {},
): DocxDocumentModel => ({
  section: section(),
  body: [],
  headers: { default: null, first: null, even: null },
  footers: { default: null, first: null, even: null },
  ...overrides,
});

const bodyStory: StoryContext = {
  story: 'body',
  containers: [],
  lineNumberingEligible: true,
};

const cellStory: StoryContext = {
  story: 'body',
  containers: [{ kind: 'tableCell' }],
  lineNumberingEligible: false,
};

describe('layout context resolvers', () => {
  it('normalizes document settings and detects East Asian body text once', () => {
    const run = textRun({ text: 'あ' });
    const settings = resolveDocumentLayoutSettings(documentModel({
      body: [{ type: 'paragraph', ...paragraph({ runs: [{ type: 'text', ...run }] }) }],
      settings: {
        kinsoku: false,
        defaultTabStop: 18,
        characterSpacingControl: 'compressPunctuation',
        mathDefJc: 'left',
        adjustLineHeightInTable: true,
        useFeLayout: true,
        balanceSingleByteDoubleByteWidth: true,
      },
    }));

    expect(settings.defaultTabPt).toBe(18);
    expect(settings.kinsoku.enabled).toBe(false);
    expect(settings.documentHasEastAsianText).toBe(true);
    expect(settings.characterSpacingControl).toBe('compressPunctuation');
    expect(settings.mathDefJc).toBe('left');
    expect(settings.compat).toEqual({
      adjustLineHeightInTable: true,
      useFeLayout: true,
      balanceSingleByteDoubleByteWidth: true,
    });
  });

  it('resolves absent settings and default docGrid without activating a grid', () => {
    const settings = resolveDocumentLayoutSettings(documentModel());
    const context = resolveSectionLayoutContext(
      settings,
      section({ docGridType: 'default', docGridLinePitch: 20 }),
    );

    expect(settings.defaultTabPt).toBe(36);
    expect(settings.compat).toEqual({
      adjustLineHeightInTable: false,
      useFeLayout: false,
      balanceSingleByteDoubleByteWidth: false,
    });
    expect(context.grid).toEqual({
      kind: 'none',
      linePitchPt: 20,
      charSpacePt: null,
    });
  });

  it('normalizes section geometry, columns, and snapToChars grid units', () => {
    const settings = resolveDocumentLayoutSettings(documentModel());
    const context = resolveSectionLayoutContext(settings, section({
      docGridType: 'snapToChars',
      docGridLinePitch: 20,
      docGridCharSpace: 4096,
      columns: {
        count: 2,
        spacePt: 10,
        equalWidth: true,
        sep: false,
        cols: [],
      },
    }));

    expect(context.geometry).toEqual({
      pageWidth: 200,
      pageHeight: 300,
      marginTop: 20,
      marginRight: 20,
      marginBottom: 20,
      marginLeft: 20,
      headerDistance: 10,
      footerDistance: 10,
    });
    expect(context.columns).toEqual([
      { xPt: 20, wPt: 75 },
      { xPt: 105, wPt: 75 },
    ]);
    expect(context.grid).toEqual({
      kind: 'snapToChars',
      linePitchPt: 20,
      charSpacePt: 1,
    });
    expect(toLegacyDocGridContext(context)).toEqual({
      type: 'snapToChars',
      linePitchPt: 20,
      charSpacePt: 1,
    });
  });

  it('keeps invalid line pitches inactive', () => {
    const settings = resolveDocumentLayoutSettings(documentModel());
    const sectionContext = resolveSectionLayoutContext(
      settings,
      section({ docGridType: 'lines', docGridLinePitch: Number.NaN }),
    );
    const paragraphContext = resolveParagraphLayoutContext(
      settings,
      sectionContext,
      bodyStory,
      paragraph(),
    );

    expect(paragraphContext.lineGrid).toEqual({ active: false, pitchPt: null });
  });

  it('resolves physical bidi indents without changing authored spacing', () => {
    const settings = resolveDocumentLayoutSettings(documentModel());
    const sectionContext = resolveSectionLayoutContext(
      settings,
      section({ docGridType: 'lines', docGridLinePitch: 20 }),
    );
    const context = resolveParagraphLayoutContext(
      settings,
      sectionContext,
      bodyStory,
      paragraph({ bidi: true }),
    );

    expect(context.physicalIndentLeftPt).toBe(6);
    expect(context.physicalIndentRightPt).toBe(12);
    expect(context.firstIndentPt).toBe(2);
    expect(context.spaceBeforePt).toBe(3);
    expect(context.spaceAfterPt).toBe(4);
    expect(context.baseRtl).toBe(true);
    expect(context.lineGrid).toEqual({ active: true, pitchPt: 20 });
  });

  it('gates only line-grid policy for paragraphs and table cells', () => {
    const noCompat = resolveDocumentLayoutSettings(documentModel());
    const withCompat = resolveDocumentLayoutSettings(documentModel({
      settings: { adjustLineHeightInTable: true },
    }));
    const sectionContext = resolveSectionLayoutContext(
      noCompat,
      section({
        docGridType: 'linesAndChars',
        docGridLinePitch: 20,
        docGridCharSpace: 2048,
      }),
    );

    const optedOut = resolveParagraphLayoutContext(
      noCompat,
      sectionContext,
      bodyStory,
      paragraph({ snapToGrid: false }),
    );
    expect(optedOut.lineGrid.active).toBe(false);
    expect(optedOut.characterGrid.active).toBe(true);

    const exact = resolveParagraphLayoutContext(
      noCompat,
      sectionContext,
      bodyStory,
      paragraph({ lineSpacing: { value: 18, rule: 'exact', explicit: true } }),
    );
    expect(exact.lineGrid.active).toBe(false);
    expect(exact.characterGrid.active).toBe(true);

    const cellWithoutCompat = resolveParagraphLayoutContext(
      noCompat,
      sectionContext,
      cellStory,
      paragraph(),
    );
    expect(cellWithoutCompat.lineGrid.active).toBe(false);
    expect(cellWithoutCompat.characterGrid.active).toBe(true);

    const compatibleSection = resolveSectionLayoutContext(
      withCompat,
      section({
        docGridType: 'linesAndChars',
        docGridLinePitch: 20,
        docGridCharSpace: 2048,
      }),
    );
    const cellWithCompat = resolveParagraphLayoutContext(
      withCompat,
      compatibleSection,
      cellStory,
      paragraph(),
    );
    expect(cellWithCompat.lineGrid.active).toBe(true);
  });

  it('lets a run opt out of character-grid spacing without changing line policy', () => {
    const settings = resolveDocumentLayoutSettings(documentModel());
    const sectionContext = resolveSectionLayoutContext(
      settings,
      section({
        docGridType: 'snapToChars',
        docGridLinePitch: 20,
        docGridCharSpace: 4096,
      }),
    );
    const paragraphContext = resolveParagraphLayoutContext(
      settings,
      sectionContext,
      bodyStory,
      paragraph(),
    );

    expect(paragraphContext.lineGrid.active).toBe(true);
    expect(resolveRunLayoutContext(
      paragraphContext,
      textRun({ snapToGrid: false }),
    ).characterGrid).toEqual({ active: false, deltaPt: 0 });
  });
});
