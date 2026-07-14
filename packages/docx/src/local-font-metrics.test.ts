import { describe, expect, it } from 'vitest';
import { docxLocalMetricRequests } from './local-font-metrics.js';
import { buildSegments, segmentIntendedSingleLinePx } from './line-layout.js';
import type { DocRun, DocxDocumentModel, DocxTextRun } from './types.js';

function model(fonts: string[]): DocxDocumentModel {
  return {
    fontFamilyClasses: Object.fromEntries(fonts.map((font) => [font, 'modern'])),
    body: [{
      type: 'paragraph',
      runs: fonts.map((font) => ({ type: 'text', text: 'x', fontFamily: font })),
    }],
  } as unknown as DocxDocumentModel;
}

function paragraphWithFamily(family: string): unknown {
  return {
    type: 'paragraph',
    runs: [{ type: 'text', text: '平成', fontFamily: family }],
  };
}

describe('docxLocalMetricRequests', () => {
  it('snapshots every authored family and requested weight/style tuple', () => {
    const doc = {
      body: [{
        type: 'paragraph',
        runs: [
          { type: 'text', text: 'plain', fontFamily: 'Times New Roman', bold: false, italic: false },
          { type: 'text', text: 'bold italic', fontFamily: 'Times New Roman', bold: true, italic: true },
          { type: 'text', text: '国', fontFamily: 'Arial', fontFamilyEastAsia: 'Yu Mincho', bold: true, italic: false },
        ],
      }],
    } as unknown as DocxDocumentModel;

    expect(docxLocalMetricRequests(doc)).toEqual([
      { family: 'Times New Roman', localNames: ['Times New Roman'] },
      { family: 'Times New Roman', localNames: ['Times New Roman'], weight: 700, style: 'italic' },
      { family: 'Arial', localNames: ['Arial'], weight: 700 },
      { family: 'Yu Mincho', localNames: ['Yu Mincho'], weight: 700 },
    ]);
  });

  it('inventories Latin/eastAsia and complex-script axes with their independent style tuples', () => {
    const doc = {
      body: [{
        type: 'paragraph',
        runs: [
          {
            type: 'text', text: 'mixed', fontFamily: 'Latin Regular',
            fontFamilyEastAsia: 'EA Regular', fontFamilyCs: 'CS Bold',
            bold: false, italic: false, boldCs: true, italicCs: false,
          },
          {
            type: 'text', text: 'inverse', fontFamily: 'Latin Bold',
            fontFamilyEastAsia: 'EA Bold', fontFamilyCs: 'CS Regular',
            bold: true, italic: false, boldCs: false, italicCs: false,
          },
        ],
      }],
    } as unknown as DocxDocumentModel;

    expect(docxLocalMetricRequests(doc)).toEqual([
      { family: 'Latin Regular', localNames: ['Latin Regular'] },
      { family: 'EA Regular', localNames: ['EA Regular'] },
      { family: 'CS Bold', localNames: ['CS Bold'], weight: 700 },
      { family: 'Latin Bold', localNames: ['Latin Bold'], weight: 700 },
      { family: 'EA Bold', localNames: ['EA Bold'], weight: 700 },
      { family: 'CS Regular', localNames: ['CS Regular'] },
    ]);
  });
  it('maps Japanese and English Meiryo names to the exact local family', () => {
    expect(docxLocalMetricRequests(model(['メイリオ', 'Meiryo']))).toEqual([
      { family: 'メイリオ', localNames: ['Meiryo'], lineHeightMultiplier: 1.3 },
      { family: 'Meiryo', localNames: ['Meiryo'], lineHeightMultiplier: 1.3 },
    ]);
  });

  it('probes Meiryo UI under its own isolated alias rather than merging it into Meiryo', () => {
    expect(docxLocalMetricRequests(model(['Meiryo UI']))).toEqual([
      { family: 'Meiryo UI', localNames: ['Meiryo UI'] },
    ]);
  });

  it('keeps an embedded regular face authoritative over terminal-local probing and aliases', () => {
    const doc = {
      ...model(['メイリオ']),
      embeddedFonts: [{
        fontName: 'メイリオ', style: 'regular',
        partPath: 'word/fonts/font1.odttf', fontKey: '{00000000-0000-0000-0000-000000000000}',
      }],
      body: [paragraphWithFamily('メイリオ')],
    } as unknown as DocxDocumentModel;

    const requests = docxLocalMetricRequests(doc);
    expect(requests).toEqual([]);
    // Model the runtime hand-off from selected local-metric requests to the
    // segment resolver. If the embedded family leaked into `requests`, its
    // terminal alias would replace the embedded face during normal-run layout.
    const resolvedLocalFonts = Object.fromEntries(requests.map(({ family }) => [
      family.trim().toLowerCase(),
      { family: '__terminal_local_alias', lineHeightRatio: 1.5 },
    ]));
    const [segment] = buildSegments(
      (doc.body[0] as { runs: DocRun[] }).runs,
      { pageIndex: 0, totalPages: 1, resolvedLocalFonts },
    );
    expect('text' in segment && segment.fontFamily).toBe('メイリオ');
  });

  it('still probes a normal face when only a non-regular embedded face exists', () => {
    const doc = {
      ...model(['メイリオ']),
      embeddedFonts: [{
        fontName: 'メイリオ', style: 'bold',
        partPath: 'word/fonts/font1.odttf', fontKey: '{00000000-0000-0000-0000-000000000000}',
      }],
    } as unknown as DocxDocumentModel;

    expect(docxLocalMetricRequests(doc)).toEqual([
      { family: 'メイリオ', localNames: ['Meiryo'], lineHeightMultiplier: 1.3 },
    ]);
  });

  it('discovers a directly-authored run family absent from fontTable and the theme', () => {
    const doc = {
      body: [paragraphWithFamily('メイリオ')],
    } as unknown as DocxDocumentModel;

    expect(docxLocalMetricRequests(doc)).toEqual([
      { family: 'メイリオ', localNames: ['Meiryo'], lineHeightMultiplier: 1.3 },
    ]);
  });

  it('discovers the direct complex-script rFonts axis', () => {
    const doc = {
      body: [{
        type: 'paragraph',
        runs: [{
          type: 'text', text: 'العربية', fontFamily: 'Arial',
          fontFamilyEastAsia: 'Arial', fontFamilyCs: 'Meiryo',
        }],
      }],
    } as unknown as DocxDocumentModel;

    expect(docxLocalMetricRequests(doc)).toEqual([
      { family: 'Arial', localNames: ['Arial'] },
      { family: 'Meiryo', localNames: ['Meiryo'], lineHeightMultiplier: 1.3 },
    ]);
  });

  it.each([
    {
      story: 'header',
      document: { headers: { default: { body: [paragraphWithFamily('Meiryo')] }, first: null, even: null } },
    },
    {
      story: 'footer',
      document: { footers: { default: { body: [paragraphWithFamily('Meiryo')] }, first: null, even: null } },
    },
    {
      story: 'earlier-section header',
      document: {
        body: [{
          type: 'sectionBreak',
          headers: { default: { body: [paragraphWithFamily('Meiryo')] }, first: null, even: null },
        }],
      },
    },
    {
      story: 'footnote',
      document: { footnotes: [{ id: '1', content: [paragraphWithFamily('Meiryo')] }] },
    },
    {
      story: 'endnote',
      document: { endnotes: [{ id: '1', content: [paragraphWithFamily('Meiryo')] }] },
    },
    {
      story: 'shape text',
      document: {
        body: [{
          type: 'paragraph',
          runs: [{
            type: 'shape',
            textBlocks: [{
              text: '平成', fontSizePt: 10, alignment: 'left', fontFamily: 'Meiryo',
            }],
          }],
        }],
      },
    },
  ])('discovers a family used only in $story', ({ document }) => {
    expect(docxLocalMetricRequests(document as unknown as DocxDocumentModel)).toEqual([
      { family: 'Meiryo', localNames: ['Meiryo'], lineHeightMultiplier: 1.3 },
    ]);
  });

  it('routes layout through the exact local alias and its measured Word line height', () => {
    const run = {
      type: 'text', text: '平成', bold: false, italic: false, underline: false,
      strikethrough: false, fontSize: 10, color: null, fontFamily: 'メイリオ',
      fontFamilyEastAsia: 'メイリオ', vertAlign: null,
    } as unknown as DocxTextRun;
    const segments = buildSegments([run as unknown as DocRun], {
      pageIndex: 0,
      totalPages: 1,
      resolvedLocalFonts: {
        'メイリオ': { family: '__ooxml_local_meiryo', lineHeightRatio: 1.95 },
      },
    });
    const segment = segments[0];
    expect('text' in segment && segment.fontFamily).toBe('__ooxml_local_meiryo');
    expect('text' in segment && segmentIntendedSingleLinePx(segment, 10, true)).toBeCloseTo(19.5, 8);
  });

  it('uses an ordinary local alias without inventing a line-height override', () => {
    const run = {
      type: 'text', text: 'Latin', bold: false, italic: false, underline: false,
      strikethrough: false, fontSize: 10, color: null, fontFamily: 'Times New Roman',
      vertAlign: null,
    } as unknown as DocxTextRun;
    const [segment] = buildSegments([run as unknown as DocRun], {
      pageIndex: 0,
      totalPages: 1,
      resolvedLocalFonts: {
        'times new roman': {
          family: '__ooxml_local_times', requestedFamily: 'Times New Roman',
          weight: 400, style: 'normal',
        },
      },
    });

    expect(segment).toMatchObject({ fontFamily: '__ooxml_local_times' });
    expect('text' in segment && segment.resolvedLineHeightRatio).toBeUndefined();
  });

  it.each([
    { bold: true, italic: false },
    { bold: false, italic: true },
    { bold: true, italic: true },
  ])('keeps the authored family for styled faces while retaining measured line height (%o)', (style) => {
    const run = {
      type: 'text', text: '平成', ...style, underline: false,
      strikethrough: false, fontSize: 10, color: null, fontFamily: 'メイリオ',
      fontFamilyEastAsia: 'メイリオ', vertAlign: null,
    } as unknown as DocxTextRun;
    const [segment] = buildSegments([run as unknown as DocRun], {
      pageIndex: 0,
      totalPages: 1,
      resolvedLocalFonts: {
        'メイリオ': { family: '__ooxml_local_meiryo', lineHeightRatio: 1.95 },
      },
    });

    expect('text' in segment && segment.fontFamily).toBe('メイリオ');
    expect('text' in segment && segmentIntendedSingleLinePx(segment, 10, true)).toBeCloseTo(19.5, 8);
  });

  it.each([
    { boldCs: true, italicCs: false },
    { boldCs: false, italicCs: true },
  ])('keeps the authored complex-script family for styled faces (%o)', (style) => {
    const run = {
      type: 'text', text: 'العربية', bold: false, italic: false, underline: false,
      strikethrough: false, fontSize: 10, color: null, fontFamily: 'Arial',
      fontFamilyEastAsia: 'Arial', fontFamilyCs: 'メイリオ', fontSizeCs: 10,
      cs: true, ...style, vertAlign: null,
    } as unknown as DocxTextRun;
    const [segment] = buildSegments([run as unknown as DocRun], {
      pageIndex: 0,
      totalPages: 1,
      resolvedLocalFonts: {
        'メイリオ': { family: '__ooxml_local_meiryo', lineHeightRatio: 1.95 },
      },
    });

    expect('text' in segment && segment.fontFamily).toBe('メイリオ');
    expect('text' in segment && segmentIntendedSingleLinePx(segment, 10, true)).toBeCloseTo(19.5, 8);
  });

  it('uses the exact normal alias when complex-script styling is explicitly off', () => {
    const run = {
      type: 'text', text: 'العربية', bold: true, italic: true, underline: false,
      strikethrough: false, fontSize: 10, color: null, fontFamily: 'Arial',
      fontFamilyEastAsia: 'Arial', fontFamilyCs: 'メイリオ', fontSizeCs: 10,
      cs: true, boldCs: false, italicCs: false, vertAlign: null,
    } as unknown as DocxTextRun;
    const [segment] = buildSegments([run as unknown as DocRun], {
      pageIndex: 0,
      totalPages: 1,
      resolvedLocalFonts: {
        'メイリオ': { family: '__ooxml_local_meiryo', lineHeightRatio: 1.95 },
      },
    });

    expect('text' in segment && segment.fontFamily).toBe('__ooxml_local_meiryo');
  });

  it.each([
    { bold: false, italic: false, expectedFamily: '__ooxml_local_meiryo' },
    { bold: true, italic: false, expectedFamily: 'メイリオ' },
    { bold: false, italic: true, expectedFamily: 'メイリオ' },
  ])('applies the normal-vs-styled alias policy to anchor-host metrics (%o)', (style) => {
    const [segment] = buildSegments([{
      type: 'anchorHost', fontSize: 10, fontFamily: 'メイリオ',
      fontFamilyEastAsia: 'メイリオ', bold: style.bold, italic: style.italic,
    }], {
      pageIndex: 0,
      totalPages: 1,
      resolvedLocalFonts: {
        'メイリオ': { family: '__ooxml_local_meiryo', lineHeightRatio: 1.95 },
      },
    });

    expect('text' in segment && segment.fontFamily).toBe(style.expectedFamily);
    expect('text' in segment && segmentIntendedSingleLinePx(segment, 10, true)).toBeCloseTo(19.5, 8);
  });
});
