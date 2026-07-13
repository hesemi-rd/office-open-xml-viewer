import { describe, expect, it } from 'vitest';
import { docxLocalMetricRequests } from './local-font-metrics.js';
import { buildSegments, segmentIntendedSingleLinePx } from './line-layout.js';
import type { DocRun, DocxDocumentModel, DocxTextRun } from './types.js';

function model(fonts: string[]): DocxDocumentModel {
  return {
    fontFamilyClasses: Object.fromEntries(fonts.map((font) => [font, 'modern'])),
  } as unknown as DocxDocumentModel;
}

describe('docxLocalMetricRequests', () => {
  it('maps Japanese and English Meiryo names to the exact local family', () => {
    expect(docxLocalMetricRequests(model(['メイリオ', 'Meiryo']))).toEqual([
      { family: 'メイリオ', localNames: ['Meiryo'], lineHeightMultiplier: 1.3 },
      { family: 'Meiryo', localNames: ['Meiryo'], lineHeightMultiplier: 1.3 },
    ]);
  });

  it('does not merge Meiryo UI into Meiryo because their design metrics differ', () => {
    expect(docxLocalMetricRequests(model(['Meiryo UI']))).toEqual([]);
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
});
