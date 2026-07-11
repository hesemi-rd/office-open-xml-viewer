import { describe, expect, it } from 'vitest';
import { buildSegments, type LayoutTextSeg } from './line-layout.js';
import type { DocRun } from './types.js';

const ENV = { pageIndex: 0, totalPages: 1 };

function textRun(text: string): DocRun {
  return {
    type: 'text',
    text,
    bold: false,
    italic: false,
    underline: false,
    strikethrough: false,
    fontSize: 12,
    color: null,
    fontFamily: 'serif',
    isLink: false,
    background: null,
    vertAlign: null,
    allCaps: false,
    smallCaps: false,
    doubleStrikethrough: false,
  } as unknown as DocRun;
}

function textSegments(runs: DocRun[]): LayoutTextSeg[] {
  return buildSegments(runs, ENV).filter((seg): seg is LayoutTextSeg => 'text' in seg);
}

describe('buildSegments UAX #14 LB28 boundary glue', () => {
  it.each([
    ['AL × AL', '<', 'a'],
    ['AL × HL', '<', 'א'],
    ['HL × AL', 'א', 'a'],
    ['HL × HL', 'א', 'ב'],
  ])('marks the following segment joinPrev for %s', (_label, prev, next) => {
    expect(textSegments([textRun(prev), textRun(next)]).map((seg) => seg.joinPrev))
      .toEqual([undefined, true]);
  });

  it.each([
    ['trailing whitespace', [textRun('< '), textRun('a')]],
    ['leading whitespace', [textRun('<'), textRun(' a')]],
    ['zero-width space', [textRun('<\u200b'), textRun('a')]],
    ['CJK-breakable text', [textRun('<'), textRun('漢')]],
    ['SEA dictionary text', [textRun('<'), textRun('ก')]],
    [
      'non-text boundary',
      [
        textRun('<'),
        { type: 'break', breakType: 'line' } as DocRun,
        textRun('a'),
      ],
    ],
  ])('does not add glue across %s', (_label, runs) => {
    expect(textSegments(runs).every((seg) => seg.joinPrev === undefined)).toBe(true);
  });
});
