import { describe, expect, it } from 'vitest';
import {
  convergePaginationFields,
  paginatedFlowHasPaginationDependentFields,
  paginationFieldFlowGeometry,
  paginationFieldGeometryFingerprint,
  resolvePaginationFieldLayout,
} from './layout/pagination-fields.js';
import type { BodyElement, DocParagraph, DocxDocumentModel, FieldRun } from './types.js';

function paragraphWithRuns(
  runs: DocParagraph['runs'],
): Extract<BodyElement, { type: 'paragraph' }> {
  return {
    type: 'paragraph',
    alignment: 'left', indentLeft: 0, indentRight: 0, indentFirst: 0,
    spaceBefore: 0, spaceAfter: 0, lineSpacing: null, numbering: null,
    tabStops: [], runs,
  } as Extract<BodyElement, { type: 'paragraph' }>;
}

function field(fieldType: 'page' | 'numPages'): FieldRun & { type: 'field' } {
  return {
    type: 'field', fieldType,
    instruction: fieldType === 'page' ? 'PAGE' : 'NUMPAGES', fallbackText: '?',
    bold: false, italic: false, underline: false, strikethrough: false,
    fontSize: 10, color: null, fontFamily: null, background: null, vertAlign: null,
  };
}

function documentWith(body: BodyElement[]): DocxDocumentModel {
  return {
    body,
    section: {} as DocxDocumentModel['section'],
    headers: { default: null, first: null, even: null },
    footers: { default: null, first: null, even: null },
  };
}

describe('pagination field convergence seam', () => {
  it('acquires field-independent pagination exactly once', () => {
    const hints: number[] = [];
    const result = resolvePaginationFieldLayout((hint) => {
      hints.push(hint);
      return { fingerprint: 'pages:2', pageCount: 2 };
    }, false);

    expect(hints).toEqual([1]);
    expect(result.pageCount).toBe(2);
  });

  it('retains convergence when pagination fields can change geometry', () => {
    const hints: number[] = [];
    resolvePaginationFieldLayout((hint) => {
      hints.push(hint);
      return { fingerprint: 'pages:2', pageCount: 2 };
    }, true);

    expect(hints).toEqual([1, 2]);
  });

  it('finds pagination fields in body, nested table, and footnote acquisition inputs', () => {
    const fieldFree = documentWith([paragraphWithRuns([])]);
    expect(paginatedFlowHasPaginationDependentFields(fieldFree.body)).toBe(false);

    const nestedTable = documentWith([{
      type: 'table',
      rows: [{ cells: [{ content: [paragraphWithRuns([field('numPages')])] }] }],
    } as unknown as BodyElement]);
    expect(paginatedFlowHasPaginationDependentFields(nestedTable.body)).toBe(true);

    const bodyPage = documentWith([paragraphWithRuns([field('page')])]);
    expect(paginatedFlowHasPaginationDependentFields(bodyPage.body)).toBe(true);

    const sectionHeader = documentWith([{
      type: 'sectionBreak', kind: 'nextPage',
      headers: {
        default: { body: [paragraphWithRuns([field('numPages')])] },
        first: null, even: null,
      },
    }]);
    expect(paginatedFlowHasPaginationDependentFields(sectionHeader.body)).toBe(false);

    const footnote = documentWith([paragraphWithRuns([])]);
    footnote.footnotes = [{ id: '1', content: [paragraphWithRuns([field('numPages')])] }];
    expect(paginatedFlowHasPaginationDependentFields(footnote.body, footnote.footnotes)).toBe(true);

    const pageOnlyFootnote = documentWith([paragraphWithRuns([])]);
    pageOnlyFootnote.footnotes = [{ id: '1', content: [paragraphWithRuns([field('page')])] }];
    expect(paginatedFlowHasPaginationDependentFields(
      pageOnlyFootnote.body,
      pageOnlyFootnote.footnotes,
    )).toBe(true);
  });

  it('normalizes absent optional runtime placement facts before fingerprinting', () => {
    const omitted = paginationFieldGeometryFingerprint({
      pageCount: 1,
      pages: [[{ type: 'paragraph', colIndex: 0 }]],
    });
    const explicitUndefined = paginationFieldGeometryFingerprint({
      pageCount: 1,
      pages: [[{ type: 'paragraph', colIndex: 0, colY: undefined, placed: undefined }]],
    });

    expect(explicitUndefined).toBe(omitted);
  });

  it('projects paragraph geometry without parser/source objects', () => {
    const geometry = paginationFieldFlowGeometry({
      kind: 'paragraph', id: 'body:0',
      source: { story: 'body', storyInstance: 'body', path: [0] },
      flowDomainId: 'body', ordinaryFlow: true,
      flowBounds: { xPt: 10, yPt: 20, widthPt: 100, heightPt: 12 },
      inkBounds: { xPt: 10, yPt: 20, widthPt: 40, heightPt: 10 },
      advancePt: 12, spacing: { beforePt: 0, afterPt: 2 }, contextualSpacing: false,
      lines: [], borders: [], resources: [], drawings: [], textBoxes: [], events: [], exclusions: [],
    });

    expect(geometry).toMatchObject({
      kind: 'paragraph',
      flowBounds: { xPt: 10, yPt: 20, widthPt: 100, heightPt: 12 },
      advancePt: 12,
    });
    expect(JSON.stringify(geometry)).not.toContain('source');
    expect(JSON.stringify(geometry)).not.toContain('storyInstance');
  });

  it('stabilizes on the geometry acquired with the resolved page count', () => {
    const hints: number[] = [];
    const result = convergePaginationFields((hint) => {
      hints.push(hint);
      const pageCount = hint === 1 ? 2 : 2;
      return { fingerprint: `pages:${pageCount}`, pageCount };
    });

    expect(hints).toEqual([1, 2]);
    expect(result).toEqual({ fingerprint: 'pages:2', pageCount: 2 });
  });

  it('hard-fails a repeated geometry cycle', () => {
    let step = 0;
    expect(() => convergePaginationFields(() => {
      const current = step++ % 2 === 0
        ? { fingerprint: 'geometry:a', pageCount: 2 }
        : { fingerprint: 'geometry:b', pageCount: 1 };
      return current;
    })).toThrow(/repeated geometry fingerprint cycle/i);
  });

  it('hard-fails when hostile geometry never stabilizes within the policy limit', () => {
    let step = 0;
    expect(() => convergePaginationFields(
      () => ({ fingerprint: `geometry:${step++}`, pageCount: step + 1 }),
      3,
    )).toThrow(/hard iteration limit 3 reached/i);
  });
});
