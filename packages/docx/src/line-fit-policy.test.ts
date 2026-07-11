import { describe, expect, it } from 'vitest';
import { adjustForWidowOrphan, selectLargestFittingEnd } from './line-fit-policy.js';

describe('selectLargestFittingEnd — greedy break selection (layout convention; §17.4.6 gates live in the callers)', () => {
  it('returns an empty selection when the candidate range is empty', () => {
    expect(selectLargestFittingEnd(3, 3, 100, () => 1)).toEqual({ end: 3, fitValue: 0 });
  });

  it('includes an end whose height exactly equals the available height', () => {
    expect(selectLargestFittingEnd(1, 3, 5, (end) => end === 2 ? 5 : 8)).toEqual({ end: 2, fitValue: 5 });
  });

  it('returns end === start when the first line does not fit', () => {
    expect(selectLargestFittingEnd(2, 4, 6, () => 7)).toEqual({ end: 2, fitValue: 0 });
  });

  it('stops at the first overflow even if a later end would fit', () => {
    const visited: number[] = [];
    const heights = new Map([[1, 4], [2, 7], [3, 5]]);

    const selection = selectLargestFittingEnd(0, 3, 5, (end) => {
      visited.push(end);
      return heights.get(end) as number;
    });

    expect(selection).toEqual({ end: 1, fitValue: 4 });
    expect(visited).toEqual([1, 2]);
  });
});

describe('§17.3.1.44 widowControl — adjustForWidowOrphan', () => {
  const input = {
    widowControl: true,
    start: 0,
    end: 3,
    totalLines: 4,
    canRelocate: false,
  };

  it('drops the last selected line only for a one-line remainder after keeping at least two lines', () => {
    expect(adjustForWidowOrphan(input)).toEqual({ kind: 'dropLastLine' });
    expect(adjustForWidowOrphan({ ...input, end: 2 })).toEqual({ kind: 'keep' });
    expect(adjustForWidowOrphan({ ...input, start: 2 })).toEqual({ kind: 'keep' });
  });

  it('relocates only a lone first line selected below the column top', () => {
    const orphan = { ...input, end: 1, canRelocate: true };

    expect(adjustForWidowOrphan(orphan)).toEqual({ kind: 'relocate' });
    expect(adjustForWidowOrphan({ ...orphan, start: 1, end: 2 })).toEqual({ kind: 'keep' });
    expect(adjustForWidowOrphan({ ...orphan, end: 2 })).toEqual({ kind: 'keep' });
    expect(adjustForWidowOrphan({ ...orphan, canRelocate: false })).toEqual({ kind: 'keep' });
  });

  it('keeps the greedy selection when widow control is disabled', () => {
    expect(adjustForWidowOrphan({ ...input, widowControl: false })).toEqual({ kind: 'keep' });
    expect(adjustForWidowOrphan({
      ...input,
      widowControl: false,
      end: 1,
      canRelocate: true,
    })).toEqual({ kind: 'keep' });
  });
});
