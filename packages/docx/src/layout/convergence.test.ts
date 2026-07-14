import { describe, expect, it } from 'vitest';
import { convergeLayout, type LayoutIteration } from './convergence.js';

const iteration = (fingerprint: string, pageCount: number): LayoutIteration => ({
  fingerprint,
  pageCount,
});

describe('convergeLayout', () => {
  it('returns the newest iteration when its relevant geometry fingerprint stabilizes', () => {
    const seed = iteration('a', 1);
    const calls: string[] = [];
    const result = convergeLayout(seed, (current) => {
      calls.push(current.fingerprint);
      return current.fingerprint === 'a' ? iteration('b', 2) : iteration('b', 3);
    }, 5);

    expect(calls).toEqual(['a', 'b']);
    expect(result).toEqual(iteration('b', 3));
    expect(result).not.toBe(seed);
  });

  it('throws NON_CONVERGENCE for a repeated cycle', () => {
    expect(() => convergeLayout(
      iteration('a', 1),
      (current) => current.fingerprint === 'a' ? iteration('b', 2) : iteration('a', 3),
      5,
    )).toThrow(/NON_CONVERGENCE.*cycle/i);
  });

  it('throws NON_CONVERGENCE at the hard limit', () => {
    expect(() => convergeLayout(
      iteration('0', 0),
      (current) => iteration(String(Number(current.fingerprint) + 1), current.pageCount + 1),
      2,
    )).toThrow(/NON_CONVERGENCE.*limit/i);
  });
});
