import { describe, expect, it } from 'vitest';
import { snapshotPlainData } from './plain-data.js';

describe('plain layout data snapshots', () => {
  it('preserves signed unbounded finite DrawingML source-rectangle percentages exactly', () => {
    const authored = { l: -0.25, t: 1.25, r: 1.5, b: -0.75 };
    const source = { srcRect: { ...authored } };
    const snapshot = snapshotPlainData(source, 'paint resource');
    source.srcRect.l = 0;

    expect(snapshot.srcRect).toEqual(authored);
    expect(structuredClone(snapshot).srcRect).toEqual(authored);
  });

  it('clones and deeply freezes plain data while preserving optional undefined values', () => {
    const source = { optional: undefined, nested: { values: [1, 'two'] } };
    const snapshot = snapshotPlainData(source, 'layout payload');

    source.nested.values[0] = 99;

    expect(snapshot).toEqual({ optional: undefined, nested: { values: [1, 'two'] } });
    expect(Object.isFrozen(snapshot)).toBe(true);
    expect(Object.isFrozen(snapshot.nested)).toBe(true);
    expect(Object.isFrozen(snapshot.nested.values)).toBe(true);
  });

  it.each([
    [() => undefined],
    [Symbol('invalid')],
    [1n],
    [new Map([['key', 'value']])],
  ])('rejects non-plain data %#', (invalid) => {
    expect(() => snapshotPlainData({ invalid }, 'layout payload'))
      .toThrow(/structured-clone-safe plain data/i);
  });

  it('rejects cyclic plain data', () => {
    const cyclic: { self?: object } = {};
    cyclic.self = cyclic;

    expect(() => snapshotPlainData(cyclic, 'layout payload'))
      .toThrow(/structured-clone-safe plain data/i);
  });

  it('validates a shared plain-data DAG only once before cloning', () => {
    let reads = 0;
    const shared = Object.defineProperty({}, 'value', {
      enumerable: true,
      get() { reads += 1; return 7; },
    });

    const snapshot = snapshotPlainData({ first: shared, second: shared }, 'layout payload');

    expect(snapshot.first).toBe(snapshot.second);
    expect(reads).toBe(2); // one validation traversal and one structured clone traversal
  });
});
