import { describe, expect, it } from 'vitest';
import {
  attachDocumentLayoutRuntime,
  createImmutableResourceLookup,
  documentLayoutRuntimeOf,
} from './runtime-state.js';

describe('document layout runtime state', () => {
  it('requires explicit deterministic attachment', () => {
    expect(() => documentLayoutRuntimeOf({})).toThrow(/runtime.*not initialized/i);

    const owner = {};
    attachDocumentLayoutRuntime(owner, 123);
    expect(documentLayoutRuntimeOf(owner).defaultCurrentDateMs).toBe(123);
  });

  it('hides mutable handles behind immutable, fixed membership', () => {
    const first = { id: 'first' };
    const entries = new Map<string, object>([['math:a', first]]);
    const lookup = createImmutableResourceLookup(entries);
    entries.set('math:b', { id: 'late' });

    expect(lookup.keys).toEqual(['math:a']);
    expect(Object.isFrozen(lookup.keys)).toBe(true);
    expect(lookup.resolve('math:a')).toBe(first);
    expect(() => lookup.resolve('math:b')).toThrow(/Unknown runtime resource/);
  });
});
