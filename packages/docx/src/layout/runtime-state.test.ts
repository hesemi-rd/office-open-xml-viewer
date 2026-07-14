import { describe, expect, it } from 'vitest';
import {
  attachDocumentLayoutRuntime,
  attachPrivateResourceLookup,
  createImmutableResourceLookup,
  documentLayoutRuntimeOf,
  privateResourceLookupOf,
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

  it('attaches a private lookup once and enforces exact declared membership', () => {
    const owner = {};
    const handle = { id: 'a' };
    attachPrivateResourceLookup(owner, new Map([['a', handle]]), ['a']);
    expect(privateResourceLookupOf(owner)?.resolve('a')).toBe(handle);
    expect(() => attachPrivateResourceLookup(owner, new Map([['a', handle]]), ['a']))
      .toThrow(/already attached/i);

    expect(() => attachPrivateResourceLookup({}, new Map(), ['a']))
      .toThrow(/membership.*missing/i);
    expect(() => attachPrivateResourceLookup({}, new Map([['extra', handle]]), []))
      .toThrow(/membership.*extra/i);
  });
});
