import { describe, expect, it } from 'vitest';
import {
  attachDocumentLayoutRuntime,
  attachPaintResourceRegistry,
  attachPrivateResourceLookup,
  createFieldAcquisitionServicesView,
  createImmutableResourceLookup,
  documentLayoutRuntimeOf,
  fieldAcquisitionContextOf,
  paintResourceRegistryOf,
  privateResourceLookupOf,
} from './runtime-state.js';
import { createPaintResourceRegistry } from './paint-resources.js';
import type { LayoutServices } from './types.js';

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

  it('attaches one typed paint resource registry without widening its owner', () => {
    const owner = {};
    const registry = createPaintResourceRegistry([{
      kind: 'math', resourceKey: 'math:a',
    }]);

    attachPaintResourceRegistry(owner, registry);

    expect(paintResourceRegistryOf(owner)).toBe(registry);
    expect(Object.keys(owner)).toEqual([]);
    expect(() => attachPaintResourceRegistry(owner, registry)).toThrow(/already attached/i);
    expect(() => paintResourceRegistryOf({})).toThrow(/not attached/i);
  });

  it('isolates immutable field-acquisition context per service view', () => {
    const services = {
      text: {}, images: {}, math: {},
    } as unknown as LayoutServices;
    const first = createFieldAcquisitionServicesView(services, { totalPages: 2 });
    const second = createFieldAcquisitionServicesView(services, { totalPages: 12 });

    expect(first).not.toBe(services);
    expect(second).not.toBe(first);
    expect(first.text).toBe(services.text);
    expect(fieldAcquisitionContextOf(first)).toEqual({ totalPages: 2 });
    expect(fieldAcquisitionContextOf(second)).toEqual({ totalPages: 12 });
    expect(fieldAcquisitionContextOf(services)).toEqual({ totalPages: 1 });
    expect(Object.isFrozen(first)).toBe(true);
    expect(Object.keys(services)).toEqual(['text', 'images', 'math']);
    expect(() => createFieldAcquisitionServicesView(services, { totalPages: 0 }))
      .toThrow(/positive integer/i);
  });

  it('keeps PAGE occurrence resolution private to its pagination iteration view', () => {
    const services = {
      text: {}, images: {}, math: {},
    } as unknown as LayoutServices;
    const paragraph = {};
    const first = createFieldAcquisitionServicesView(services, {
      totalPages: 2,
      resolvePageField: (candidate, sourceRunIndex) =>
        candidate === paragraph && sourceRunIndex === 3
          ? { pageIndex: 1, displayPageNumber: 50, pageNumberFormat: 'upperRoman' }
          : undefined,
    });
    const second = createFieldAcquisitionServicesView(services, {
      totalPages: 2,
      resolvePageField: () => undefined,
    });

    expect(fieldAcquisitionContextOf(first).resolvePageField?.(paragraph, 3)).toEqual({
      pageIndex: 1, displayPageNumber: 50, pageNumberFormat: 'upperRoman',
    });
    expect(fieldAcquisitionContextOf(first).resolvePageField?.(paragraph, 4)).toBeUndefined();
    expect(fieldAcquisitionContextOf(second).resolvePageField?.(paragraph, 3)).toBeUndefined();
    expect(fieldAcquisitionContextOf(services).resolvePageField).toBeUndefined();
  });
});
