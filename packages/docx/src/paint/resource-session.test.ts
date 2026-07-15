import { describe, expect, it } from 'vitest';
import {
  createPaintResourceRegistry,
} from '../layout/paint-resources.js';
import type { PaintResourceDescriptor } from '../layout/types.js';
import {
  createPaintResourceSession,
  createProductionPaintResourceSession,
  isUnavailablePaintResourceHandle,
  unavailablePaintResourceHandle,
} from './resource-session.js';

const descriptors: readonly PaintResourceDescriptor[] = [{
  kind: 'image', resourceKey: 'image:body:0', partPath: 'word/media/image.png',
  mimeType: 'image/png', intrinsicSize: { widthPt: 20, heightPt: 10 },
}, {
  kind: 'chart', resourceKey: 'chart:body:1',
  intrinsicSize: { widthPt: 100, heightPt: 60 }, model: {} as never,
}, {
  kind: 'math', resourceKey: 'math:body:2:inline',
}];

describe('paint-owned resource session', () => {
  it('keeps realm-local opaque handles outside the clone-safe layout registry', () => {
    const registry = createPaintResourceRegistry(descriptors);
    const drawable = { realm: 'main' };
    const session = createPaintResourceSession(registry, [{
      resourceKey: 'image:body:0', kind: 'image', handle: drawable,
    }]);

    expect(session.resolve('image:body:0', 'image')).toEqual({
      descriptor: registry.resolve('image:body:0', 'image'),
      handle: drawable,
    });
    expect(registry.descriptors).not.toContain(drawable);
    expect(() => session.resolve('math:body:2:inline', 'math'))
      .toThrow(/Missing paint resource handle.*math:body:2:inline/);
  });

  it('rejects duplicate, unknown, kind-mismatched, and forged unavailable entries', () => {
    const registry = createPaintResourceRegistry(descriptors);
    const entry = { resourceKey: 'math:body:2:inline', kind: 'math' as const, handle: {} };

    expect(() => createPaintResourceSession(registry, [entry, { ...entry }]))
      .toThrow(/Duplicate paint resource handle.*math:body:2:inline/);
    expect(() => createPaintResourceSession(registry, [{
      resourceKey: 'missing', kind: 'image', handle: {},
    }])).toThrow(/Unknown paint resource: missing/);
    expect(() => createPaintResourceSession(registry, [{
      resourceKey: 'chart:body:1', kind: 'image', handle: {},
    }])).toThrow(/kind mismatch.*expected image.*chart/i);
    expect(() => createPaintResourceSession(registry, [{
      resourceKey: 'math:body:2:inline', kind: 'math',
      handle: { status: 'unavailable', reason: '   ' },
    }])).toThrow(/unavailable paint resource reason.*non-empty/i);
  });

  it('distinguishes an explicit unavailable drawable from a missing contract entry', () => {
    const registry = createPaintResourceRegistry(descriptors);
    const unavailable = unavailablePaintResourceHandle('optional math renderer unavailable');
    const session = createPaintResourceSession(registry, [{
      resourceKey: 'math:body:2:inline', kind: 'math', handle: unavailable,
    }]);

    expect(isUnavailablePaintResourceHandle(
      session.resolve('math:body:2:inline', 'math').handle,
    )).toBe(true);
    expect(structuredClone(registry.descriptors)).toEqual(registry.descriptors);
    expect(JSON.stringify(registry.descriptors)).not.toContain(unavailable.reason);
  });

  it('builds the complete production session without moving handles into layout', () => {
    const registry = createPaintResourceRegistry(descriptors);
    const image = { realm: 'main' };
    const unavailable = unavailablePaintResourceHandle('optional math renderer unavailable');
    const session = createProductionPaintResourceSession(registry, (descriptor) => {
      if (descriptor.kind === 'image') return image;
      if (descriptor.kind === 'math') return unavailable;
      return undefined;
    });

    expect(session.resolve('image:body:0', 'image').handle).toBe(image);
    expect(session.resolve('chart:body:1', 'chart').handle).toBeNull();
    expect(session.resolve('math:body:2:inline', 'math').handle).toBe(unavailable);
  });
});
