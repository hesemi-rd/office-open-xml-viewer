import { describe, expect, it } from 'vitest';
import { createPaintResourceRegistry } from '../layout/paint-resources.js';
import type {
  AcquiredParagraphLayoutInput,
  PaintResourceDescriptor,
} from '../layout/types.js';
import { layoutParagraph } from '../layout/paragraph.js';
import { stableFingerprint } from '../layout/fingerprint.js';
import {
  createPaintResourceSession,
  isUnavailablePaintResourceHandle,
  unavailablePaintResourceHandle,
} from './resource-session.js';

const descriptors: readonly PaintResourceDescriptor[] = [{
  kind: 'image',
  resourceKey: 'image:body:0',
  partPath: 'word/media/image1.png',
  mimeType: 'image/png',
  intrinsicSize: { widthPt: 20, heightPt: 10 },
}, {
  kind: 'math',
  resourceKey: 'math:body:1:inline',
}];

function retainedParagraph(): ReturnType<typeof layoutParagraph> {
  const input: AcquiredParagraphLayoutInput = {
    kind: 'paragraph',
    id: 'body:0',
    source: { story: 'body', storyInstance: 'body', path: [0] },
    flowDomainId: 'body',
    ordinaryFlow: true,
    flowBounds: { xPt: 10, yPt: 20, widthPt: 200, heightPt: 12 },
    inkBounds: { xPt: 10, yPt: 20, widthPt: 20, heightPt: 10 },
    spacing: { beforePt: 0, afterPt: 0 },
    lines: [{
      range: { start: 0, end: 1 },
      bounds: { xPt: 10, yPt: 20, widthPt: 20, heightPt: 10 },
      baselinePt: 28,
      advancePt: 12,
      placements: [{
        kind: 'resource',
        range: { start: 0, end: 1 },
        resourceKey: 'image:body:0',
        resourceKind: 'image',
        bounds: { xPt: 10, yPt: 20, widthPt: 20, heightPt: 10 },
        advancePt: 20,
      }],
    }],
    borders: [],
    resources: [{
      kind: 'image',
      resourceKey: 'image:body:0',
      intrinsicSize: { widthPt: 20, heightPt: 10 },
    }],
    drawings: [],
    textBoxes: [],
    events: [],
    exclusions: [],
  };
  return layoutParagraph(input);
}

describe('worker retained-resource boundary', () => {
  it('reconstructs clone-safe retained data while keeping handles local to each realm', () => {
    const mainRegistry = createPaintResourceRegistry(descriptors);
    const workerRegistry = createPaintResourceRegistry(
      structuredClone(mainRegistry.descriptors) as PaintResourceDescriptor[],
    );
    const mainImage = { realm: 'main' };
    const workerImage = { realm: 'worker' };
    const mainSession = createPaintResourceSession(mainRegistry, [{
      resourceKey: 'image:body:0', kind: 'image', handle: mainImage,
    }]);
    const workerSession = createPaintResourceSession(workerRegistry, [{
      resourceKey: 'image:body:0', kind: 'image', handle: workerImage,
    }]);
    const paragraph = retainedParagraph();
    const workerParagraph = structuredClone(paragraph);

    expect(workerRegistry.descriptors).toEqual(mainRegistry.descriptors);
    expect(workerParagraph).toEqual(paragraph);
    expect(stableFingerprint('paragraph', workerParagraph))
      .toBe(stableFingerprint('paragraph', paragraph));
    expect(mainSession.resolve('image:body:0', 'image').handle).toBe(mainImage);
    expect(workerSession.resolve('image:body:0', 'image').handle).toBe(workerImage);
    expect(mainSession.resolve('image:body:0', 'image').handle).not.toBe(workerImage);
    expect(JSON.stringify(workerRegistry.descriptors)).not.toContain('worker');
    expect(JSON.stringify(workerParagraph)).not.toContain('worker');
  });

  it('distinguishes an explicit realm-local unavailable handle from a missing entry', () => {
    const registry = createPaintResourceRegistry(descriptors);
    const unavailable = unavailablePaintResourceHandle('optional worker math renderer unavailable');
    const session = createPaintResourceSession(registry, [{
      resourceKey: 'math:body:1:inline', kind: 'math', handle: unavailable,
    }]);

    expect(isUnavailablePaintResourceHandle(
      session.resolve('math:body:1:inline', 'math').handle,
    )).toBe(true);
    expect(() => session.resolve('image:body:0', 'image'))
      .toThrow(/Missing paint resource handle.*image:body:0/);
    expect(structuredClone(registry.descriptors)).toEqual(registry.descriptors);
    expect(JSON.stringify(registry.descriptors)).not.toContain(unavailable.reason);
  });

  it('rejects a forged unavailable marker without an actionable reason', () => {
    const registry = createPaintResourceRegistry(descriptors);

    expect(() => createPaintResourceSession(registry, [{
      resourceKey: 'math:body:1:inline',
      kind: 'math',
      handle: { status: 'unavailable', reason: '   ' },
    }])).toThrow(/unavailable paint resource reason.*non-empty/i);
  });
});
