import { describe, expect, it } from 'vitest';
import type { SectionLayoutContext } from '../layout-context.js';
import { deepFreezeDocumentLayout, layoutFingerprint } from './invariants.js';
import type { DocumentLayout } from './types.js';

describe('DocumentLayout data boundary', () => {
  it('is deeply immutable and structured-clone safe without platform objects', () => {
    const layout: DocumentLayout = {
      pages: [{
        pageIndex: 0,
        geometry: {
          xPt: 0,
          yPt: 0,
          widthPt: 612,
          heightPt: 792,
          contentTopPt: 72,
          contentBottomPt: 720,
        },
        section: {} as SectionLayoutContext,
        layers: {
          paintOrder: [{ layer: 'front', nodeId: 'drawing-1' }],
          background: [],
          behindText: [],
          header: [],
          body: [],
          notes: [],
          front: [{
            kind: 'drawing',
            id: 'drawing-1',
            source: { story: 'body', storyInstance: 'body', path: [0, 1] },
            flowBounds: { xPt: 10, yPt: 20, widthPt: 30, heightPt: 40 },
            inkBounds: { xPt: 9, yPt: 19, widthPt: 32, heightPt: 42 },
            clipBounds: { xPt: 10, yPt: 20, widthPt: 30, heightPt: 40 },
            advancePt: 0,
            ordinaryFlow: false,
            transform: { a: 1, b: 0, c: 0, d: 1, e: 10, f: 20 },
            clip: {
              kind: 'polygon',
              points: [{ xPt: 0, yPt: 0 }, { xPt: 30, yPt: 0 }, { xPt: 30, yPt: 40 }],
            },
            commands: [{
              kind: 'fill-rect',
              rect: { xPt: 10, yPt: 20, widthPt: 30, heightPt: 40 },
              fill: '#123456',
            }],
          }],
          footer: [],
        },
        readingOrder: ['drawing-1'],
      }],
      diagnostics: [{
        code: 'UNSUPPORTED_FEATURE',
        severity: 'warning',
        source: { story: 'body', storyInstance: 'body', path: [0, 1] },
        message: 'synthetic diagnostic',
      }],
    };

    const frozen = deepFreezeDocumentLayout(layout);
    const cloned = structuredClone(frozen);
    const clonedNode = cloned.pages[0]?.layers.front[0];
    if (!clonedNode || clonedNode.kind !== 'drawing') throw new Error('drawing clone missing');

    expect(Object.isFrozen(frozen)).toBe(true);
    expect(Object.isFrozen(frozen.pages)).toBe(true);
    expect(Object.isFrozen(frozen.pages[0]?.layers.front[0]?.source.path)).toBe(true);
    expect(Object.getPrototypeOf(clonedNode.transform)).toBe(Object.prototype);
    expect(layoutFingerprint(cloned)).toBe(layoutFingerprint(frozen));
    expect(() => (frozen as unknown as { pages: unknown[] }).pages.push({})).toThrow();
  });
});
