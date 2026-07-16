import { describe, expect, expectTypeOf, it } from 'vitest';
import { layoutFlowBlocks } from './flow.js';
import type {
  BlockLayoutAlgorithms,
  FlowContainer,
  FlowDomain,
  LayoutServices,
} from './types.js';

describe('flow ownership types', () => {
  it('keeps acquisition bounds separate from retained dual-space domains', () => {
    expectTypeOf<FlowContainer>().toHaveProperty('bounds');
    expectTypeOf<FlowDomain>().toHaveProperty('logicalBounds');
    expectTypeOf<FlowDomain>().toHaveProperty('physicalBounds');
    // @ts-expect-error retained page domains intentionally have no ambiguous bounds alias
    const ambiguous = ({} as FlowDomain).bounds;
    expect(ambiguous).toBeUndefined();
  });

  it('continues to acquire flow in container-local logical bounds', () => {
    const algorithms: BlockLayoutAlgorithms = {
      layoutParagraph(input, placement) {
        const bounds = { xPt: 10, yPt: 20, widthPt: 40, heightPt: 10 };
        return {
          layout: {
            kind: 'paragraph', id: 'paragraph', source: input.source,
            flowDomainId: placement.container.id, flowBounds: bounds, inkBounds: bounds,
            advancePt: 10, ordinaryFlow: true, spacing: { beforePt: 0, afterPt: 0 },
            contextualSpacing: false, lines: [], borders: [], resources: [], drawings: [],
            textBoxes: [], events: [], exclusions: [],
          },
          nextCursor: { xPt: 10, yPt: 30 },
        };
      },
      layoutTable() { throw new Error('not used'); },
    };
    const services = {} as LayoutServices;

    const result = layoutFlowBlocks({
      source: { story: 'body', storyInstance: 'body', path: [0] },
      container: {
        id: 'column-local', kind: 'body',
        bounds: { xPt: 10, yPt: 20, widthPt: 100, heightPt: 200 },
      },
      cursor: { xPt: 10, yPt: 20 },
      blocks: [{ kind: 'paragraph', source: { story: 'body', storyInstance: 'body', path: [1] } }],
    }, services, algorithms);

    expect(result.flowBounds).toEqual({ xPt: 10, yPt: 20, widthPt: 40, heightPt: 10 });
  });
});
