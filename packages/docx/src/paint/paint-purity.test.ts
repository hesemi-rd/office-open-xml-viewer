import { describe, expect, it } from 'vitest';
import type { SectionLayoutContext } from '../layout-context.js';
import type { DocumentLayout } from '../layout/types.js';
import { paintLayoutPage } from './canvas-page.js';

describe('paintLayoutPage', () => {
  it('paints retained geometry without measuring text', async () => {
    const fills: Array<{ fill: string; args: number[] }> = [];
    let currentFill = '';
    const context = {
      get fillStyle() { return currentFill; },
      set fillStyle(value: string | CanvasGradient | CanvasPattern) { currentFill = String(value); },
      save() {},
      restore() {},
      setTransform() {},
      clearRect() {},
      fillRect(...args: number[]) {
        fills.push({ fill: currentFill, args });
      },
      measureText() {
        throw new Error('paint must not measure text');
      },
    } as unknown as CanvasRenderingContext2D;
    const target = {
      width: 0,
      height: 0,
      getContext: () => context,
    } as unknown as HTMLCanvasElement;
    const layout: DocumentLayout = {
      pages: [{
        pageIndex: 0,
        geometry: {
          xPt: 0,
          yPt: 0,
          widthPt: 100,
          heightPt: 200,
          contentTopPt: 10,
          contentBottomPt: 190,
        },
        section: {} as SectionLayoutContext,
        layers: {
          paintOrder: [{ layer: 'body', nodeId: 'drawing-1' }],
          background: [],
          behindText: [],
          header: [],
          body: [{
            kind: 'drawing',
            id: 'drawing-1',
            source: { story: 'body', storyInstance: 'body', path: [0] },
            flowBounds: { xPt: 10, yPt: 20, widthPt: 30, heightPt: 40 },
            inkBounds: { xPt: 10, yPt: 20, widthPt: 30, heightPt: 40 },
            advancePt: 40,
            ordinaryFlow: true,
            commands: [{
              kind: 'fill-rect',
              rect: { xPt: 10, yPt: 20, widthPt: 30, heightPt: 40 },
              fill: '#ff0000',
            }],
          }],
          notes: [],
          front: [],
          footer: [],
        },
        readingOrder: ['drawing-1'],
      }],
      diagnostics: [],
    };

    await expect(paintLayoutPage(layout, 0, target, { scale: 1, dpr: 1 })).resolves.toBeUndefined();
    expect(fills).toEqual([{ fill: '#ff0000', args: [10, 20, 30, 40] }]);
  });
});
