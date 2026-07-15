import { describe, expect, it } from 'vitest';
import {
  enqueueDeferredFrontPaint,
  withDeferredFrontPaintSession,
  type DeferredFrontPaintState,
} from './deferred-front-session.js';

describe('deferred front paint session', () => {
  it('flushes queued front paint in insertion order after the story body', () => {
    const state: DeferredFrontPaintState = {};
    const order: string[] = [];

    withDeferredFrontPaintSession(state, () => {
      order.push('body');
      expect(enqueueDeferredFrontPaint(state, () => order.push('front-1'))).toBe(true);
      expect(enqueueDeferredFrontPaint(state, () => order.push('front-2'))).toBe(true);
      expect(order).toEqual(['body']);
    });

    expect(order).toEqual(['body', 'front-1', 'front-2']);
    expect(state.frontPaintSession).toBeUndefined();
  });

  it('restores an outer session before flushing a nested session', () => {
    const state: DeferredFrontPaintState = {};
    const order: string[] = [];

    withDeferredFrontPaintSession(state, () => {
      const outerSession = state.frontPaintSession;
      expect(enqueueDeferredFrontPaint(state, () => order.push('outer-front'))).toBe(true);
      order.push('outer-body-before');

      withDeferredFrontPaintSession(state, () => {
        expect(state.frontPaintSession).not.toBe(outerSession);
        order.push('inner-body');
        expect(enqueueDeferredFrontPaint(state, () => {
          // Flushing must paint immediately. If the restored outer session were
          // visible here, the inner front paint would leak into its queue.
          expect(enqueueDeferredFrontPaint(state, () => order.push('leaked'))).toBe(false);
          order.push('inner-front');
        })).toBe(true);
      });

      expect(state.frontPaintSession).toBe(outerSession);
      order.push('outer-body-after');
    });

    expect(order).toEqual([
      'outer-body-before',
      'inner-body',
      'inner-front',
      'outer-body-after',
      'outer-front',
    ]);
  });

  it('restores the previous queue without flushing when story paint throws', () => {
    const state: DeferredFrontPaintState = {};
    let queuedPaintRan = false;

    withDeferredFrontPaintSession(state, () => {
      const previousSession = state.frontPaintSession;
      expect(() => withDeferredFrontPaintSession(state, () => {
        enqueueDeferredFrontPaint(state, () => { queuedPaintRan = true; });
        throw new Error('story paint failed');
      })).toThrow('story paint failed');

      expect(state.frontPaintSession).toBe(previousSession);
    });

    expect(state.frontPaintSession).toBeUndefined();
    expect(queuedPaintRan).toBe(false);
  });

  it('closes the shared session before replay so cloned states cannot re-enqueue', () => {
    const state: DeferredFrontPaintState = {};
    let clonedState: DeferredFrontPaintState | undefined;
    let replayCount = 0;

    withDeferredFrontPaintSession(state, () => {
      // RenderState views for cells and nested stories are shallow copies. They
      // retain the same session object even after the owner state restores its
      // previous session, so replay must close the shared object itself.
      clonedState = { ...state };
      expect(enqueueDeferredFrontPaint(state, () => {
        replayCount += 1;
        expect(enqueueDeferredFrontPaint(clonedState!, () => {
          replayCount += 100;
        })).toBe(false);
      })).toBe(true);
    });

    expect(replayCount).toBe(1);
  });

  it('replays retained front drawings by relative height then source order', () => {
    const state: DeferredFrontPaintState = {};
    const order: string[] = [];

    withDeferredFrontPaintSession(state, () => {
      enqueueDeferredFrontPaint(state, () => order.push('high-late'), {
        relativeHeight: 20, sourceOrder: 3,
      });
      enqueueDeferredFrontPaint(state, () => order.push('low'), {
        relativeHeight: 10, sourceOrder: 2,
      });
      enqueueDeferredFrontPaint(state, () => order.push('high-early'), {
        relativeHeight: 20, sourceOrder: 1,
      });
    });

    expect(order).toEqual(['low', 'high-early', 'high-late']);
  });
});
