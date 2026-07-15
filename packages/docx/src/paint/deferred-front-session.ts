export type DeferredFrontPaint = () => void;
export interface DeferredFrontPaintOrder {
  readonly relativeHeight: number;
  readonly sourceOrder: number;
}

export interface DeferredFrontPaintSession {
  readonly kind: 'deferred-front-paint-session';
}

export interface DeferredFrontPaintState {
  frontPaintSession?: DeferredFrontPaintSession | null;
}

interface MutableDeferredFrontPaintSession extends DeferredFrontPaintSession {
  readonly paints: Array<Readonly<{
    paint: DeferredFrontPaint;
    order: DeferredFrontPaintOrder;
    enqueueOrder: number;
  }>>;
  accepting: boolean;
}

export function enqueueDeferredFrontPaint(
  state: DeferredFrontPaintState,
  paint: DeferredFrontPaint,
  order?: DeferredFrontPaintOrder,
): boolean {
  const session = state.frontPaintSession as MutableDeferredFrontPaintSession | null | undefined;
  if (!session?.accepting) return false;
  session.paints.push({
    paint,
    order: order ?? { relativeHeight: 0, sourceOrder: session.paints.length },
    enqueueOrder: session.paints.length,
  });
  return true;
}

/**
 * Owns the mutable front-layer queue while a story paints.
 *
 * B3 replaces closure collection with immutable PageLayers. Until then, keeping
 * queue replacement/restoration here prevents each story renderer from growing
 * a subtly different nesting and exception contract.
 */
export function withDeferredFrontPaintSession(
  state: DeferredFrontPaintState,
  paintStory: () => void,
): void {
  const previousSession = state.frontPaintSession;
  const session: MutableDeferredFrontPaintSession = {
    kind: 'deferred-front-paint-session',
    paints: [],
    accepting: true,
  };
  state.frontPaintSession = session;
  try {
    paintStory();
  } finally {
    // Child RenderState views are shallow copies and can retain this same
    // session object after the owner restores its previous session. Closing the
    // shared object prevents replayed front paints from appending to the array
    // currently being iterated, which would otherwise grow without bound.
    session.accepting = false;
    state.frontPaintSession = previousSession;
  }
  const paints = [...session.paints].sort((a, b) =>
    a.order.relativeHeight - b.order.relativeHeight
    || a.order.sourceOrder - b.order.sourceOrder
    || a.enqueueOrder - b.enqueueOrder);
  for (const entry of paints) {
    // The story's enclosing session is already restored for nesting, but a
    // replayed front object must paint now rather than enqueue into that owner.
    state.frontPaintSession = null;
    try {
      entry.paint();
    } finally {
      state.frontPaintSession = previousSession;
    }
  }
}
