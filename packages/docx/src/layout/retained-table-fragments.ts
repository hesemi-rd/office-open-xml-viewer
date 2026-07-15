import type { TableLayout } from './types.js';

export interface RetainedTableEnvelope {
  readonly fragment: TableLayout;
  readonly xPt?: number;
  readonly yPt: number;
  readonly widthPt?: number;
}

const retainedTableEnvelopes = new WeakMap<object, RetainedTableEnvelope>();

export function retainTableEnvelope(
  envelope: object,
  placement: RetainedTableEnvelope,
): void {
  retainedTableEnvelopes.set(envelope, Object.freeze({ ...placement }));
}

export function retainedTableEnvelopeFor(
  envelope: object,
): RetainedTableEnvelope | undefined {
  return retainedTableEnvelopes.get(envelope);
}

export function retainedTableSliceSize(
  envelope: object,
  scale: number,
): Readonly<{ widthPx: number; heightPx: number }> {
  const fragment = retainedTableEnvelopes.get(envelope)?.fragment;
  if (!fragment) {
    throw new Error('Floating-table wrap registration requires a retained table slice');
  }
  return Object.freeze({
    widthPx: fragment.columnWidthsPt.reduce((sum, width) => sum + width, 0) * scale,
    heightPx: fragment.advancePt * scale,
  });
}
