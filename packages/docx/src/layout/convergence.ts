import type { DocumentLayout } from './types.js';
import { LayoutInvariantError } from './diagnostics.js';

export interface LayoutIteration {
  readonly fingerprint: string;
  readonly pageCount: number;
  readonly layout?: DocumentLayout;
}

export function convergeLayout<T extends LayoutIteration>(
  seed: T,
  step: (iteration: T) => T,
  limit: number,
): T {
  if (!Number.isInteger(limit) || limit < 1) {
    throw new LayoutInvariantError('NON_CONVERGENCE', 'limit must be a positive integer');
  }
  const seen = new Set([seed.fingerprint]);
  let current = seed;
  for (let iteration = 0; iteration < limit; iteration += 1) {
    const next = step(current);
    if (next.fingerprint === current.fingerprint) return next;
    if (seen.has(next.fingerprint)) {
      throw new LayoutInvariantError('NON_CONVERGENCE', `repeated geometry fingerprint cycle at ${next.fingerprint}`);
    }
    seen.add(next.fingerprint);
    current = next;
  }
  throw new LayoutInvariantError('NON_CONVERGENCE', `hard iteration limit ${limit} reached`);
}
