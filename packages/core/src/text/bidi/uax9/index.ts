// From-scratch UAX#9 engine — the swappable implementation behind BidiEngine.
//
// NOTE: Task 1 scaffold ships a trivial LTR/RTL-uniform placeholder so the seam
// is wired and importable. The full P2-P3 / X / W / N / I / L rule set is
// implemented in Task 3 (`./rules.ts`), conformance-tested against Unicode's
// BidiCharacterTest.txt / BidiTest.txt. Nothing consumes the engine until the
// renderers (PR2+), so this placeholder is never user-visible.

import type { BidiEngine } from '../engine.js';
import type { BaseDirection, BidiLevels } from '../types.js';

class Uax9Engine implements BidiEngine {
  computeLevels(
    text: string,
    base: BaseDirection,
  ): { levels: BidiLevels; paragraphLevel: number } {
    const paragraphLevel = base === 'rtl' ? 1 : 0;
    const levels = new Uint8Array(text.length).fill(paragraphLevel);
    return { levels, paragraphLevel };
  }

  reorderVisual(_levels: BidiLevels, start: number, end: number): number[] {
    const order: number[] = [];
    for (let i = start; i < end; i++) order.push(i);
    return order;
  }

  getMirror(_codePoint: number): number | null {
    return null;
  }
}

export function createUax9Engine(): BidiEngine {
  return new Uax9Engine();
}
