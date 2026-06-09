// The swap seam. The low-level UAX#9 surface is isolated behind this interface
// so a from-scratch implementation can be replaced by a thin third-party
// adapter (e.g. bidi-js: getEmbeddingLevels / getReorderSegments /
// getMirroredCharacter) without touching renderers or `toVisualSegments`,
// which sit ABOVE this seam.

import type { BaseDirection, BidiLevels } from './types.js';
import { createUax9Engine } from './uax9/index.js';

export interface BidiEngine {
  /**
   * UAX#9 P2-P3 + X..I rules. Returns the resolved embedding level per code
   * unit and the paragraph base level. Mirrors bidi-js `getEmbeddingLevels`.
   */
  computeLevels(
    text: string,
    base: BaseDirection,
  ): { levels: BidiLevels; paragraphLevel: number };

  /**
   * UAX#9 L2. Given resolved levels for the line range [start, end), returns the
   * visual-order permutation: result[i] is the logical index drawn at visual
   * position i. Mirrors bidi-js `getReorderSegments` / `getReorderedIndices`.
   */
  reorderVisual(levels: BidiLevels, start: number, end: number): number[];

  /**
   * UAX#9 L4. Mirror glyph for a character displayed at an odd (RTL) level, or
   * null when the character has no mirror. Mirrors bidi-js
   * `getMirroredCharacter`.
   */
  getMirror(codePoint: number): number | null;
}

let current: BidiEngine | null = null;

/** The active engine. Lazily instantiates the default UAX#9 engine. */
export function getDefaultBidiEngine(): BidiEngine {
  if (current === null) current = createUax9Engine();
  return current;
}

/** Swap the active engine (tests, or a third-party adapter). */
export function setBidiEngine(engine: BidiEngine): void {
  current = engine;
}

/** Restore the built-in UAX#9 engine. */
export function resetBidiEngine(): void {
  current = null;
}
