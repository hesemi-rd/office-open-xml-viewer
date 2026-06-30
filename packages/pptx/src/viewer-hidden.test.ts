import { describe, it, expect } from 'vitest';
import type { PptxViewer, PptxViewerOptions, HiddenSlideMode } from './viewer.js';
import { nextVisibleIndex, resolveVisibleIndex } from './hidden.js';

/**
 * Compile-time API-surface assertions (erased at runtime, enforced by
 * `pnpm typecheck`). PptxViewer is DOM-bound and the vitest env is `node`, so —
 * like notes.test.ts — the viewer is verified at the type level; its policy
 * logic is delegated to the pure helpers tested in hidden.test.ts.
 */
type Expect<T extends true> = T;
type Equal<A, B> =
  (<T>() => T extends A ? 1 : 2) extends (<T>() => T extends B ? 1 : 2) ? true : false;

type _Modes = Expect<Equal<HiddenSlideMode, 'show' | 'skip' | 'dim'>>;
type _SetMode = Expect<Equal<PptxViewer['setHiddenSlideMode'], (mode: HiddenSlideMode) => Promise<void>>>;
type _ModeGetter = Expect<Equal<PptxViewer['hiddenSlideMode'], HiddenSlideMode>>;
type _VisibleCount = Expect<Equal<PptxViewer['visibleSlideCount'], number>>;
const _opts: PptxViewerOptions = { hiddenSlideMode: 'dim', hiddenSlideDim: { color: '#fff', opacity: 0.5 } };

describe('PptxViewer hidden-slide policy (delegated pure helpers)', () => {
  it('skip navigation jumps over hidden slides via the tested helpers', () => {
    const isHidden = (i: number) => i === 1; // visible: 0, 2
    expect(nextVisibleIndex(0, 1, isHidden, 3)).toBe(2); // nextSlide from 0 → 2
    expect(resolveVisibleIndex(1, isHidden, 3)).toBe(2); // entering skip on hidden → 2
    void _opts;
  });
});
