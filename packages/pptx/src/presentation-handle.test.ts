import { describe, it, expect } from 'vitest';
import type { PresentationHandle } from './presentation-handle.js';

/**
 * Compile-time assertions enforced by `pnpm typecheck`. They are erased at
 * runtime, so the `it(...)` body keeps the file a valid (trivially green)
 * vitest suite.
 */
type HasKey<T, K extends PropertyKey> = K extends keyof T ? true : false;
type Expect<T extends true> = T;

/**
 * v1.0 API freeze: the live-playback handle's teardown method is named
 * `destroy()`, matching `PptxViewer.destroy()`, `DocxDocument.destroy()`,
 * `XlsxWorkbook.destroy()` etc. The old `dispose()` name is removed so the
 * whole public surface uses one teardown verb.
 */
type _HasDestroy = Expect<HasKey<PresentationHandle, 'destroy'>>;
// `dispose` must NOT exist on the handle any more.
type _NoDispose = Expect<HasKey<PresentationHandle, 'dispose'> extends false ? true : false>;

describe('PresentationHandle teardown method', () => {
  it('exposes destroy() (renamed from dispose()) for API parity', () => {
    // Enforced by `pnpm typecheck` via the `_HasDestroy` / `_NoDispose` types.
    expect(true).toBe(true);
  });
});
