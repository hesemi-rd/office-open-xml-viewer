import { describe, it, expect } from 'vitest';
import type { XlsxViewerOptions } from './viewer.js';

/**
 * Compile-time equality check. `Equal<A, B>` resolves to `true` only when the
 * two types are mutually assignable AND identical; otherwise `false`. Feeding
 * the result to `Expect<...>` makes `tsc --build` (run via `pnpm typecheck`)
 * fail when the assertion is violated. These are erased at runtime, so the
 * `it(...)` body below also keeps the file as a valid (trivially green) vitest
 * suite — the real gate is the type checker.
 */
type Equal<A, B> =
  (<T>() => T extends A ? 1 : 2) extends (<T>() => T extends B ? 1 : 2) ? true : false;
type Expect<T extends true> = T;

type SheetCb = NonNullable<XlsxViewerOptions['onSheetChange']>;

/**
 * Cross-viewer API consistency (v1.0 API freeze): `onSheetChange` must share
 * the `(index: number, total: number) => void` shape used by the docx
 * (`onPageChange`) and pptx (`onSlideChange`) viewers, so downstream code can
 * treat the three viewers uniformly. The second argument is the *total* number
 * of sheets — NOT the sheet name. Consumers that need the name look it up via
 * `workbook.sheetNames[index]`.
 */
type _AssertSheetCb = Expect<Equal<SheetCb, (index: number, total: number) => void>>;

describe('XlsxViewer onSheetChange signature', () => {
  it('matches the docx/pptx (index, total) contract', () => {
    // Enforced by `pnpm typecheck` via the `_AssertSheetCb` type above.
    // `total` is the sheet count; the name is read from `workbook.sheetNames`.
    expect(true).toBe(true);
  });
});
