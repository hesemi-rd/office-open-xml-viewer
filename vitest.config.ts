import { defineConfig } from 'vitest/config';

// Unit tests only. `*.spec.ts` is reserved for the Playwright visual-regression
// suites (run via `pnpm vrt`), so restrict vitest to `*.test.ts` under src and
// keep it out of the tests/visual directories.
export default defineConfig({
  test: {
    include: ['packages/**/src/**/*.test.ts'],
    exclude: ['**/node_modules/**', '**/dist/**', '**/tests/visual/**'],
    // The suite includes WASM parsing and native Skia render probes whose CPU
    // work happens outside Vitest's JavaScript scheduler. Shared CI runners can
    // therefore starve two otherwise independent workers long enough that a
    // synchronous render returns only after Vitest's timeout should have fired.
    // Serialize files in CI; retain bounded parallelism for local development.
    maxWorkers: process.env.CI ? 1 : '50%',
  },
});
