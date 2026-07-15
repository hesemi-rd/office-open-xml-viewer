import { defineConfig } from 'vitest/config';

// Unit tests only. `*.spec.ts` is reserved for the Playwright visual-regression
// suites (run via `pnpm vrt`), so restrict vitest to `*.test.ts` under src and
// keep it out of the tests/visual directories.
export default defineConfig({
  test: {
    include: ['packages/**/src/**/*.test.ts'],
    exclude: ['**/node_modules/**', '**/dist/**', '**/tests/visual/**'],
    // The suite includes WASM parsing and native Skia render probes. Letting
    // Vitest run near the available CPU parallelism makes those files contend
    // for host resources, increasing wall time enough to trigger unrelated
    // per-test timeouts. Half the available parallelism leaves headroom for the
    // native work performed inside each worker and reduces CI wall time.
    maxWorkers: '50%',
  },
});
