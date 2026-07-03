import { defineConfig, devices } from '@playwright/test';

export default defineConfig({
  testDir: '.',
  testMatch: '**/*.spec.ts',
  fullyParallel: false,
  reporter: [['list']],
  use: {
    baseURL: 'http://localhost:6007',
    actionTimeout: 30_000,
  },
  // The smoke assertion is canvasHasInk (a count of non-white pixels), which is
  // font-independent, so it stays stable across engines. Running webkit and
  // firefox alongside chromium catches engine-specific breakage in the
  // parse -> render pipeline (OffscreenCanvas quirks, worker transfer, canvas
  // API gaps) that a Chrome-only smoke would miss. (VRT match-% comparisons,
  // which ARE font-sensitive, stay Chrome-only and local — see visual.spec.ts.)
  projects: [
    {
      name: 'chrome',
      use: {
        channel: 'chrome',
        deviceScaleFactor: 1,
        viewport: { width: 1400, height: 900 },
      },
    },
    {
      name: 'webkit',
      use: {
        ...devices['Desktop Safari'],
        deviceScaleFactor: 1,
        viewport: { width: 1400, height: 900 },
      },
    },
    {
      name: 'firefox',
      use: {
        ...devices['Desktop Firefox'],
        deviceScaleFactor: 1,
        viewport: { width: 1400, height: 900 },
      },
    },
  ],
  webServer: {
    command: 'pnpm storybook --port 6007 --no-open',
    url: 'http://localhost:6007/iframe.html',
    // Locally, reuse a Storybook already serving on 6007; in CI always boot a
    // fresh one so the run never binds to a stale/unrelated server.
    reuseExistingServer: !process.env.CI,
    timeout: 120_000,
  },
});
