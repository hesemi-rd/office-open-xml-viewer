import { defineConfig } from '@playwright/test';

export default defineConfig({
  testDir: '.',
  testMatch: '**/*.spec.ts',
  fullyParallel: false,
  reporter: [['list']],
  use: {
    baseURL: 'http://localhost:6007',
    actionTimeout: 30_000,
  },
  projects: [
    {
      name: 'chrome',
      use: {
        channel: 'chrome',
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
