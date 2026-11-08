import { defineConfig } from '@playwright/test';

const browserName = process.env.HANG_BROWSER ?? 'chromium';
if (!['chromium', 'firefox', 'webkit'].includes(browserName))
  throw new Error('Invalid HANG_BROWSER');

export default defineConfig({
  testDir: '.',
  outputDir: `../test-results/fixtures-${browserName}`,
  testMatch: 'fixtures.stress.ts',
  timeout: 60_000,
  workers: 1,
  use: {
    browserName: browserName as 'chromium' | 'firefox' | 'webkit',
    viewport: { width: 1440, height: 900 },
    actionTimeout: 10_000,
  },
  projects: [
    { name: 'react', use: { baseURL: 'http://localhost:5303' } },
    { name: 'vue', use: { baseURL: 'http://localhost:5304' } },
  ],
  webServer: [
    {
      command: 'bunx vite --port 5303 --strictPort',
      cwd: '../examples/vite',
      url: 'http://localhost:5303',
      reuseExistingServer: false,
    },
    {
      command: 'bunx vite --port 5304 --strictPort',
      cwd: '../examples/vue',
      url: 'http://localhost:5304',
      reuseExistingServer: false,
    },
  ],
});
