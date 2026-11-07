import { defineConfig } from '@playwright/test';

const browserName = process.env.HANG_BROWSER ?? 'chromium';
if (!['chromium', 'firefox', 'webkit'].includes(browserName))
  throw new Error('Invalid HANG_BROWSER');

// Each adapter runs the same adversarial interaction sequence against this checkout.
export default defineConfig({
  testDir: '.',
  outputDir: `../test-results/interactions-${browserName}`,
  testMatch: 'interactions.stress.ts',
  timeout: 90_000,
  workers: 1,
  use: {
    browserName: browserName as 'chromium' | 'firefox' | 'webkit',
    viewport: { width: 1440, height: 900 },
    actionTimeout: 10_000,
  },
  projects: [
    { name: 'react', use: { baseURL: 'http://localhost:5297' } },
    { name: 'vue', use: { baseURL: 'http://localhost:5299' } },
  ],
  webServer: [
    {
      command: 'bunx vite --port 5297 --strictPort',
      cwd: '../examples/vite',
      url: 'http://localhost:5297',
      reuseExistingServer: false,
    },
    {
      command: 'bunx vite --port 5299 --strictPort',
      cwd: '../examples/vue',
      url: 'http://localhost:5299',
      reuseExistingServer: false,
    },
  ],
});
