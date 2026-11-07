import { defineConfig } from '@playwright/test';

const browserName = process.env.HANG_BROWSER ?? 'chromium';
if (!['chromium', 'firefox', 'webkit'].includes(browserName))
  throw new Error('Invalid HANG_BROWSER');

export default defineConfig({
  outputDir: `../test-results/hang-${browserName}`,
  testDir: '.',
  testMatch: /hang\.stress\.ts$/,
  timeout: 60_000,
  workers: 1,
  reporter: 'list',
  use: {
    browserName: browserName as 'chromium' | 'firefox' | 'webkit',
    baseURL: 'http://localhost:5301',
    viewport: { width: 1440, height: 900 },
  },
  webServer: {
    command: 'bunx vite --port 5301 --strictPort',
    cwd: '../examples/vite',
    url: 'http://localhost:5301',
    reuseExistingServer: false,
  },
});
