import { defineConfig } from '@playwright/test';

// Run the same inspector tasks against both source adapters, without shared development servers.
export default defineConfig({
  testDir: '.',
  testMatch: /inspector(?:-tasks|-effects)?\.spec\.ts$/,
  outputDir: '../test-results/inspector',
  timeout: 60_000,
  workers: 1,
  use: {
    viewport: { width: 1440, height: 900 },
    acceptDownloads: true,
    actionTimeout: 10_000,
  },
  projects: [
    { name: 'react', use: { baseURL: 'http://localhost:5307' } },
    { name: 'vue', use: { baseURL: 'http://localhost:5309' } },
  ],
  webServer: [
    {
      command: 'bunx vite --port 5307 --strictPort',
      cwd: '../examples/vite',
      url: 'http://localhost:5307',
      reuseExistingServer: false,
    },
    {
      command: 'bunx vite --port 5309 --strictPort',
      cwd: '../examples/vue',
      url: 'http://localhost:5309',
      reuseExistingServer: false,
    },
  ],
});
