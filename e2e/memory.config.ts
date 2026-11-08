import { defineConfig } from '@playwright/test';

// The browser memory bench (`bun run bench:memory:browser`). It drives the React demo, and the
// Vue demo for the Vue mount scenario, each on its own server (never a reused one), so the
// numbers always come from this checkout. `MEMORY_PORT` moves both ports (React, then Vue).
const PORT = Number(process.env.MEMORY_PORT ?? 5520);
const VUE_PORT = PORT + 1;

export default defineConfig({
  testDir: '.',
  testMatch: /memory\.bench\.ts$/,
  timeout: 900_000,
  workers: 1,
  reporter: 'list',
  use: {
    baseURL: `http://localhost:${PORT}`,
    viewport: { width: 1440, height: 900 },
    acceptDownloads: true,
  },
  webServer: [
    {
      command: `bunx vite --port ${PORT} --strictPort`,
      cwd: '../examples/vite',
      url: `http://localhost:${PORT}`,
      reuseExistingServer: false,
      timeout: 120_000,
    },
    {
      command: `bunx vite --port ${VUE_PORT} --strictPort`,
      cwd: '../examples/vue',
      url: `http://localhost:${VUE_PORT}`,
      reuseExistingServer: false,
      timeout: 120_000,
    },
  ],
});
