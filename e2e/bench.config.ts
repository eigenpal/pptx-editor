import { defineConfig } from '@playwright/test';

// The browser editing bench (`bun run bench:edit:browser`). Same demo as the e2e suite, but its
// own file pattern, so `bun run test:e2e` never runs a benchmark, and its own server (never a
// reused one), so the numbers always come from this checkout.
const PORT = Number(process.env.BENCH_PORT ?? 5298);

export default defineConfig({
  testDir: '.',
  testMatch: /.*\.bench\.ts$/,
  timeout: 600_000,
  workers: 1,
  reporter: 'list',
  use: {
    baseURL: `http://localhost:${PORT}`,
    viewport: { width: 1440, height: 900 },
  },
  webServer: {
    command: `bunx vite --port ${PORT} --strictPort`,
    cwd: '../examples/vite',
    url: `http://localhost:${PORT}`,
    reuseExistingServer: false,
    timeout: 120_000,
  },
});
