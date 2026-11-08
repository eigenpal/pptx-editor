import { defineConfig } from '@playwright/test';

// The browser editing bench (`bun run bench:edit:browser`). Same demo as the e2e suite, but its
// own file pattern, so `bun run test:e2e` never runs a benchmark, and its own server (never a
// reused one), so the numbers always come from this checkout. `BENCH_BASE_URL` points the
// benches at a demo that already runs instead, such as a production build served by
// `vite preview`, whose React runs without its development checks.
const PORT = Number(process.env.BENCH_PORT ?? 5298);
const BASE_URL = process.env.BENCH_BASE_URL;

export default defineConfig({
  testDir: '.',
  testMatch: /.*\.bench\.ts$/,
  timeout: 600_000,
  workers: 1,
  reporter: 'list',
  use: {
    baseURL: BASE_URL ?? `http://localhost:${PORT}`,
    viewport: { width: 1440, height: 900 },
  },
  ...(BASE_URL
    ? {}
    : {
        webServer: {
          command: `bunx vite --port ${PORT} --strictPort`,
          cwd: '../examples/vite',
          url: `http://localhost:${PORT}`,
          reuseExistingServer: false,
          timeout: 120_000,
        },
      }),
});
