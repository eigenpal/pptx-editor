import { defineConfig } from '@playwright/test';

/** The React demo (`examples/vite`): the `baseURL` of the browser spec. */
const PORT = 5297;
/** The Vue demo (`examples/vue`), for the specs that run in both adapters (`VUE_DEMO_URL`). */
const VUE_PORT = 5296;

// The demo is the system under test: the same app a person runs with `bun run dev`. It serves
// the workspace source, so no package build runs first. Only what needs real layout, a real
// canvas, or real frames runs here: `browser.spec.ts`, and `document-refresh.spec.ts` in both
// adapters. Every other demo flow is a happy-dom test in `test:dom`.
export default defineConfig({
  testDir: '.',
  testMatch: /.*\.spec\.ts$/,
  // These tasks use their own React and Vue servers in inspector.config.ts.
  testIgnore: /inspector(?:-tasks|-effects|-accessibility)?\.spec\.ts$/,
  timeout: 60_000,
  use: {
    baseURL: `http://localhost:${PORT}`,
    viewport: { width: 1440, height: 900 },
    acceptDownloads: true,
  },
  // Never reuse a running server: another worktree may serve its own code on this port, and the
  // run would then test that code. A busy port fails the run instead.
  metadata: { vueUrl: `http://localhost:${VUE_PORT}` },
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
