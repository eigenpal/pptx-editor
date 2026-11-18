import { defineConfig } from '@playwright/test';

/** The React demo (`examples/vite`): the `baseURL` of the browser spec. */
const PORT = 5297;
/** Production builds of the React and Vue demos, for `document-refresh.spec.ts`. */
const REACT_BUILD_PORT = 5294;
const VUE_BUILD_PORT = 5293;

/**
 * Build a demo and serve the build with `vite preview`. React's development build records
 * performance marks and async stack tasks for every component, which makes one refresh
 * render take several times longer, so the long-task checks measure a production build.
 */
const preview = (port: number) =>
  `bunx vite build --outDir dist/e2e --emptyOutDir --logLevel warn && ` +
  `bunx vite preview --outDir dist/e2e --port ${port} --strictPort`;

// The demo is the system under test: the same app a person runs with `bun run dev`. It serves
// the workspace source, so no package build runs first. Only what needs real layout, a real
// canvas, or real frames runs here: `browser.spec.ts`, and `document-refresh.spec.ts` in both
// adapters. `document-refresh.spec.ts` uses production builds of the demos, because it checks
// for long tasks. Every other demo flow is a happy-dom test in `test:dom`.
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
  metadata: {
    reactBuildUrl: `http://localhost:${REACT_BUILD_PORT}`,
    vueBuildUrl: `http://localhost:${VUE_BUILD_PORT}`,
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
      command: preview(REACT_BUILD_PORT),
      cwd: '../examples/vite',
      url: `http://localhost:${REACT_BUILD_PORT}`,
      reuseExistingServer: false,
      timeout: 180_000,
    },
    {
      command: preview(VUE_BUILD_PORT),
      cwd: '../examples/vue',
      url: `http://localhost:${VUE_BUILD_PORT}`,
      reuseExistingServer: false,
      timeout: 180_000,
    },
  ],
});
