import { defineConfig } from 'tsup';

// `core` stays external and a peer, as in the adapters: the engine must resolve to ONE copy.
export default defineConfig({
  entry: { index: 'src/index.ts', browser: 'src/browser.ts' },
  format: ['esm'],
  dts: true,
  clean: true,
  metafile: true,
  sourcemap: true,
  external: [/^@pptx-editor\.dev\//],
});
