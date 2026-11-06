import { defineConfig } from 'tsup';

// `core` stays external and a peer: the engine must resolve to ONE copy (shaper, caches).
export default defineConfig({
  entry: { index: 'src/index.ts' },
  format: ['esm'],
  dts: true,
  clean: true,
  metafile: true,
  sourcemap: false,
  external: ['react', 'react-dom', /^@pptx-editor\.dev\//],
});
