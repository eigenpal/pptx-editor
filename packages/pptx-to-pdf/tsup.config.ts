import { defineConfig } from 'tsup';

// `core` stays external and a peer, as in the adapters: the engine must resolve to ONE copy.
export default defineConfig({
  entry: { index: 'src/index.ts' },
  format: ['esm'],
  dts: true,
  clean: true,
  metafile: true,
  sourcemap: true,
  external: ['pdf-lib', 'fontkit', /^@pptx-editor\.dev\//],
});
