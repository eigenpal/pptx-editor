import { defineConfig } from 'tsup';

// `core` stays external and a peer, as in the adapters: the engine must resolve to ONE copy.
// Minified and without source maps, as DOCX Editor builds its pro package.
export default defineConfig({
  entry: { index: 'src/index.ts' },
  format: ['esm'],
  dts: true,
  clean: true,
  metafile: true,
  sourcemap: false,
  minify: true,
  external: ['yjs', /^y-protocols/, /^@pptx-editor\.dev\//],
});
