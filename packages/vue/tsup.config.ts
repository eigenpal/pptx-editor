import { defineConfig } from 'tsup';

// `core` stays external and a peer, as in the React adapter: the engine must resolve to ONE
// copy (shaper, caches). Simple parts use `h()`; complex parts are `.tsx` with Vue's JSX
// runtime, as in DOCX Editor's Vue adapter.
export default defineConfig({
  entry: { index: 'src/index.ts' },
  format: ['esm'],
  dts: true,
  clean: true,
  metafile: true,
  sourcemap: false,
  external: ['vue', /^@pptx-editor\.dev\//],
  esbuildOptions(options) {
    options.jsx = 'automatic';
    options.jsxImportSource = 'vue';
  },
});
