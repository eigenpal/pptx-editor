import { readFileSync } from 'node:fs';
import { defineConfig, type Options } from 'tsup';

// The license header opens every JS and .d.ts file in dist as a legal comment (`/*!`), which minifiers keep.
const banner = `/*!\n${readFileSync(new URL('./license-header.txt', import.meta.url), 'utf8').trim()}\n*/`;

// Minified and without source maps, so the published files do not carry the source.
// `core` and the adapters stay external and peers: the engine must resolve to ONE copy.
const shared: Options = {
  format: ['esm'],
  sourcemap: false,
  minify: true,
  banner: { js: banner },
  external: ['yjs', /^y-protocols/, /^@pptx-editor\.dev\//, 'react', 'react-dom', 'vue'],
};

// One JS build for every entry, so the entries share one copy of each module (the review module
// and the license key). Each `.tsx` file names its JSX runtime with a `@jsxImportSource` pragma,
// which esbuild reads per file. TypeScript cannot switch the JSX mode per file, so the
// declarations come from two programs, as in DOCX Editor's pro package: the Vue entry's from
// `tsconfig.vue.json`. The package `build` script empties `dist/` first.
export default defineConfig([
  {
    ...shared,
    entry: {
      index: 'src/index.ts',
      drawing: 'src/drawing/index.ts',
      react: 'src/react/index.ts',
      vue: 'src/vue/index.ts',
    },
    metafile: true,
    dts: {
      banner,
      entry: {
        index: 'src/index.ts',
        drawing: 'src/drawing/index.ts',
        react: 'src/react/index.ts',
      },
    },
    esbuildOptions(options) {
      options.jsx = 'automatic';
    },
  },
  {
    ...shared,
    entry: { vue: 'src/vue/index.ts' },
    tsconfig: 'tsconfig.vue.json',
    dts: { banner, only: true },
  },
]);
