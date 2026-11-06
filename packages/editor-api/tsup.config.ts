import { readFileSync } from 'node:fs';
import { defineConfig } from 'tsup';

// The license header opens every JS and .d.ts file in dist as a legal comment (`/*!`), which minifiers keep.
const banner = `/*!\n${readFileSync(new URL('./license-header.txt', import.meta.url), 'utf8').trim()}\n*/`;

// Minified and without source maps, so the published files do not carry the source.
// `core` stays external and a peer, as in the adapters: the engine must resolve to ONE copy.
export default defineConfig({
  entry: { index: 'src/index.ts', browser: 'src/browser.ts' },
  format: ['esm'],
  dts: { banner },
  clean: true,
  metafile: true,
  sourcemap: false,
  minify: true,
  banner: { js: banner },
  external: [/^@pptx-editor\.dev\//],
});
