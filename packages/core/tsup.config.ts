import { readFileSync } from 'node:fs';
import { copyFile, readFile } from 'node:fs/promises';
import { dirname, relative, resolve, sep } from 'node:path';
import { fileURLToPath } from 'node:url';
import { defineConfig, type Options } from 'tsup';

// HarfBuzz is inlined, adapted from DOCX Editor (docx-editor/packages/core/tsup.config.ts). Left
// external, the harfbuzzjs glue reaches the consumer's bundler with an `await import("module")`
// that no browser target can resolve, so Bun and esbuild builds fail. Inlined, `module` is aliased
// to `src/layout/node-module-shim.ts`, the glue's WASM URL goes through `setHarfBuzzWasmUrl`, and
// `harfbuzz.wasm` ships in `dist/`. DOCX Editor commit adf0e178.

const here = dirname(fileURLToPath(import.meta.url));

/** The package version, inlined for `openFailureReport` (`__PPTX_EDITOR_VERSION__`). */
const { version } = JSON.parse(readFileSync(resolve(here, 'package.json'), 'utf8')) as {
  version: string;
};

/** How the inlined runtime locates its binary, verbatim from the minified glue. */
export const WASM_URL_EXPRESSION = 'new URL("harfbuzz.wasm",import.meta.url).href';

/**
 * Route the runtime's WASM URL through `resolveHarfBuzzWasmBinaryUrl`, override first, the
 * original expression as the fallback argument, so asset-emitting bundlers (Vite, webpack)
 * still see the pattern they rewrite. Two guards, because either alone fails open: the count
 * guard catches a glue that changed shape, the applied guard a glue that moved.
 */
const harfBuzzWasmUrlEscapeHatch: NonNullable<Options['esbuildPlugins']>[number] = {
  name: 'harfbuzz-wasm-url-escape-hatch',
  setup(build) {
    let applied = 0;
    build.onLoad({ filter: /harfbuzzjs[/\\]dist[/\\]harfbuzz\.js$/ }, async (args) => {
      const source = await readFile(args.path, 'utf8');
      const occurrences = source.split(WASM_URL_EXPRESSION).length - 1;
      if (occurrences !== 1) {
        throw new Error(
          `expected exactly one \`${WASM_URL_EXPRESSION}\` in ${args.path}, found ${occurrences}. ` +
            'The harfbuzzjs glue changed shape; re-point this plugin before shipping.'
        );
      }
      applied += 1;
      // A path relative to the glue, so no build machine path reaches the output.
      const target = relative(
        dirname(args.path),
        resolve(here, 'src/layout/harfbuzz-wasm-binary.ts')
      );
      const resolver = JSON.stringify(`./${target.split(sep).join('/')}`);
      const contents =
        `import{resolveHarfBuzzWasmBinaryUrl as __pptxResolveHarfBuzzWasmBinaryUrl}from${resolver};` +
        // A replacement function, not a string: `$&` and friends are patterns in a string.
        source.replace(
          WASM_URL_EXPRESSION,
          () => `__pptxResolveHarfBuzzWasmBinaryUrl(${WASM_URL_EXPRESSION})`
        );
      return { contents, loader: 'js' };
    });
    build.onEnd(() => {
      if (applied !== 1) {
        throw new Error(
          `harfbuzz-wasm-url-escape-hatch patched ${applied} files, expected exactly 1. ` +
            'The glue is no longer at harfbuzzjs/dist/harfbuzz.js; re-point the onLoad filter.'
        );
      }
    });
  },
};

/** Put `harfbuzz.wasm` beside the bundle that now contains its loader. */
async function copyHarfBuzzBinary(): Promise<void> {
  const entry = fileURLToPath(import.meta.resolve('harfbuzzjs'));
  await copyFile(resolve(dirname(entry), 'harfbuzz.wasm'), resolve(here, 'dist/harfbuzz.wasm'));
}

// One entry per lane subpath (core-lane-graph.test.ts checks this table against the lanes).
export default defineConfig({
  entry: {
    index: 'src/index.ts',
    contracts: 'src/contracts/index.ts',
    store: 'src/store/index.ts',
    cascade: 'src/cascade/index.ts',
    layout: 'src/layout/index.ts',
    automation: 'src/automation/index.ts',
    output: 'src/output/index.ts',
    editor: 'src/editor/index.ts',
  },
  format: ['esm'],
  // The shared OOXML layer is a private workspace package: inline its code and its types, so the
  // published bundle and declarations never name it.
  dts: { resolve: ['@eigenpal/ooxml-core'] },
  clean: true,
  metafile: true,
  sourcemap: false,
  splitting: true,
  noExternal: ['harfbuzzjs', '@eigenpal/ooxml-core'],
  define: { __PPTX_EDITOR_VERSION__: JSON.stringify(version) },
  esbuildPlugins: [harfBuzzWasmUrlEscapeHatch],
  esbuildOptions(options) {
    options.alias = { ...options.alias, module: resolve(here, 'src/layout/node-module-shim.ts') };
  },
  onSuccess: copyHarfBuzzBinary,
});
