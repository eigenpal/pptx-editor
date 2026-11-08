import { defineConfig } from 'vite';
import vue from '@vitejs/plugin-vue';
import { fileURLToPath } from 'node:url';
import { pdfExportPlugin } from '../shared/pdf-export-plugin.ts';

const pkg = (p: string): string => fileURLToPath(new URL(`../../packages/${p}`, import.meta.url));

// The Vue twin of examples/vite. It builds against workspace SOURCE, so
// `bun install && bun run dev:vue` works from a clean clone with no package build step.
// It serves the React demo's decks and icons (`publicDir`), so both demos open the same files.
// Port 5275 is next to the React demo's 5274.
export default defineConfig({
  plugins: [vue(), pdfExportPlugin()],
  base: process.env.VITE_BASE_PATH ?? '/',
  publicDir: fileURLToPath(new URL('../vite/public', import.meta.url)),
  resolve: {
    alias: [
      {
        find: /^@pptx-editor\.dev\/core\/styles\/editor\.css$/,
        replacement: pkg('core/src/styles/editor.css'),
      },
      { find: /^@pptx-editor\.dev\/core\/(\w+)$/, replacement: pkg('core/src/$1/index.ts') },
      { find: /^@pptx-editor\.dev\/core$/, replacement: pkg('core/src/index.ts') },
      { find: /^@pptx-editor\.dev\/vue$/, replacement: pkg('vue/src/index.ts') },
      { find: /^@pptx-editor\.dev\/i18n$/, replacement: pkg('i18n/src/index.ts') },
      { find: /^@pptx-editor\.dev\/pro\/drawing$/, replacement: pkg('pro/src/drawing/index.ts') },
      { find: /^@pptx-editor\.dev\/pro\/vue$/, replacement: pkg('pro/src/vue/index.ts') },
      { find: /^@pptx-editor\.dev\/pro$/, replacement: pkg('pro/src/index.ts') },
      // The development server's PDF export (`pdfExportPlugin`) runs the converter from source.
      { find: /^@pptx-editor\.dev\/pptx-to-pdf$/, replacement: pkg('pptx-to-pdf/src/index.ts') },
    ],
  },
  // Both packages resolve assets relative to their own module URL; pre-bundling would move them.
  optimizeDeps: { exclude: ['harfbuzzjs', '@docx-editor.dev/fonts'] },
  build: { target: 'es2022' },
  // The adapter's complex parts are Vue TSX (`/** @jsxImportSource vue */`); its tsconfig keeps
  // `jsx: preserve` for the type checker, so the demo compiles their JSX with Vue's runtime.
  oxc: { target: 'es2022', jsx: { runtime: 'automatic', importSource: 'vue' } },
  server: { port: 5275, strictPort: true, fs: { allow: ['../..'] } },
});
