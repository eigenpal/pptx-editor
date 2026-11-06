import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import { fileURLToPath } from 'node:url';
import { pdfExportPlugin } from '../shared/pdf-export-plugin.ts';

const pkg = (p: string): string => fileURLToPath(new URL(`../../packages/${p}`, import.meta.url));

// The demo builds against workspace SOURCE, so `bun install && bun run dev` works from a clean
// clone with no package build step. Port 5274 avoids DOCX Editor's 5173.
export default defineConfig({
  plugins: [react(), pdfExportPlugin()],
  base: process.env.VITE_BASE_PATH ?? '/',
  resolve: {
    alias: [
      {
        find: /^@pptx-editor\.dev\/core\/styles\/editor\.css$/,
        replacement: pkg('core/src/styles/editor.css'),
      },
      { find: /^@pptx-editor\.dev\/core\/(\w+)$/, replacement: pkg('core/src/$1/index.ts') },
      { find: /^@pptx-editor\.dev\/core$/, replacement: pkg('core/src/index.ts') },
      { find: /^@pptx-editor\.dev\/react$/, replacement: pkg('react/src/index.ts') },
      { find: /^@pptx-editor\.dev\/i18n$/, replacement: pkg('i18n/src/index.ts') },
      { find: /^@pptx-editor\.dev\/pro\/drawing$/, replacement: pkg('pro/src/drawing/index.ts') },
      // The development server's PDF export (`pdfExportPlugin`) runs the converter from source.
      { find: /^@pptx-editor\.dev\/pptx-to-pdf$/, replacement: pkg('pptx-to-pdf/src/index.ts') },
    ],
  },
  // Both packages resolve assets relative to their own module URL; pre-bundling would move them.
  optimizeDeps: { exclude: ['harfbuzzjs', '@docx-editor.dev/fonts'] },
  build: {
    target: 'es2022',
    rollupOptions: {
      input: {
        main: fileURLToPath(new URL('index.html', import.meta.url)),
        harness: fileURLToPath(new URL('harness.html', import.meta.url)),
      },
    },
  },
  esbuild: { target: 'es2022' },
  server: { port: 5274, strictPort: true, fs: { allow: ['../..'] } },
});
