import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import { fileURLToPath } from 'node:url';

const pkg = (p: string): string => fileURLToPath(new URL(`../../packages/${p}`, import.meta.url));

// The chat example builds against workspace source, as the Vite demo does, and reuses that
// demo's sample decks. Port 5276 sits near the demo's 5274.
export default defineConfig({
  plugins: [react()],
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
      { find: /^@pptx-editor\.dev\/react$/, replacement: pkg('react/src/index.ts') },
      { find: /^@pptx-editor\.dev\/i18n$/, replacement: pkg('i18n/src/index.ts') },
    ],
  },
  // Both packages resolve assets relative to their own module URL; pre-bundling would move them.
  optimizeDeps: { exclude: ['harfbuzzjs', '@docx-editor.dev/fonts'] },
  build: { target: 'es2022' },
  server: { port: 5276, strictPort: true, fs: { allow: ['../..'] } },
});
