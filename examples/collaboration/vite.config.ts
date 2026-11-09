import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import { sharedPublicDir, sourceAliases } from '../shared/collaboration/vite-aliases.ts';

// Port 5277 sits after the chat example's 5276. The room key needs `crypto.subtle`, which a
// browser offers on `localhost` and on HTTPS only.
export default defineConfig({
  plugins: [react()],
  publicDir: sharedPublicDir,
  resolve: { alias: sourceAliases },
  // Both packages resolve assets relative to their own module URL; pre-bundling would move them.
  optimizeDeps: { exclude: ['harfbuzzjs', '@docx-editor.dev/fonts'] },
  build: { target: 'es2022' },
  server: { port: 5277, strictPort: true, fs: { allow: ['../..'] } },
});
