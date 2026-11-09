import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import { sharedPublicDir, sourceAliases } from '../shared/collaboration/vite-aliases.ts';

// Port 5278 sits after the peer-to-peer example's 5277. The room server listens on 1234.
export default defineConfig({
  plugins: [react()],
  publicDir: sharedPublicDir,
  resolve: { alias: sourceAliases },
  // Both packages resolve assets relative to their own module URL; pre-bundling would move them.
  optimizeDeps: { exclude: ['harfbuzzjs', '@docx-editor.dev/fonts'] },
  build: { target: 'es2022' },
  server: { port: 5278, strictPort: true, fs: { allow: ['../..'] } },
});
