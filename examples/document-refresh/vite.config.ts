import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import { fileURLToPath } from 'node:url';
import { samplePptx, sampleUpdate } from './sample-presentation.ts';
import { agentPptx, agentStepParam } from './agent-script.ts';

const pkg = (p: string): string => fileURLToPath(new URL(`../../packages/${p}`, import.meta.url));

// The example builds against workspace source, as the Vite demo does. A middleware acts as the
// processing server: it returns controlled sample files, never an edit of the uploaded bytes.
// Port 5279 follows the other examples.
export default defineConfig({
  plugins: [
    react(),
    {
      name: 'mock-presentation-server',
      configureServer(server) {
        server.middlewares.use(async (req, res, next) => {
          if (req.url === '/api/presentation' && req.method === 'GET') {
            res.setHeader('Content-Type', 'application/octet-stream');
            res.end(samplePptx());
            return;
          }
          const url = new URL(req.url ?? '/', 'http://localhost');
          // The replay page: the presentation after `step` steps of the scripted agent.
          if (url.pathname === '/api/agent' && req.method === 'GET') {
            const step = agentStepParam(url.searchParams.get('step'));
            if (step === null) {
              res.statusCode = 400;
              res.end('Invalid step');
              return;
            }
            res.setHeader('Content-Type', 'application/octet-stream');
            res.end(agentPptx(step));
            return;
          }
          if (url.pathname !== '/api/update' || req.method !== 'POST') {
            next();
            return;
          }
          const submissionId = req.headers['x-submission-id'];
          if (typeof submissionId !== 'string' || !submissionId) {
            res.statusCode = 400;
            res.end('Missing submission ID');
            return;
          }
          const round = Number(url.searchParams.get('round'));
          const sequence = Number(url.searchParams.get('sequence'));
          if (
            !Number.isSafeInteger(round) ||
            round < 1 ||
            round > 10000 ||
            ![1, 2].includes(sequence)
          ) {
            res.statusCode = 400;
            res.end('Invalid sample request');
            return;
          }
          // Discard the uploaded presentation. This mock generates controlled fixtures only.
          req.resume();
          await new Promise((resolve) => setTimeout(resolve, 1500));
          if (res.destroyed) return;
          res.setHeader('Content-Type', 'application/json');
          res.end(JSON.stringify(sampleUpdate(submissionId, round, sequence)));
        });
      },
    },
  ],
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
      { find: /^@pptx-editor\.dev\/pro$/, replacement: pkg('pro/src/index.ts') },
    ],
  },
  // Both packages resolve assets relative to their own module URL; pre-bundling would move them.
  optimizeDeps: { exclude: ['harfbuzzjs', '@docx-editor.dev/fonts'] },
  build: {
    target: 'es2022',
    rollupOptions: {
      input: {
        main: fileURLToPath(new URL('index.html', import.meta.url)),
        agent: fileURLToPath(new URL('agent.html', import.meta.url)),
      },
    },
  },
  server: { port: 5279, strictPort: true, fs: { allow: ['../..'] } },
});
