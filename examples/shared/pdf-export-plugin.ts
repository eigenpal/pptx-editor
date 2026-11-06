// Serve `POST /api/convert` (the root `api/convert.ts` function) from the React and Vue
// development servers, after DOCX Editor's `examples/shared/pdf-export-plugin.ts`. The demo then
// exports PDF locally through the same handler as the deployment. The demos' aliases point
// `@pptx-editor.dev/pptx-to-pdf` at its source, so no package build step is needed.

import type { Plugin } from 'vite';
import { fileURLToPath } from 'node:url';

export function pdfExportPlugin(): Plugin {
  return {
    name: 'pptx-pdf-export',
    configureServer(server) {
      server.middlewares.use((request, response, next) => {
        if (request.url?.split('?')[0] !== '/api/convert') return next();
        const handlerPath = fileURLToPath(new URL('../../api/convert.ts', import.meta.url));
        void server
          .ssrLoadModule(handlerPath)
          .then(({ default: handler }) => handler(request, response))
          .catch((error: unknown) => {
            server.config.logger.error(String(error));
            if (response.headersSent) return response.end();
            response.writeHead(500, { 'Content-Type': 'application/json' });
            response.end(
              JSON.stringify({ ok: false, message: 'The PDF converter did not start.' })
            );
          });
      });
    },
  };
}
