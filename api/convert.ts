/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pptx-to-pdf/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
/**
 * `POST /api/convert`: the demo's PDF export, after DOCX Editor's `api/convert.ts`.
 *
 * The React and Vue demos post the saved deck here when the user picks Export > Presentation as
 * PDF. The conversion runs on Node.js because `@pptx-editor.dev/pptx-to-pdf` reads font files
 * through `node:fs`. The Vite development servers serve this same handler
 * (`examples/shared/pdf-export-plugin.ts`), so local runs and the deployment share one contract.
 *
 * The success body is the PDF itself (`application/pdf`). A failure is JSON with a fixed
 * `message`: raw exception text can name parts of the upload, so it never goes back.
 *
 * Limits: a 4 MiB upload and a 4 MiB PDF (the platform caps each function body at 4.5 MB), a
 * 15 second upload, a 60 second conversion, and one conversion at a time per instance. The body
 * is read before the instance lock, so a slow upload never holds the lock. A closed connection
 * cancels the conversion. Neither the upload nor the PDF is kept.
 *
 * The request must carry an `Origin` that matches the host: a browser sends `Origin` on every
 * `POST`, so a request without one is not the demo page. This keeps other pages from using the
 * function. It is not authentication: a script can send any `Origin`. Rate limits belong to the
 * host's firewall.
 */
import type { IncomingMessage, ServerResponse } from 'node:http';
import { exportPdf } from '@pptx-editor.dev/pptx-to-pdf';

const MAX_UPLOAD = 4 * 1024 * 1024;
const MAX_OUTPUT = 4 * 1024 * 1024;
const UPLOAD_MS = 15_000;
const DEADLINE_MS = 60_000;
const TOO_LARGE = { ok: false, error: 'UploadTooLarge', message: 'The upload limit is 4 MiB.' };

/** Time the function may hold the request; the conversion's own deadline is shorter. */
export const maxDuration = 90;
/** The body is the PPTX bytes; the platform must not parse it. */
export const config = { api: { bodyParser: false } };

/** One conversion per instance, so one instance never lays out two decks at once. */
let active = false;

function json(res: ServerResponse, status: number, value: Record<string, unknown>): void {
  if (res.destroyed || res.writableEnded) return;
  res.writeHead(status, { 'Content-Type': 'application/json', 'Cache-Control': 'no-store' });
  res.end(JSON.stringify(value));
}

/** The body, `'large'` past `MAX_UPLOAD`, or `'slow'` when it takes more than `UPLOAD_MS`. */
async function readBody(req: IncomingMessage): Promise<Uint8Array | 'large' | 'slow'> {
  let timer: ReturnType<typeof setTimeout> | undefined;
  const slow = new Promise<'slow'>((resolve) => {
    timer = setTimeout(() => resolve('slow'), UPLOAD_MS);
  });
  const read = (async () => {
    const chunks: Buffer[] = [];
    let length = 0;
    for await (const chunk of req) {
      const bytes = chunk as Buffer;
      length += bytes.length;
      if (length > MAX_UPLOAD) return 'large' as const;
      chunks.push(bytes);
    }
    return new Uint8Array(Buffer.concat(chunks));
  })();
  try {
    return await Promise.race([read, slow]);
  } finally {
    clearTimeout(timer);
  }
}

/** Fixed public failure responses. Never return raw exception messages. */
export function pdfFailureResponse(error: unknown): {
  status: number;
  body: Record<string, unknown>;
} {
  const name = error instanceof Error ? error.name : 'Error';
  const code = error instanceof Error && 'code' in error ? error.code : undefined;
  const fail = (status: number, message: string) => ({
    status,
    body: { ok: false, error: name, message },
  });
  if (name === 'PdfDocumentOpenError')
    return fail(400, 'The file is not a presentation this converter can open.');
  if (name === 'PdfWorkLimitError' && code === 'timeout')
    return fail(408, 'The conversion took more than 60 seconds.');
  if (
    ['RangeError', 'PdfWorkLimitError', 'PdfPageLimitError', 'PdfOutputLimitError'].includes(name)
  )
    return fail(507, 'The presentation is larger than this demo converts.');
  return fail(500, 'The conversion failed.');
}

export default async function handler(req: IncomingMessage, res: ServerResponse): Promise<void> {
  if (req.method !== 'POST') return json(res, 405, { ok: false, message: 'Use POST.' });
  const host = req.headers.host ?? '';
  const origin = req.headers.origin;
  if (!origin || (origin !== `https://${host}` && origin !== `http://${host}`))
    return json(res, 403, { ok: false, message: 'Use the same-origin demo.' });
  if (Number(req.headers['content-length']) > MAX_UPLOAD) return json(res, 413, TOO_LARGE);
  const bytes = await readBody(req);
  if (bytes === 'large') return json(res, 413, TOO_LARGE);
  if (bytes === 'slow') {
    json(res, 408, { ok: false, error: 'UploadTimeout', message: 'The upload was too slow.' });
    // Stop reading the rest of the upload.
    req.destroy?.();
    return;
  }
  if (bytes.byteLength === 0)
    return json(res, 400, { ok: false, message: 'The request has no presentation.' });
  if (active)
    return json(res, 503, {
      ok: false,
      error: 'Busy',
      message: 'A conversion is already running. Try again when it finishes.',
    });
  active = true;
  // A closed connection (a closed tab, the client's timeout) stops the conversion.
  const abort = new AbortController();
  const closed = (): void => {
    if (!res.writableFinished) abort.abort();
  };
  res.on?.('close', closed);
  try {
    // The demo exports what the editor shows, so approximations become diagnostics, not refusals.
    const result = await exportPdf(bytes, {
      fidelityPolicy: 'best-effort',
      timeoutMs: DEADLINE_MS,
      maxOutputBytes: MAX_OUTPUT,
      signal: abort.signal,
    });
    if (res.destroyed || res.writableEnded) return;
    res.writeHead(200, {
      'Content-Type': 'application/pdf',
      'Content-Length': String(result.bytes.byteLength),
      'Cache-Control': 'no-store',
      'X-Pdf-Page-Count': String(result.pageCount),
      'X-Pdf-Diagnostics': String(result.diagnostics.length),
    });
    res.end(Buffer.from(result.bytes.buffer, result.bytes.byteOffset, result.bytes.byteLength));
  } catch (error) {
    const failure = pdfFailureResponse(error);
    return json(res, failure.status, failure.body);
  } finally {
    active = false;
    res.off?.('close', closed);
  }
}
