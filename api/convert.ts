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
 * 15 second upload, and a 40 second conversion. One instance can get parallel requests, so it
 * converts one deck at a time and keeps one more request waiting. A third request gets 503. The
 * body is read before the request waits, so a slow upload never holds a turn. A closed connection
 * cancels its conversion. A request that closes while it waits keeps its turn until the
 * conversion before it ends, and then gives it up. Neither the upload nor the PDF is kept.
 *
 * The time budget: 15 s of upload, at most 40 s of waiting, and 40 s of conversion make 95 s. A
 * request that waits longer gets 503, whatever the conversion before it does. The conversion
 * checks its deadline between slides, so `maxDuration` keeps 25 s more for a slow slide. The
 * browser waits 130 s, longer than the function can run.
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
const DEADLINE_MS = 40_000;
const WAIT_MS = 40_000;
/** The conversion that runs and the one that waits. */
const MAX_IN_FLIGHT = 2;
const TOO_LARGE = { ok: false, error: 'UploadTooLarge', message: 'The upload limit is 4 MiB.' };

/** Time the function may hold the request. See the time budget above. */
export const maxDuration = 120;

/** The end of the last conversion in line on this instance. */
let line: Promise<void> = Promise.resolve();
/** The requests that convert or wait on this instance. */
let inFlight = 0;

/**
 * Run `work` after the conversions before it, so one instance never lays out two decks at once.
 * `'busy'` when `MAX_IN_FLIGHT` requests already convert or wait, or after `WAIT_MS` of waiting.
 */
async function inTurn<T>(work: () => Promise<T>): Promise<T | 'busy'> {
  if (inFlight >= MAX_IN_FLIGHT) return 'busy';
  inFlight++;
  const before = line;
  let done!: () => void;
  line = new Promise<void>((resolve) => (done = resolve));
  let timer: ReturnType<typeof setTimeout> | undefined;
  const turn = await Promise.race([
    before.then(() => true),
    new Promise<false>((resolve) => (timer = setTimeout(() => resolve(false), WAIT_MS))),
  ]);
  clearTimeout(timer);
  if (!turn) {
    // Give up the wait. The line stays in order: the next request starts after the one before.
    inFlight--;
    void before.then(done);
    return 'busy';
  }
  try {
    return await work();
  } finally {
    inFlight--;
    done();
  }
}

function json(res: ServerResponse, status: number, value: Record<string, unknown>): void {
  if (res.destroyed || res.writableEnded) return;
  res.writeHead(status, { 'Content-Type': 'application/json', 'Cache-Control': 'no-store' });
  res.end(JSON.stringify(value));
}

/** How long, and how much more, a refused upload is read before its connection closes. */
const DISCARD_MS = 5_000;
const DISCARD_BYTES = 2 * MAX_UPLOAD;

/**
 * Refuse an upload that may still arrive. The reply goes out first, and the rest of the upload
 * is read and thrown away, so the client can finish its upload and read the reply. A socket that
 * closes with unread data resets the connection, and the client can then lose the reply. After
 * `DISCARD_MS` or `DISCARD_BYTES` the connection closes anyway.
 */
function refuseUpload(
  req: IncomingMessage,
  res: ServerResponse,
  status: number,
  value: Record<string, unknown>
): void {
  json(res, status, value);
  let discarded = 0;
  const stop = (): void => {
    clearTimeout(timer);
    req.destroy?.();
  };
  const timer = setTimeout(stop, DISCARD_MS);
  timer.unref?.();
  req.on('data', (chunk: Buffer) => {
    discarded += chunk.length;
    if (discarded > DISCARD_BYTES) stop();
  });
  req.once('end', () => clearTimeout(timer));
  req.resume();
}

/**
 * The body; `'large'` past `MAX_UPLOAD`; `'slow'` when it takes more than `UPLOAD_MS`; or
 * `'aborted'` when the client stops the upload. It pauses the request and does not destroy it,
 * so the refusal can still reach the client (`refuseUpload`).
 */
function readBody(req: IncomingMessage): Promise<Uint8Array | 'large' | 'slow' | 'aborted'> {
  return new Promise((resolve) => {
    const chunks: Buffer[] = [];
    let length = 0;
    let settled = false;
    const finish = (value: Uint8Array | 'large' | 'slow' | 'aborted'): void => {
      if (settled) return;
      settled = true;
      clearTimeout(timer);
      req.off('data', onData);
      req.off('end', onEnd);
      req.off('error', onAbort);
      req.off('close', onClose);
      if (typeof value === 'string') req.pause();
      resolve(value);
    };
    const onData = (chunk: Buffer): void => {
      length += chunk.length;
      if (length > MAX_UPLOAD) finish('large');
      else chunks.push(chunk);
    };
    const onEnd = (): void => finish(new Uint8Array(Buffer.concat(chunks)));
    const onAbort = (): void => finish('aborted');
    // A request that closes before its end lost its client.
    const onClose = (): void =>
      finish(req.readableEnded ? new Uint8Array(Buffer.concat(chunks)) : 'aborted');
    const timer = setTimeout(() => finish('slow'), UPLOAD_MS);
    req.on('data', onData);
    req.on('end', onEnd);
    req.on('error', onAbort);
    req.on('close', onClose);
  });
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
    return fail(408, 'The conversion took more than 40 seconds.');
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
  if (bytes === 'large') return refuseUpload(req, res, 413, TOO_LARGE);
  if (bytes === 'slow')
    return refuseUpload(req, res, 408, {
      ok: false,
      error: 'UploadTimeout',
      message: 'The upload was too slow.',
    });
  if (bytes === 'aborted') return;
  if (bytes.byteLength === 0)
    return json(res, 400, { ok: false, message: 'The request has no presentation.' });
  // A closed connection (a closed tab, the client's timeout) stops the conversion.
  const abort = new AbortController();
  const closed = (): void => {
    if (!res.writableFinished) abort.abort();
  };
  res.on?.('close', closed);
  try {
    const result = await inTurn(() => {
      // The client left while the request waited.
      if (abort.signal.aborted) return Promise.resolve(null);
      // The demo exports what the editor shows, so approximations become diagnostics.
      return exportPdf(bytes, {
        fidelityPolicy: 'best-effort',
        timeoutMs: DEADLINE_MS,
        maxOutputBytes: MAX_OUTPUT,
        signal: abort.signal,
      });
    });
    if (result === 'busy')
      return json(res, 503, {
        ok: false,
        error: 'Busy',
        message: 'Too many conversions are running. Try again soon.',
      });
    if (result === null) return;
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
    res.off?.('close', closed);
  }
}
