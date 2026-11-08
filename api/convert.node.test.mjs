/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pptx-to-pdf/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// The contract of the demo's PDF export function, driven in-process with plain request and response
// stand-ins, after DOCX Editor's `api/convert.node.test.mjs`. The handler imports the built
// package, so this runs after `bun run build:pdf` and skips itself elsewhere. DOCX Editor commit
// adf0e178.
//
//   node --test api/convert.node.test.mjs
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { existsSync } from 'node:fs';
import { readFile } from 'node:fs/promises';
import { createServer, request as httpRequest } from 'node:http';
import { Readable } from 'node:stream';

const built = existsSync(new URL('../packages/pptx-to-pdf/dist/index.js', import.meta.url));
const skip = built ? false : 'packages/pptx-to-pdf is not built; run `bun run build:pdf`';
const HOST = 'demo.example';
const ORIGIN = { origin: `https://${HOST}` };

function request({ method = 'POST', headers = {}, body } = {}) {
  const stream = body === undefined ? Readable.from([]) : Readable.from([Buffer.from(body)]);
  return Object.assign(stream, {
    method,
    url: '/api/convert',
    headers: { host: HOST, ...headers },
  });
}

/**
 * A response stand-in. With `leaveAfterStart`, the client goes away just after the handler
 * listens for `close`, before the request gets its turn to convert.
 */
function response({ leaveAfterStart = false } = {}) {
  const out = { status: 0, headers: {}, chunks: [], destroyed: false, writableEnded: false };
  const listeners = new Set();
  const disconnect = () => {
    out.destroyed = true;
    for (const listener of listeners) listener();
  };
  return {
    out,
    on(event, listener) {
      if (event !== 'close') return this;
      listeners.add(listener);
      if (leaveAfterStart) queueMicrotask(disconnect);
      return this;
    },
    off(event, listener) {
      listeners.delete(listener);
      return this;
    },
    get writableFinished() {
      return out.writableEnded;
    },
    get destroyed() {
      return out.destroyed;
    },
    get writableEnded() {
      return out.writableEnded;
    },
    writeHead(status, headers) {
      out.status = status;
      out.headers = headers;
      return this;
    },
    end(chunk) {
      if (chunk !== undefined) out.chunks.push(Buffer.from(chunk));
      out.writableEnded = true;
    },
  };
}

async function call(handler, options) {
  const res = response();
  await handler(request(options), res);
  const body = Buffer.concat(res.out.chunks);
  const json = res.out.headers['Content-Type'] === 'application/json';
  return {
    status: res.out.status,
    headers: res.out.headers,
    body,
    json: json ? JSON.parse(body.toString('utf8')) : null,
  };
}

test(
  'the function refuses anything but a same-origin POST within the limits',
  { skip },
  async () => {
    const { default: handler } = await import('./convert.ts');
    assert.equal((await call(handler, { method: 'GET', headers: ORIGIN })).status, 405);
    assert.equal((await call(handler, { headers: {} })).status, 403);
    assert.equal(
      (await call(handler, { headers: { origin: 'https://elsewhere.example' } })).status,
      403
    );
    // The platform caps a function body at 4.5 MB, so the function's own limit is 4 MiB.
    const declared = await call(handler, {
      headers: { ...ORIGIN, 'content-length': String(5 * 1024 * 1024) },
    });
    assert.equal(declared.status, 413);
    assert.equal(declared.json.message, 'The upload limit is 4 MiB.');
    // A body that does not declare its length is counted as it arrives.
    const sent = await call(handler, { headers: ORIGIN, body: Buffer.alloc(5 * 1024 * 1024) });
    assert.equal(sent.status, 413);
    const empty = await call(handler, { headers: ORIGIN, body: Buffer.alloc(0) });
    assert.equal(empty.status, 400);
    assert.equal(empty.json.message, 'The request has no presentation.');
  }
);

test('one conversion converts and one waits; a third request is refused', { skip }, async () => {
  const { default: handler } = await import('./convert.ts');
  const deck = await readFile(new URL('../examples/vite/public/sample.pptx', import.meta.url));
  // The three bodies arrive within a few ticks, long before the first conversion ends, so two
  // requests are in line when the third one asks for a turn.
  const results = await Promise.all(
    [1, 2, 3].map(() => call(handler, { headers: ORIGIN, body: deck }))
  );
  assert.deepEqual(results.map((r) => r.status).sort(), [200, 200, 503]);
  assert.equal(results.find((r) => r.status === 503).json.error, 'Busy');
  const result = results.find((r) => r.status === 200);
  assert.equal(result.headers['Content-Type'], 'application/pdf');
  assert.equal(result.body.subarray(0, 5).toString('latin1'), '%PDF-');
  assert.equal(result.headers['X-Pdf-Page-Count'], '5');
  assert.equal(result.headers['Content-Length'], String(result.body.length));
});

test(
  'a client that leaves before its turn gets no conversion, and its turn is freed',
  { skip },
  async () => {
    const { default: handler } = await import('./convert.ts');
    const deck = await readFile(new URL('../examples/vite/public/sample.pptx', import.meta.url));
    const gone = response({ leaveAfterStart: true });
    await handler(request({ headers: ORIGIN, body: deck }), gone);
    assert.equal(gone.out.status, 0, 'nothing is written to a closed connection');
    const next = await call(handler, { headers: ORIGIN, body: deck });
    assert.equal(next.status, 200);
  }
);

test(
  'a client that leaves while it waits behind a conversion gets nothing, and the line moves on',
  { skip },
  async () => {
    const { default: handler } = await import('./convert.ts');
    const deck = await readFile(new URL('../examples/vite/public/sample.pptx', import.meta.url));
    // The first request gets its turn first: both bodies are one chunk, read in order.
    const running = call(handler, { headers: ORIGIN, body: deck });
    const gone = response({ leaveAfterStart: true });
    const waiting = handler(request({ headers: ORIGIN, body: deck }), gone);
    assert.equal((await running).status, 200);
    await waiting;
    assert.equal(gone.out.status, 0, 'nothing is written to a closed connection');
    // Both turns are free again: two new requests convert side by side in line.
    const after = await Promise.all(
      [1, 2].map(() => call(handler, { headers: ORIGIN, body: deck }))
    );
    assert.deepEqual(
      after.map((r) => r.status),
      [200, 200]
    );
  }
);

test("failures answer with fixed text, never with the file's own details", { skip }, async () => {
  const { default: handler, pdfFailureResponse } = await import('./convert.ts');
  const notPptx = await call(handler, { headers: ORIGIN, body: Buffer.from('not a zip at all') });
  assert.equal(notPptx.status, 400);
  assert.equal(notPptx.json.error, 'PdfDocumentOpenError');
  assert.equal(notPptx.json.message, 'The file is not a presentation this converter can open.');

  const { PdfPageLimitError, PdfOutputLimitError, PdfWorkLimitError } =
    await import('@pptx-editor.dev/pptx-to-pdf');
  for (const error of [
    new PdfPageLimitError(1, 2),
    new PdfOutputLimitError(1, 2),
    new PdfWorkLimitError('operations'),
  ]) {
    const result = pdfFailureResponse(error);
    assert.equal(result.status, 507);
    assert.equal(result.body.error, error.name);
    assert.equal(result.body.message, 'The presentation is larger than this demo converts.');
  }
  assert.equal(pdfFailureResponse(new PdfWorkLimitError('timeout')).status, 408);
  const unknown = pdfFailureResponse(new Error('secret slide title'));
  assert.equal(unknown.status, 500);
  assert.equal(unknown.body.message, 'The conversion failed.');
});

test('a real client gets the 413 reply for a chunked upload past the limit', { skip }, async () => {
  // A socket, not a stand-in: the reply must go out before the server stops the upload.
  const { default: handler } = await import('./convert.ts');
  const server = createServer((req, res) => void handler(req, res));
  await new Promise((resolve) => server.listen(0, '127.0.0.1', resolve));
  const { port } = server.address();
  try {
    const reply = await new Promise((resolve, reject) => {
      const req = httpRequest(
        {
          host: '127.0.0.1',
          port,
          method: 'POST',
          path: '/api/convert',
          // No `Content-Length`: the body is chunked, so only the count while reading stops it.
          headers: { origin: `http://127.0.0.1:${port}`, 'transfer-encoding': 'chunked' },
        },
        (res) => {
          const chunks = [];
          res.on('data', (chunk) => chunks.push(chunk));
          res.on('end', () => resolve({ status: res.statusCode, body: Buffer.concat(chunks) }));
        }
      );
      // A lost reply fails the test; it does not hang it.
      setTimeout(() => reject(new Error('no reply within 10 s')), 10_000).unref();
      req.on('error', reject);
      const piece = Buffer.alloc(256 * 1024);
      let sent = 0;
      const more = () => {
        while (sent < 5 * 1024 * 1024) {
          sent += piece.length;
          if (!req.write(piece)) return void req.once('drain', more);
        }
        req.end();
      };
      more();
    });
    assert.equal(reply.status, 413);
    assert.equal(JSON.parse(reply.body.toString('utf8')).error, 'UploadTooLarge');
  } finally {
    server.closeAllConnections();
    await new Promise((resolve) => server.close(resolve));
  }
});
