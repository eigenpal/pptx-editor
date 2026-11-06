/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pptx-to-pdf/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// The contract of the demo's PDF export function, driven in-process with plain request and
// response stand-ins, after DOCX Editor's `api/convert.node.test.mjs`. The handler imports the
// built package, so this runs after `bun run build:pdf` and skips itself elsewhere.
//
//   node --test api/convert.node.test.mjs
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { existsSync } from 'node:fs';
import { readFile } from 'node:fs/promises';
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

function response() {
  const out = { status: 0, headers: {}, chunks: [], destroyed: false, writableEnded: false };
  return {
    out,
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
    assert.equal(
      (await call(handler, { headers: { ...ORIGIN, 'content-length': String(21 * 1024 * 1024) } }))
        .status,
      413
    );
    const empty = await call(handler, { headers: ORIGIN, body: Buffer.alloc(0) });
    assert.equal(empty.status, 400);
    assert.equal(empty.json.message, 'The request has no presentation.');
  }
);

test('one conversion at a time per instance, then a PDF of every slide', { skip }, async () => {
  const { default: handler } = await import('./convert.ts');
  const deck = await readFile(new URL('../examples/vite/public/sample.pptx', import.meta.url));
  const first = call(handler, { headers: ORIGIN, body: deck });
  // Arrives while the first request still reads its body.
  const second = await call(handler, { headers: ORIGIN, body: deck });
  assert.equal(second.status, 503);
  const result = await first;
  assert.equal(result.status, 200);
  assert.equal(result.headers['Content-Type'], 'application/pdf');
  assert.equal(result.body.subarray(0, 5).toString('latin1'), '%PDF-');
  assert.equal(result.headers['X-Pdf-Page-Count'], '5');
  assert.equal(result.headers['Content-Length'], String(result.body.length));
});

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
