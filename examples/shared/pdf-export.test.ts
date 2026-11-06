// The demo's PDF export client: each failure of the server function becomes a translated notice
// key, never the server's English text or the browser's raw error.
import { afterEach, describe, expect, test } from 'bun:test';
import type { PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import {
  PDF_MAX_UPLOAD,
  PdfExportFailure,
  convertToPdf,
  exportSlides,
  watchPdfExport,
} from './demo-files.ts';

const realFetch = globalThis.fetch;
afterEach(() => {
  globalThis.fetch = realFetch;
});

function answer(status: number, type: string, body = ''): void {
  globalThis.fetch = (async () =>
    new Response(body, { status, headers: { 'Content-Type': type } })) as unknown as typeof fetch;
}

async function failure(deck = new Uint8Array(8)): Promise<string> {
  try {
    await convertToPdf(deck);
  } catch (e) {
    expect(e).toBeInstanceOf(PdfExportFailure);
    return (e as PdfExportFailure).key;
  }
  throw new Error('convertToPdf did not fail');
}

describe('convertToPdf', () => {
  test('returns the PDF of a successful conversion', async () => {
    answer(200, 'application/pdf', '%PDF-1.7');
    const pdf = await convertToPdf(new Uint8Array(8));
    expect(await pdf.text()).toBe('%PDF-1.7');
  });

  test('maps each failure status to a notice key', async () => {
    answer(413, 'text/plain', 'Request Entity Too Large');
    expect(await failure()).toBe('app.pdfTooLarge');
    answer(507, 'application/json', '{"message":"The presentation is larger"}');
    expect(await failure()).toBe('app.pdfTooComplex');
    answer(503, 'application/json', '{"message":"A conversion is already running"}');
    expect(await failure()).toBe('app.pdfBusy');
    answer(500, 'application/json', '{"message":"The conversion failed."}');
    expect(await failure()).toBe('app.pdfExportFailed');
    // A page that is not the function (a static host with no `/api/convert`).
    answer(200, 'text/html', '<!doctype html>');
    expect(await failure()).toBe('app.pdfExportFailed');
  });

  test('a network failure and a deck past the limit fail without the raw error', async () => {
    globalThis.fetch = (async () => {
      throw new TypeError('Failed to fetch');
    }) as unknown as typeof fetch;
    expect(await failure()).toBe('app.pdfExportFailed');
    let sent = false;
    globalThis.fetch = (async () => {
      sent = true;
      return new Response('');
    }) as unknown as typeof fetch;
    expect(await failure(new Uint8Array(PDF_MAX_UPLOAD + 1))).toBe('app.pdfTooLarge');
    expect(sent).toBe(false);
  });
});

/** An editor with a deck, as far as `exportSlides` reads it. */
function fakeEditor() {
  const rejected: string[] = [];
  const editor = {
    session: {},
    snapshot: () => ({ currentSlide: 0 }),
    save: () => new Uint8Array(8),
    reject: (reason: string) => (rejected.push(reason), { ok: false }),
  } as unknown as PptxEditorInstance;
  return { editor, rejected };
}

describe('exportSlides as PDF', () => {
  test('a second pick while the PDF is made sends nothing, and the busy state follows', async () => {
    let finish!: (r: Response) => void;
    let requests = 0;
    globalThis.fetch = (() => {
      requests++;
      return new Promise<Response>((resolve) => (finish = resolve));
    }) as unknown as typeof fetch;
    const states: boolean[] = [];
    const stop = watchPdfExport((busy) => states.push(busy));
    const { editor, rejected } = fakeEditor();
    const first = exportSlides(editor, 'pdf', 'deck');
    await exportSlides(editor, 'pdf', 'deck');
    expect(requests).toBe(1);
    finish(new Response('', { status: 500, headers: { 'Content-Type': 'application/json' } }));
    await first;
    stop();
    expect(states).toEqual([false, true, false]);
    // The failure shows as the editor's notice, by its key.
    expect(rejected).toEqual(['app.pdfExportFailed']);
  });

  test('each failure reaches the notice as its own key', async () => {
    answer(503, 'application/json', '{}');
    const { editor, rejected } = fakeEditor();
    await exportSlides(editor, 'pdf', 'deck');
    expect(rejected).toEqual(['app.pdfBusy']);
  });
});
