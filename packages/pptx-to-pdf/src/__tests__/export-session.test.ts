/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pptx-to-pdf/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
import { fontConfiguration } from '../../../../scripts/fixtures/font-configuration.ts';
// `openDocumentForExport` and `exportPdfFrom`: one layout, many exports (DOCX Editor's
// reusable session).

import { missingFontsDeck } from '../../../core/src/editor/__tests__/font-test-helpers.ts';
import { describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import { getDocument } from 'pdfjs-dist/legacy/build/pdf.mjs';
import { buildPptx, para, run, shape } from '../../../../scripts/fixtures/builder.ts';
import { sessionState } from '../open-session.ts';
import {
  ExportResourceError,
  exportPdf,
  exportPdfFrom,
  openDocumentForExport,
  PdfFidelityError,
  PdfOutputLimitError,
  PdfWorkLimitError,
  type PdfExportOptions,
  type PdfExportSession,
  type PdfProjectionOptions,
} from '../index.ts';

const FIXTURES = join(import.meta.dir, '../../../../fixtures');
const deck = (name: string): Uint8Array =>
  new Uint8Array(readFileSync(join(FIXTURES, `${name}.pptx`)));
const BEST = { fidelityPolicy: 'best-effort' } as const;

/** Four text slides; the second and fourth are hidden. No fixture deck has a hidden slide. */
const withHidden = (): Uint8Array =>
  buildPptx({
    slides: [1, 2, 3, 4].map((n) => ({
      layout: 'blank' as const,
      hidden: n % 2 === 0,
      shapes: shape({
        id: 2,
        x: 914_400,
        y: 914_400,
        cx: 5_486_400,
        cy: 914_400,
        text: para(run(`Slide ${n}`)),
      }),
    })),
  });

async function open(name: string | Uint8Array, signal?: AbortSignal): Promise<PdfExportSession> {
  const bytes = typeof name === 'string' ? deck(name) : name;
  const opened = await openDocumentForExport(bytes, signal ? { signal } : {});
  if (!opened.ok) throw new Error(opened.reason);
  return opened.session;
}

describe('openDocumentForExport and exportPdfFrom', () => {
  test('two exports from one session equal a one-shot export, without reopening', async () => {
    const session = await open('text-runs');
    try {
      const first = await exportPdfFrom(session, BEST);
      const second = await exportPdfFrom(session, BEST);
      const once = await exportPdf(deck('text-runs'), BEST);
      expect(Buffer.from(first.bytes).equals(Buffer.from(once.bytes))).toBe(true);
      expect(Buffer.from(second.bytes).equals(Buffer.from(once.bytes))).toBe(true);
      expect(first.timings.openMs).toBe(0);
      expect(second.timings.openMs).toBe(0);
      expect(first.diagnostics).toEqual(once.diagnostics);
      expect(first.fontResolution).toEqual(once.fontResolution);
      expect(first.fontResolution).toBe(session.fontResolution);
      // Each result owns its bytes.
      first.bytes[0] = 0;
      expect(second.bytes[0]).toBe(0x25);
    } finally {
      session.dispose();
    }
  });

  test('each slide is laid out once per session', async () => {
    const session = await open('preset-shapes');
    try {
      const document = sessionState(session)!.document;
      const before = document.layout(0);
      expect(before).toBeDefined();
      await exportPdfFrom(session, BEST);
      await exportPdfFrom(session, BEST);
      expect(document.layout(0)).toBe(before!);
    } finally {
      session.dispose();
    }
  });

  test('projections choose the slides, and slideNumbers maps pages to slides', async () => {
    const session = await open(withHidden());
    try {
      const shown = await exportPdfFrom(session, BEST);
      const all = await exportPdfFrom(session, { ...BEST, includeHiddenSlides: true });
      expect(shown.slideNumbers).toEqual([1, 3]);
      expect(shown.pageCount).toBe(2);
      expect(all.slideNumbers).toEqual([1, 2, 3, 4]);
      for (const d of all.diagnostics)
        if (d.pageIndex !== undefined) expect(all.slideNumbers[d.pageIndex]).toBe(d.slideNumber!);
    } finally {
      session.dispose();
    }
  });

  test('a failed or cancelled export leaves the session usable', async () => {
    const session = await open('fills-lines');
    try {
      await expect(exportPdfFrom(session)).rejects.toBeInstanceOf(PdfFidelityError);
      const controller = new AbortController();
      controller.abort(new Error('stop'));
      const cancelled = exportPdfFrom(session, { ...BEST, signal: controller.signal });
      await expect(cancelled).rejects.toMatchObject({ code: 'aborted' });
      await expect(cancelled).rejects.toBeInstanceOf(PdfWorkLimitError);
      const result = await exportPdfFrom(session, BEST);
      expect(result.pageCount).toBeGreaterThan(0);
    } finally {
      session.dispose();
    }
  });

  test('a disposed session and an aborted session signal refuse exports', async () => {
    const session = await open('text-runs');
    session.dispose();
    session.dispose();
    await expect(exportPdfFrom(session, BEST)).rejects.toMatchObject({ code: 'disposed' });
    await expect(exportPdfFrom(session, BEST)).rejects.toBeInstanceOf(ExportResourceError);

    const controller = new AbortController();
    const live = await open('text-runs', controller.signal);
    try {
      const reason = new Error('closed by host');
      controller.abort(reason);
      const run = exportPdfFrom(live, BEST);
      await expect(run).rejects.toMatchObject({ code: 'aborted', cause: reason });
    } finally {
      live.dispose();
    }
  });

  test('the reference PDF leaves hidden slides out, and so does the export', async () => {
    // `reference/hidden-slides.pdf`: the reference PDF of `reference/hidden-slides.pptx`,
    // a deck that `scripts/build-reference-decks.ts` builds (slides 2 and 4 hidden).
    const reference = join(import.meta.dir, 'reference');
    const bytes = new Uint8Array(readFileSync(join(reference, 'hidden-slides.pptx')));
    const session = await open(bytes);
    expect(sessionState(session)!.document.info.slides.map((s) => s.hidden)).toEqual([
      false,
      true,
      false,
      true,
    ]);
    session.dispose();
    const pdf = await getDocument({
      data: new Uint8Array(readFileSync(join(reference, 'hidden-slides.pdf'))),
      useSystemFonts: false,
    }).promise;
    const pages = [];
    for (let i = 1; i <= pdf.numPages; i += 1) {
      const content = await (await pdf.getPage(i)).getTextContent();
      pages.push(content.items.map((item) => ('str' in item ? item.str : '')).join(''));
    }
    expect(pages).toEqual(['Slide 1', 'Slide 3']);
    const result = await exportPdf(bytes, BEST);
    expect(result.slideNumbers).toEqual([1, 3]);
    expect(result.pageCount).toBe(pdf.numPages);
  });

  test('a dispose or a session abort during an export stops that export', async () => {
    const session = await open('slide-structure');
    // The export runs until its first wait; the dispose lands before the next slide.
    const running = exportPdfFrom(session, BEST);
    session.dispose();
    await expect(running).rejects.toMatchObject({ code: 'disposed' });

    const controller = new AbortController();
    const live = await open('slide-structure', controller.signal);
    try {
      const inFlight = exportPdfFrom(live, BEST);
      controller.abort(new Error('closed by host'));
      await expect(inFlight).rejects.toBeInstanceOf(ExportResourceError);
      await expect(inFlight).rejects.toMatchObject({ code: 'aborted' });
      // The session's signal ends the session: the abort released the document.
      expect(sessionState(live)!.disposed).toBe(true);
    } finally {
      live.dispose();
    }
  });

  test('each export has its own deadline and size limit, and exports can run together', async () => {
    const session = await open('slide-structure');
    try {
      await expect(exportPdfFrom(session, { ...BEST, timeoutMs: 1 })).rejects.toMatchObject({
        code: 'timeout',
      });
      await expect(exportPdfFrom(session, { ...BEST, maxOutputBytes: 1 })).rejects.toBeInstanceOf(
        PdfOutputLimitError
      );
      const [a, b] = await Promise.all([
        exportPdfFrom(session, BEST),
        exportPdfFrom(session, BEST),
      ]);
      expect(Buffer.from(a.bytes).equals(Buffer.from(b.bytes))).toBe(true);
      // The session still serves an export after the failures.
      expect((await exportPdfFrom(session, BEST)).pageCount).toBe(a.pageCount);
    } finally {
      session.dispose();
    }
  });

  test('session fonts reach the session: caller faces first, substitutes off on request', async () => {
    const opened = await openDocumentForExport(deck('text-runs'));
    if (!opened.ok) throw new Error('fixture');
    const carlito = sessionState(opened.session)!.document.faces.find(
      (f) => f.family === 'Carlito' && !f.bold && !f.italic
    )!;
    opened.session.dispose();
    const font = {
      family: 'Calibri',
      bold: false,
      italic: false,
      bytes: carlito.bytes,
      origin: 'host-font',
    };
    const withFont = await openDocumentForExport(deck('text-runs'), {
      fonts: fontConfiguration([font]),
    });
    if (!withFont.ok) throw new Error('fixture');
    expect(withFont.session.fontResolution.faces[0]).toEqual({
      family: 'Calibri',
      origin: 'host',
    });
    withFont.session.dispose();
    const bare = await openDocumentForExport(deck('text-runs'), { substitutes: false });
    if (!bare.ok) throw new Error('fixture');
    expect(bare.session.fontResolution.faces.some((f) => f.origin === 'substitute')).toBe(false);
    bare.session.dispose();
  });

  test('a font source with a wrong field is refused before the deck is read', async () => {
    const bytes = deck('text-runs');
    for (const bad of [
      { family: 5, bytes: new Uint8Array(4) },
      { family: 'F', bold: 'yes', bytes: new Uint8Array(4) },
    ]) {
      const options = { fonts: bad } as unknown as PdfExportOptions;
      await expect(exportPdf(bytes, options)).rejects.toThrow('fonts must be a font configuration');
      await expect(openDocumentForExport(bytes, options)).rejects.toThrow(
        'fonts must be a font configuration'
      );
    }
  });

  test('a slide that cannot be laid out has no page, and slideNumbers skips it', async () => {
    // Core lays out every slide it lists, so the session's layout stands in for a failure.
    const session = await open(withHidden());
    const document = sessionState(session)!.document;
    const layout = document.layout.bind(document);
    document.layout = (index: number) => (index === 2 ? undefined : layout(index));
    const result = await exportPdfFrom(session, { ...BEST, includeHiddenSlides: true });
    session.dispose();
    expect(result.slideNumbers).toEqual([1, 2, 4]);
    const notLaidOut = result.diagnostics.find((d) => d.code === 'slide.not-laid-out');
    expect(notLaidOut).toMatchObject({ slideNumber: 3, severity: 'unsupported' });
    expect(notLaidOut?.pageNumber).toBeUndefined();
  });

  test('options go to the call that owns them', async () => {
    const bytes = deck('text-runs');
    for (const key of [
      'fidelityPolicy',
      'includeHiddenSlides',
      'timeoutMs',
      'maxOutputBytes',
      'maxPages',
    ]) {
      const value = key === 'fidelityPolicy' ? 'strict' : key === 'includeHiddenSlides' ? true : 5;
      const options = { [key]: value } as unknown as PdfExportOptions;
      await expect(openDocumentForExport(bytes, options)).rejects.toThrow(
        `${key} must be configured in exportPdfFrom`
      );
    }
    const session = await open('text-runs');
    try {
      for (const [key, value] of [
        ['fonts', []],
        ['substitutes', true],
      ] as const) {
        const options = { [key]: value } as unknown as PdfProjectionOptions;
        await expect(exportPdfFrom(session, options)).rejects.toThrow(
          `${key} must be configured when opening the PDF session`
        );
      }
      await expect(exportPdfFrom(session, { maxPages: 0 })).rejects.toBeInstanceOf(RangeError);
      const unknown = { displayMode: 'final' } as unknown as PdfProjectionOptions;
      await expect(exportPdfFrom(session, unknown)).rejects.toThrow(
        'displayMode is not an option of exportPdfFrom'
      );
    } finally {
      session.dispose();
    }
    const notASession = { fontResolution: session.fontResolution, dispose() {} };
    await expect(exportPdfFrom(notASession, BEST)).rejects.toBeInstanceOf(TypeError);
    await expect(openDocumentForExport('deck.pptx' as never)).rejects.toBeInstanceOf(TypeError);
  });

  test('a refused deck returns the typed reason, as in DOCX Editor', async () => {
    expect(await openDocumentForExport(new Uint8Array())).toEqual({ ok: false, reason: 'empty' });
    const garbage = await openDocumentForExport(new Uint8Array([1, 2, 3, 4]));
    expect(garbage).toMatchObject({ ok: false, reason: 'zip' });
    const controller = new AbortController();
    controller.abort();
    const aborted = await openDocumentForExport(deck('text-runs'), { signal: controller.signal });
    expect(aborted).toEqual({ ok: false, reason: 'aborted' });
    // An abort during the open wins, also over a deck that fails to open.
    const during = new AbortController();
    const pending = openDocumentForExport(new Uint8Array([1, 2, 3, 4]), { signal: during.signal });
    during.abort();
    expect(await pending).toEqual({ ok: false, reason: 'aborted' });
  });
});

describe('substitute font failures', () => {
  test('give one document-wide diagnostic per face, which strict refuses', async () => {
    const opened = await openDocumentForExport(missingFontsDeck(), {});
    if (!opened.ok) throw new Error('deck');
    // Stand in for a fonts package with a missing file: the session resolves its fonts once.
    const state = sessionState(opened.session)!;
    const failure = { family: 'Calibri', file: 'Carlito-Bold.ttf', diagnostic: 'HTTP 404' };
    (state as { fontResolution: unknown }).fontResolution = {
      ...state.fontResolution,
      substituteFailures: [failure],
    };
    try {
      const result = await exportPdfFrom(opened.session, BEST);
      expect(result.diagnostics.filter((d) => d.code === 'font.substitute-failed')).toEqual([
        {
          code: 'font.substitute-failed',
          severity: 'approximation',
          message: 'The Calibri substitute (Carlito-Bold.ttf) did not load: HTTP 404.',
        },
      ]);
      expect(result.fontResolution.substituteFailures).toEqual([failure]);
      const error = await exportPdfFrom(opened.session, {}).catch((e: unknown) => e);
      expect(error).toBeInstanceOf(PdfFidelityError);
      expect((error as Error).message).toContain('document: font.substitute-failed');
    } finally {
      opened.session.dispose();
    }
  });
});
