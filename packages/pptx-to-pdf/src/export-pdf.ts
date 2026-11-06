/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pptx-to-pdf/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// `exportPdf` and `exportPdfFrom`: PPTX → PDF bytes on Node.js, after DOCX Editor's
// `docx-to-pdf` (`packages/docx-to-pdf/src/export-pdf.ts`). Both run the editor's own pipeline,
// `openPresentation` → `session.layout(i)`, and paint each layout into a PDF page.

import { PDFDocument } from 'pdf-lib';
import type { RegisteredFace, SlideLayout } from '@pptx-editor.dev/core/layout';
import { deckMetadata } from './metadata.ts';
import {
  checkSession,
  openExportSession,
  sessionState,
  type SessionState,
} from './open-session.ts';
import { validateOptions, withPdfWork, type Work } from './options.ts';
import { clip, SlideDiagnostics, type PaintResources } from './page-ops.ts';
import { paintSlide } from './paint.ts';
import { PROJECTION_KEYS, SESSION_KEYS } from './pdf-session.ts';
import {
  PdfDocumentOpenError,
  PdfFidelityError,
  PdfOutputLimitError,
  PdfPageLimitError,
  PdfWorkLimitError,
  type PdfDiagnostic,
  type PdfExportOptions,
  type PdfExportResult,
  type PdfExportSession,
  type PdfProjectionOptions,
} from './types.ts';

/** A fixed date, so the same deck gives the same bytes (as in DOCX Editor). */
const FIXED_DATE = new Date('2020-01-01T00:00:00Z');

/**
 * Convert PPTX bytes to a PDF, one page per slide, on Node.js. Disposes its session after
 * success or failure.
 *
 * The default `fidelityPolicy` is `strict`: if a slide has content the PDF cannot reproduce,
 * the conversion throws `PdfFidelityError` with every diagnostic. Pass
 * `fidelityPolicy: 'best-effort'` to get the PDF and read `diagnostics` instead.
 *
 * @throws `TypeError` or `RangeError` for invalid options, before the deck is read.
 * @throws `PdfDocumentOpenError` when core refuses the file.
 * @throws `PdfPageLimitError` when the deck has more slides than `maxPages`.
 * @throws `PdfFidelityError` under the strict policy.
 * @throws `PdfOutputLimitError` when the PDF is larger than `maxOutputBytes`.
 * @throws `PdfWorkLimitError` on `timeoutMs`, on `signal` abort (its `reason` is the `cause`),
 *   or when the paint budget (operations or content bytes) runs out.
 * @throws `PdfEncodingError` for any other failure while writing the PDF.
 * @public
 */
export async function exportPdf(
  source: Uint8Array,
  options: PdfExportOptions = {}
): Promise<PdfExportResult> {
  if (!(source instanceof Uint8Array)) throw new TypeError('source must be a Uint8Array');
  validateOptions(options);
  return withPdfWork(options, async (work, phase) => {
    const { fonts, substitutes } = options;
    const opened = await openExportSession(
      source,
      {
        ...(fonts !== undefined ? { fonts } : {}),
        ...(substitutes !== undefined ? { substitutes } : {}),
      },
      work.signal
    );
    if (!opened.ok) {
      // An abort during open is the work's abort, not a refused file.
      if (opened.reason === 'aborted') {
        work.check();
        throw new PdfWorkLimitError('aborted');
      }
      throw new PdfDocumentOpenError(opened.reason, opened.detail);
    }
    const state = sessionState(opened.session) as SessionState;
    try {
      work.check();
      return await renderSession(state, options, work, phase, phase());
    } finally {
      opened.session.dispose();
    }
  });
}

/**
 * Encode a caller-owned session from `openDocumentForExport` without reopening the deck. Each
 * slide is laid out once per session, so a second export reuses the layout. Does not dispose
 * the session: the caller disposes it after all exports. Cancellation stops this export only;
 * the session's own signal controls the session.
 *
 * @throws `TypeError` for a value that is not an export session, or for session options.
 * @throws `ExportResourceError` when the session is disposed or its signal aborted.
 * @throws The errors of `exportPdf`, other than `PdfDocumentOpenError`.
 * @public
 */
export async function exportPdfFrom(
  session: PdfExportSession,
  options: PdfProjectionOptions = {}
): Promise<PdfExportResult> {
  validateOptions(options);
  for (const key of Object.keys(options)) {
    if ([...PROJECTION_KEYS, 'signal'].includes(key)) continue;
    throw new TypeError(
      SESSION_KEYS.includes(key)
        ? `${key} must be configured when opening the PDF session`
        : `${key} is not an option of exportPdfFrom`
    );
  }
  const state = sessionState(session);
  if (!state) throw new TypeError('exportPdfFrom requires a session from openDocumentForExport');
  checkSession(state);
  return withPdfWork(options, (work, phase) => renderSession(state, options, work, phase, 0));
}

/** Lay out the selected slides, paint them, and save the PDF. */
async function renderSession(
  state: SessionState,
  options: PdfProjectionOptions,
  work: Work,
  phase: () => number,
  openMs: number
): Promise<PdfExportResult> {
  const { maxBytes, maxPages } = validateOptions(options);
  const session = state.document;
  const step = (): void => {
    work.check();
    checkSession(state);
  };
  step();
  const slides = session.info.slides
    .map((slide, index) => ({ slide, index }))
    .filter(({ slide }) => options.includeHiddenSlides === true || !slide.hidden);
  if (slides.length > maxPages) throw new PdfPageLimitError(maxPages, slides.length);
  work.allowPages(slides.length);

  const diagnostics: PdfDiagnostic[] = [];
  const layouts: { index: number; layout: SlideLayout }[] = [];
  for (const { index } of slides) {
    step();
    // The session caches each slide's layout, so a second export reuses it.
    const layout = session.layout(index);
    if (layout) layouts.push({ index, layout });
    else {
      const diag = new SlideDiagnostics(index, undefined);
      diag.add('slide.not-laid-out', 'unsupported', 'The slide could not be laid out.');
      for (const d of diag.list) diagnostics.push(d);
    }
  }
  // A packaged substitute that failed to load changes how text measures on every page: one
  // document-wide diagnostic per face, with no slide or page, as DOCX Editor reports
  // font source failures.
  for (const f of state.fontResolution.substituteFailures)
    diagnostics.push({
      code: 'font.substitute-failed',
      severity: 'approximation',
      message: `The ${clip(f.family)} substitute${f.file ? ` (${clip(f.file)})` : ''} did not load: ${clip(f.diagnostic, 200)}.`,
    });
  const layoutMs = phase() - openMs;

  const doc = await PDFDocument.create({ updateMetadata: false });
  doc.setProducer('pptx-editor.dev');
  doc.setCreator('pptx-editor.dev');
  doc.setCreationDate(FIXED_DATE);
  doc.setModificationDate(FIXED_DATE);
  const metadata = deckMetadata(session.store.package);
  if (metadata.title) doc.setTitle(metadata.title);
  if (metadata.author) doc.setAuthor(metadata.author);
  if (metadata.subject) doc.setSubject(metadata.subject);
  if (metadata.keywords) doc.setKeywords([metadata.keywords]);

  const faces = new Map<string, RegisteredFace>(session.faces.map((f) => [f.id, f]));
  const res: PaintResources = {
    doc,
    pkg: session.store.package,
    work,
    faces,
    fonts: new Map(),
    images: new Map(),
    gradients: new Map(),
  };
  for (const [pageIndex, { index, layout }] of layouts.entries()) {
    step();
    const diag = new SlideDiagnostics(index, pageIndex);
    // Let the timer and the caller's signal run between slides, as DOCX Editor does.
    await work.yield();
    await paintSlide(layout, res, diag);
    for (const d of diag.list) diagnostics.push(d);
  }
  // Write each face's subset and its ToUnicode map, after every page has used it.
  for (const font of res.fonts.values()) if (font) await font.finish(work);
  const paintMs = phase() - openMs - layoutMs;

  const frozen = Object.freeze(diagnostics.map((d) => Object.freeze(d)));
  if (
    (options.fidelityPolicy ?? 'strict') === 'strict' &&
    frozen.some((d) => d.severity !== 'information')
  )
    throw new PdfFidelityError(frozen);
  step();
  // pdf-lib yields between batches; the deadline and the signal are checked again after.
  const bytes = await doc.save({ useObjectStreams: false, objectsPerTick: 1000 });
  step();
  if (bytes.byteLength > maxBytes) throw new PdfOutputLimitError(maxBytes, bytes.byteLength);
  return Object.freeze({
    bytes,
    pageCount: layouts.length,
    slideNumbers: Object.freeze(layouts.map(({ index }) => index + 1)),
    fontResolution: state.fontResolution,
    diagnostics: frozen,
    timings: Object.freeze({
      openMs,
      layoutMs,
      paintMs,
      saveMs: phase() - openMs - layoutMs - paintMs,
    }),
  });
}
