/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pptx-to-pdf/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// One way to open a deck for PDF output, after DOCX Editor's `docx-to-pdf`
// (`packages/docx-to-pdf/src/open-session.ts`, commit 398963c7). `exportPdf`, `openDocumentForExport`, and the
// evaluation scripts share it, so a tool sees exactly the layout the PDF is painted from.

import { openPresentation, type DocumentSession } from '@pptx-editor.dev/core/editor';
import type { FontSource } from '@pptx-editor.dev/core/layout';
import {
  ExportResourceError,
  type OpenPdfDocumentForExportOptions,
  type OpenPdfDocumentForExportResult,
  type PdfExportSession,
  type PdfFontResolution,
} from './types.ts';

/** The options that shape the session, without the ones that shape one PDF. */
export type SessionOptions = Omit<OpenPdfDocumentForExportOptions, 'signal'>;

/** What an export session holds. Callers see only the `PdfExportSession` face. */
export interface SessionState {
  readonly document: DocumentSession;
  readonly fontResolution: PdfFontResolution;
  /** The signal given at open. It controls the whole session lifetime. */
  readonly signal: AbortSignal | undefined;
  disposed: boolean;
}

const states = new WeakMap<PdfExportSession, SessionState>();

/** The state of a session that this package opened, or undefined for any other value. */
export function sessionState(session: unknown): SessionState | undefined {
  return session !== null && typeof session === 'object'
    ? states.get(session as PdfExportSession)
    : undefined;
}

/** Throw when the session can no longer serve an export. An abort wins over the dispose it causes. */
export function checkSession(state: SessionState): void {
  if (state.signal?.aborted)
    throw new ExportResourceError('aborted', { cause: state.signal.reason });
  if (state.disposed) throw new ExportResourceError('disposed');
}

/** The faces a session measures with, frozen once at open: the deck does not change. */
function fontResolution(document: DocumentSession): PdfFontResolution {
  const report = document.fontReport;
  return Object.freeze({
    faces: Object.freeze(report.faces.map((f) => Object.freeze({ ...f }))),
    substitutions: Object.freeze(document.substitutedFamilies().map((s) => Object.freeze(s))),
    refusedEmbedded: Object.freeze([...report.refusedEmbedded]),
  });
}

/**
 * Open PPTX bytes the way `exportPdf` does: caller fonts first, then the deck's embedded fonts,
 * then the packaged substitutes. Never throws for a malformed file; it returns the reason.
 */
export async function openExportSession(
  source: Uint8Array,
  options: SessionOptions,
  signal?: AbortSignal
): Promise<OpenPdfDocumentForExportResult> {
  if (signal?.aborted) return { ok: false, reason: 'aborted' };
  const fonts: readonly FontSource[] | undefined =
    options.fonts === undefined
      ? undefined
      : Array.isArray(options.fonts)
        ? options.fonts
        : [options.fonts as FontSource];
  const opened = await openPresentation(source, {
    ...(fonts ? { fonts } : {}),
    ...(options.substitutes === false ? { substitutes: false } : {}),
  });
  // An abort during the open wins over the open's own result, as the signal's doc says.
  if (signal?.aborted) {
    if (opened.ok) opened.session.dispose();
    return { ok: false, reason: 'aborted' };
  }
  if (!opened.ok) {
    const detail = 'detail' in opened.error ? String(opened.error.detail) : undefined;
    return { ok: false, reason: opened.error.reason, ...(detail ? { detail } : {}) };
  }
  const state: SessionState = {
    document: opened.session,
    fontResolution: fontResolution(opened.session),
    signal,
    disposed: false,
  };
  // The session's signal ends the session, as in DOCX Editor: an abort releases the document.
  const onAbort = (): void => session.dispose();
  const session: PdfExportSession = Object.freeze({
    fontResolution: state.fontResolution,
    dispose(): void {
      if (state.disposed) return;
      state.disposed = true;
      signal?.removeEventListener('abort', onAbort);
      state.document.dispose();
    },
  });
  signal?.addEventListener('abort', onAbort, { once: true });
  states.set(session, state);
  return { ok: true, session };
}
