/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pptx-to-pdf/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
/**
 * `@pptx-editor.dev/pptx-to-pdf`: PPTX to PDF on Node.js through the PPTX Editor engine.
 *
 * The names follow DOCX Editor's `@docx-editor.dev/docx-to-pdf` where the concept is the same.
 *
 * @example
 * ```ts
 * import { exportPdf } from '@pptx-editor.dev/pptx-to-pdf';
 *
 * const { bytes, diagnostics } = await exportPdf(pptx, { fidelityPolicy: 'best-effort' });
 * ```
 *
 * To export one deck several times from one layout, open it once with `openDocumentForExport`
 * and call `exportPdfFrom` for each export. Dispose the session after the last export.
 *
 * @packageDocumentation
 */

export { exportPdf, exportPdfFrom } from './export-pdf.ts';
export { openDocumentForExport } from './pdf-session.ts';
export {
  ExportResourceError,
  PdfDocumentOpenError,
  PdfEncodingError,
  PdfFidelityError,
  PdfOutputLimitError,
  PdfPageLimitError,
  PdfWorkLimitError,
} from './types.ts';
export type {
  OpenPdfDocumentForExportOptions,
  OpenPdfDocumentForExportResult,
  PdfDiagnostic,
  PdfExportOptions,
  PdfExportResult,
  PdfExportSession,
  PdfExportTimings,
  PdfFontResolution,
  PdfFontsSource,
  PdfProjectionOptions,
} from './types.ts';
export type { FontSource } from '@pptx-editor.dev/core/layout';
export type {
  FontSubstitution,
  OpenError,
  SubstituteFontFailure,
} from '@pptx-editor.dev/core/editor';
