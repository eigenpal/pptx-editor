/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pptx-to-pdf/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Adapted from DOCX Editor's `docx-to-pdf` (`packages/docx-to-pdf/src/pdf-session.ts`, commit
// 398963c7f67207cebf3c38dcc41d5f0bd1339b5f).

import { openExportSession } from './open-session.ts';
import { validateOptions } from './options.ts';
import type { OpenPdfDocumentForExportOptions, OpenPdfDocumentForExportResult } from './types.ts';

/** The options that shape one PDF. `exportPdfFrom` takes them, not the session. */
export const PROJECTION_KEYS: readonly string[] = [
  'fidelityPolicy',
  'includeHiddenSlides',
  'timeoutMs',
  'maxOutputBytes',
  'maxPages',
];

/** The options that shape the session. `openDocumentForExport` takes them, not `exportPdfFrom`. */
export const SESSION_KEYS: readonly string[] = ['fonts', 'substitutes'];

/**
 * Open a reusable PDF session over immutable PPTX bytes. Pass the session to `exportPdfFrom`
 * as many times as you need: each slide is laid out once. The caller must dispose the session.
 * Its signal controls the entire session lifetime.
 *
 * @throws `TypeError` or `RangeError` for invalid options, before the deck is read.
 * @public
 */
export async function openDocumentForExport(
  source: Uint8Array,
  options: OpenPdfDocumentForExportOptions = {}
): Promise<OpenPdfDocumentForExportResult> {
  if (!(source instanceof Uint8Array)) throw new TypeError('source must be a Uint8Array');
  validateOptions(options);
  for (const key of PROJECTION_KEYS) {
    if ((options as Record<string, unknown>)[key] !== undefined)
      throw new TypeError(`${key} must be configured in exportPdfFrom`);
  }
  return openExportSession(source, options, options.signal);
}
