/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pptx-to-pdf/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Option checks, the work budget, and the cancellation scope, after DOCX Editor's `docx-to-pdf`
// (`src/options.ts`, `src/context.ts` budgets, `src/pdf-work.ts` `withPdfWork`).

import {
  ExportResourceError,
  PdfDocumentOpenError,
  PdfEncodingError,
  PdfFidelityError,
  PdfPageLimitError,
  PdfWorkLimitError,
  type PdfExportOptions,
} from './types.ts';

/** The largest PDF this package writes. */
export const MAX_OUTPUT_BYTES = 64 * 1024 * 1024;
/** Paint operations for the whole document, before the per-page budget (DOCX Editor's value). */
export const MAX_OPERATIONS = 2_000_000;
/** Paint work grows with the page count, so each page adds this many operations. */
const OPERATIONS_PER_PAGE = 6_000;
/** The ceiling on the scaled budget, so the largest allowed export stays bounded. */
const MAX_SCALED_OPERATIONS = 40_000_000;
/** PDF content (operators, functions, pictures) for each page, before compression (DOCX Editor's value). */
const CONTENT_BYTES_PER_PAGE = 192 * 1024;
/** The ceiling on the scaled content budget. */
const MAX_SCALED_CONTENT_BYTES = 8 * MAX_OUTPUT_BYTES;
const MAX_PAGES = 10_000;
const MAX_TIMEOUT_MS = 2_147_483_647;

function positiveLimit(value: number, max: number, name: string): number {
  if (!Number.isSafeInteger(value) || value < 1 || value > max)
    throw new RangeError(`${name} must be an integer from 1 to ${max}`);
  return value;
}

/** Validate caller options before any document byte is read. */
export function validateOptions(options: PdfExportOptions): {
  timeoutMs: number;
  maxBytes: number;
  maxPages: number;
} {
  if (!options || typeof options !== 'object' || Array.isArray(options))
    throw new TypeError('options must be an object');
  const policy = options.fidelityPolicy;
  if (policy !== undefined && policy !== 'strict' && policy !== 'best-effort')
    throw new RangeError('Invalid fidelityPolicy: expected strict, best-effort');
  for (const name of ['substitutes', 'includeHiddenSlides'] as const) {
    if (options[name] !== undefined && typeof options[name] !== 'boolean')
      throw new TypeError(`${name} must be boolean`);
  }
  const fonts = options.fonts;
  if (fonts !== undefined) {
    const list = Array.isArray(fonts) ? fonts : [fonts];
    for (const f of list)
      if (
        !f ||
        typeof f !== 'object' ||
        !(f.bytes instanceof Uint8Array) ||
        typeof f.family !== 'string' ||
        !f.family ||
        // Core matches faces with `===` on these, so any other value would never match a face.
        (f.bold !== undefined && typeof f.bold !== 'boolean') ||
        (f.italic !== undefined && typeof f.italic !== 'boolean')
      )
        throw new TypeError('fonts must be a font source or an array of font sources');
  }
  const signal = options.signal;
  if (
    signal !== undefined &&
    (!signal ||
      typeof signal.aborted !== 'boolean' ||
      typeof signal.addEventListener !== 'function' ||
      typeof signal.removeEventListener !== 'function')
  )
    throw new TypeError('signal must be an AbortSignal');
  const { timeoutMs = 60_000, maxOutputBytes = MAX_OUTPUT_BYTES, maxPages = MAX_PAGES } = options;
  return {
    timeoutMs: positiveLimit(timeoutMs, MAX_TIMEOUT_MS, 'timeoutMs'),
    maxBytes: positiveLimit(maxOutputBytes, MAX_OUTPUT_BYTES, 'maxOutputBytes'),
    maxPages: positiveLimit(maxPages, MAX_PAGES, 'maxPages'),
  };
}

/** The deadline, the cancellation signal, and the paint budget of one conversion. */
export class Work {
  #operations = 0;
  #operationLimit = MAX_OPERATIONS;
  #contentBytes = 0;
  #contentLimit = MAX_OUTPUT_BYTES;

  constructor(
    readonly signal: AbortSignal,
    private readonly deadline: number
  ) {}

  /** Throw the abort reason (a `PdfWorkLimitError`) when the conversion must stop. */
  check(): void {
    if (this.signal.aborted) throw this.signal.reason;
    if (Date.now() >= this.deadline) throw new PdfWorkLimitError('timeout');
  }

  /** Let the timer and the signal run between long steps, then check the limits again. */
  async yield(): Promise<void> {
    await new Promise((resolve) => setImmediate(resolve));
    this.check();
  }

  /** Scale the paint budgets to the page count, as DOCX Editor does. */
  allowPages(pages: number): void {
    this.#operationLimit = Math.min(
      MAX_SCALED_OPERATIONS,
      MAX_OPERATIONS + pages * OPERATIONS_PER_PAGE
    );
    this.#contentLimit = Math.max(
      MAX_OUTPUT_BYTES,
      Math.min(MAX_SCALED_CONTENT_BYTES, pages * CONTENT_BYTES_PER_PAGE)
    );
  }

  /**
   * Count paint operations and check the limits. A file cannot make the writer run without bound,
   * and a long paint stops at the deadline (DOCX Editor's `Work.tick`).
   */
  count(operations: number): void {
    this.check();
    this.#operations += operations;
    if (this.#operations > this.#operationLimit) throw new PdfWorkLimitError('operations');
  }

  /**
   * Reserve PDF content bytes before they are kept in memory: operators, shading functions, and
   * pictures (DOCX Editor's `Work.reserveContent`). `maxOutputBytes` is checked only after the
   * save, so this bounds the PDF content that the paint keeps. A decoded PNG is bounded only by
   * the pixel limit of each picture.
   */
  reserveContent(bytes: number): void {
    this.#contentBytes += bytes;
    if (this.#contentBytes > this.#contentLimit) throw new PdfWorkLimitError('operations');
  }
}

/**
 * Run one conversion inside a deadline and the caller's signal. An already aborted signal stops
 * before the deck opens. Errors that are not this package's typed errors become
 * `PdfEncodingError`, after the limits are checked, so a timeout is reported as a timeout.
 */
export async function withPdfWork<T>(
  options: PdfExportOptions,
  run: (work: Work, phase: () => number) => Promise<T>
): Promise<T> {
  const { timeoutMs } = validateOptions(options);
  const controller = new AbortController();
  const signal = options.signal;
  const abort = (): void =>
    controller.abort(new PdfWorkLimitError('aborted', { cause: signal?.reason }));
  if (signal?.aborted) abort();
  signal?.addEventListener('abort', abort, { once: true });
  const timer = setTimeout(() => controller.abort(new PdfWorkLimitError('timeout')), timeoutMs);
  const work = new Work(controller.signal, Date.now() + timeoutMs);
  const started = performance.now();
  try {
    work.check();
    return await run(work, () => Math.round(performance.now() - started));
  } catch (error) {
    if (
      error instanceof PdfFidelityError ||
      error instanceof PdfDocumentOpenError ||
      error instanceof ExportResourceError ||
      error instanceof PdfEncodingError ||
      error instanceof PdfWorkLimitError ||
      // Options are validated before the work starts, so a TypeError or RangeError from inside
      // the work is an engine failure and becomes PdfEncodingError (as in DOCX Editor).
      error instanceof PdfPageLimitError
    )
      throw error;
    work.check();
    throw new PdfEncodingError('PDF encoding failed', { cause: error });
  } finally {
    clearTimeout(timer);
    signal?.removeEventListener('abort', abort);
  }
}
