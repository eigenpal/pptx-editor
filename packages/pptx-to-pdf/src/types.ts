/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pptx-to-pdf/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// The public types, adapted from DOCX Editor's `docx-to-pdf` (`packages/docx-to-pdf/src/types.ts`
// at adf0e178) where the concept is the same.

import type {
  FontSubstitution,
  OpenError,
  SubstituteFontFailure,
} from '@pptx-editor.dev/core/editor';
import type { FontSource } from '@pptx-editor.dev/core/layout';

/** A bounded explanation of content the PDF cannot reproduce. @public */
export interface PdfDiagnostic {
  /** Machine-readable diagnostic identifier, from a fixed set. Handle unknown codes by severity. */
  readonly code: string;
  /** Human-readable explanation. Do not use this text for program control. */
  readonly message: string;
  /** Zero-based PDF page index. Omitted when the slide produced no page. */
  readonly pageIndex?: number;
  /** One-based PDF page number for display. */
  readonly pageNumber?: number;
  /** Zero-based slide index in the deck, hidden slides included. Omitted for document-wide diagnostics. */
  readonly slideIndex?: number;
  /** One-based slide number in the deck, hidden slides included. */
  readonly slideNumber?: number;
  /** Strict export permits only information diagnostics. */
  readonly severity: 'unsupported' | 'approximation' | 'information';
}

/** One font face or several. Earlier faces win. @public */
export type PdfFontsSource = FontSource | readonly FontSource[];

/** Node conversion controls. @public */
export interface PdfExportOptions {
  /** Font faces tried before embedded and substitute fonts. */
  readonly fonts?: PdfFontsSource;
  /**
   * Load the metric-compatible substitutes from `@docx-editor.dev/fonts`. Defaults to true.
   * DOCX Editor configures this through its font sources instead.
   */
  readonly substitutes?: boolean;
  /** Resolve fonts from the pinned Google Fonts catalog, including Arabic and CJK fallbacks.
   * Makes network requests. Defaults to false. Supplied fonts keep precedence.
   */
  readonly googleFonts?: boolean;
  /** Include slides marked hidden (`p:sld/@show="0"`). Defaults to false. */
  readonly includeHiddenSlides?: boolean;
  /** Reject unsupported or approximate content. Defaults to `strict`. */
  readonly fidelityPolicy?: 'strict' | 'best-effort';
  /** Cooperative conversion deadline. Defaults to 60000 milliseconds. */
  readonly timeoutMs?: number;
  /** Maximum encoded bytes, from 1 to 67108864. Defaults to 67108864. */
  readonly maxOutputBytes?: number;
  /** Maximum output pages, from 1 to 10000. Defaults to 10000. */
  readonly maxPages?: number;
  /** Cancels the conversion. Its `reason` is the `cause` of the thrown `PdfWorkLimitError`. */
  readonly signal?: AbortSignal;
}

/** Elapsed milliseconds for each conversion phase. @public */
export interface PdfExportTimings {
  /** Reading and opening the deck. */
  readonly openMs: number;
  /** Laying out every exported slide. */
  readonly layoutMs: number;
  /** Painting every page into the PDF. */
  readonly paintMs: number;
  /** Encoding the PDF bytes. */
  readonly saveMs: number;
}

/** The faces the conversion drew with, and the families it substituted. @public */
export interface PdfFontResolution {
  /** Each face drawn with: its family, and its origin (for example `embedded`, `provided`, or `substitute`). */
  readonly faces: readonly { readonly family: string; readonly origin: string }[];
  /** Families the deck asked for that no face covered, with the family used instead. */
  readonly substitutions: readonly FontSubstitution[];
  /** Embedded faces the deck carries that may not be used, with the reason. */
  readonly refusedEmbedded: readonly string[];
  /** Packaged substitute faces that failed to load, each with its family, file, and reason. */
  readonly substituteFailures: readonly SubstituteFontFailure[];
}

/** Result owns its bytes. @public */
export interface PdfExportResult {
  /** Owned, mutable PDF bytes. Use application/pdf when serving them. */
  readonly bytes: Uint8Array;
  /** Physical PDF page count: one page per exported slide. */
  readonly pageCount: number;
  /**
   * The one-based deck slide number of each page, in page order. Hidden slides (unless
   * `includeHiddenSlides` is true) and slides that could not be laid out have no page, so page
   * `i` shows slide `slideNumbers[i]`.
   */
  readonly slideNumbers: readonly number[];
  /** Selected faces and substitutions. */
  readonly fontResolution: PdfFontResolution;
  /** Immutable output limitations and informational notices. */
  readonly diagnostics: readonly PdfDiagnostic[];
  /** Wall-clock milliseconds spent opening, laying out, painting, and saving. */
  readonly timings: PdfExportTimings;
}

/** Where a diagnostic is, for messages: `slide 3`, or the document. */
function place(d: PdfDiagnostic): string {
  return d.slideNumber !== undefined ? `slide ${d.slideNumber}` : 'document';
}

/** Requested content could not be represented faithfully. @public */
export class PdfFidelityError extends Error {
  /** Stable error code. */
  readonly code = 'fidelityUnsupported';
  /** Every diagnostic of the conversion; the blocking ones name the slide and feature. */
  readonly diagnostics: readonly PdfDiagnostic[];
  /** @param diagnostics - Every diagnostic of the conversion. */
  constructor(diagnostics: readonly PdfDiagnostic[]) {
    const blocking = diagnostics.filter((d) => d.severity !== 'information');
    const first = blocking[0];
    const more = blocking.length > 1 ? ` (+${blocking.length - 1} more)` : '';
    super(
      first
        ? `PDF export cannot faithfully represent all requested content: ${place(first)}: ${first.code}${more}`
        : 'PDF export cannot faithfully represent all requested content'
    );
    this.diagnostics = diagnostics;
    this.name = 'PdfFidelityError';
  }
}

/** Core refused the input package. @public */
export class PdfDocumentOpenError extends Error {
  /** Stable error code. */
  readonly code = 'documentOpenFailed';
  /** Why core refused the package, as `openPresentation` reports it. */
  readonly reason: OpenError['reason'];
  /** More detail, when core gives it. */
  readonly detail?: string;
  /**
   * @param reason - Why core refused the package.
   * @param detail - More detail, when core gives it.
   */
  constructor(reason: OpenError['reason'], detail?: string) {
    super(`Cannot open PPTX: ${reason}${detail ? `: ${detail}` : ''}`);
    this.reason = reason;
    if (detail !== undefined) this.detail = detail;
    this.name = 'PdfDocumentOpenError';
  }
}

/**
 * A caller-owned export session cannot serve an export: it was disposed, or the signal given to
 * `openDocumentForExport` aborted (which also disposes it). For `aborted`, `cause` is the
 * signal's `reason`. DOCX Editor throws the same name from its Core. @public
 */
export class ExportResourceError extends Error {
  /** Stable error code: the session was disposed, or its signal aborted. */
  readonly code: 'disposed' | 'aborted';
  /**
   * @param code - Why the session cannot serve the export.
   * @param options - For `aborted`, `cause` is the signal's `reason`.
   */
  constructor(code: 'disposed' | 'aborted', options?: ErrorOptions) {
    super(
      code === 'disposed'
        ? 'The PDF export session is disposed'
        : 'The PDF export session was aborted',
      options
    );
    this.code = code;
    this.name = 'ExportResourceError';
  }
}

/** PDF encoding failed after a document was opened. @public */
export class PdfEncodingError extends Error {
  /** Stable error code: `outputTooLarge` on the {@link PdfOutputLimitError} subclass. */
  readonly code: 'encodingFailed' | 'outputTooLarge' = 'encodingFailed';
  /**
   * @param message - What failed.
   * @param options - `cause` holds the original error.
   */
  constructor(message: string, options?: ErrorOptions) {
    super(message, options);
    this.name = 'PdfEncodingError';
  }
}

/** Encoded output exceeds the configured byte limit. @public */
export class PdfOutputLimitError extends PdfEncodingError {
  /** Stable error code. */
  override readonly code = 'outputTooLarge';
  /** The `maxOutputBytes` option. */
  readonly limit: number;
  /** The encoded size in bytes. */
  readonly actual: number;
  /**
   * @param limit - The `maxOutputBytes` option.
   * @param actual - The encoded size in bytes.
   */
  constructor(limit: number, actual: number) {
    super(`PDF exceeds maxOutputBytes (${limit}); encoded ${actual} bytes`);
    this.limit = limit;
    this.actual = actual;
    this.name = 'PdfOutputLimitError';
  }
}

/** The deck exceeds the configured page limit. @public */
export class PdfPageLimitError extends RangeError {
  /** Stable error code. */
  readonly code = 'pageLimitExceeded';
  /** The `maxPages` option. */
  readonly limit: number;
  /** The number of pages the deck would export. */
  readonly actual: number;
  /**
   * @param limit - The `maxPages` option.
   * @param actual - The number of pages the deck would export.
   */
  constructor(limit: number, actual: number) {
    super(`PDF exceeds maxPages (${limit}); the deck has ${actual} pages`);
    this.limit = limit;
    this.actual = actual;
    this.name = 'PdfPageLimitError';
  }
}

/**
 * The conversion passed its deadline, was cancelled, or ran past its paint budget. For
 * `aborted`, `cause` is the signal's `reason`. @public
 */
export class PdfWorkLimitError extends Error {
  /**
   * Stable error code: past `timeoutMs`, cancelled through `signal`, or past the paint budget
   * (paint operations or PDF content bytes).
   */
  readonly code: 'timeout' | 'aborted' | 'operations';
  /**
   * @param code - Which limit stopped the work.
   * @param options - For `aborted`, `cause` is the signal's `reason`.
   */
  constructor(code: 'timeout' | 'aborted' | 'operations', options?: ErrorOptions) {
    super(
      code === 'timeout'
        ? 'PDF export exceeded timeoutMs'
        : code === 'aborted'
          ? 'PDF export was aborted'
          : 'PDF export exceeded its paint budget',
      options
    );
    this.code = code;
    this.name = 'PdfWorkLimitError';
  }
}

/** Output controls for a caller-owned PDF session. @public */
export type PdfProjectionOptions = Pick<
  PdfExportOptions,
  'fidelityPolicy' | 'includeHiddenSlides' | 'timeoutMs' | 'maxOutputBytes' | 'maxPages' | 'signal'
>;

/** Font controls for a reusable session over immutable PPTX bytes. @public */
export type OpenPdfDocumentForExportOptions = Omit<
  PdfExportOptions,
  'fidelityPolicy' | 'includeHiddenSlides' | 'timeoutMs' | 'maxOutputBytes' | 'maxPages' | 'signal'
> & {
  /**
   * Ends the session. Aborted before or during the open, the result is `{ ok: false, reason:
   * 'aborted' }` (the engine's open cannot be cancelled, so it is checked when the open ends). Aborted later, the session is disposed, and exports reject with
   * `ExportResourceError` (`aborted`), whose `cause` is the signal's `reason`.
   */
  readonly signal?: AbortSignal;
};

/**
 * An open deck with its fonts and slide layouts, for `exportPdfFrom`. Each slide is laid out
 * once and reused by every export. The caller owns the session and must call `dispose()`. @public
 */
export interface PdfExportSession {
  /** The faces the session measures with, and the families it substituted. */
  readonly fontResolution: PdfFontResolution;
  /** Release the session. Later exports reject with `ExportResourceError` (`disposed`). */
  dispose(): void;
}

/** Open result, with DOCX Editor's success and refusal shape. @public */
export type OpenPdfDocumentForExportResult =
  | { readonly ok: true; readonly session: PdfExportSession }
  | {
      readonly ok: false;
      readonly reason: OpenError['reason'] | 'aborted';
      readonly detail?: string;
    };
