// The loading screen's state, shared by the React and Vue adapters (`PptxEditor.Loading` and
// `PptxEditorLoading`), after DOCX Editor's `DocxEditor.Loading`. Both adapters derive their
// answer from here, so the two screens cannot disagree about when a deck is still opening.
//
// Every export is adapter glue, not a stable API: `@internal`.

import type {
  DocumentOpenFailure,
  DocumentOpenFailureReason,
  EditorSnapshot,
} from '../contracts/index.ts';

/** EMU in one CSS pixel (ECMA-376 Part 1, 20.1.2.1: 914400 EMU per inch, 96 px per inch). */
const EMU_PER_PX = 9525;
/** The default slide size, 16:9 (13.333 in by 7.5 in), until a deck gives its own. */
const DEFAULT_SLIDE = { width: 12_192_000, height: 6_858_000 };

/**
 * Whether the loading screen shows: no deck yet and nothing went wrong (`empty`), or a deck
 * is opening (`loading`). A parse failure (`error`) clears it, so a broken file never spins
 * forever. @internal
 */
export function showsLoading(snapshot: EditorSnapshot): boolean {
  return snapshot.status === 'empty' || snapshot.status === 'loading';
}

/** The loading sheet's size: the slide size at the current zoom, in CSS pixels. @internal */
export interface LoadingSheetSize {
  readonly width: number;
  readonly height: number;
}

/**
 * The size of the slide-shaped sheet the loading screen shows: the open deck's slide size
 * (the previous deck while the next one opens), or 16:9 before any deck. @internal
 */
export function loadingSheetSize(snapshot: EditorSnapshot): LoadingSheetSize {
  const known = snapshot.slideWidth > 0 && snapshot.slideHeight > 0;
  const width = known ? snapshot.slideWidth : DEFAULT_SLIDE.width;
  const height = known ? snapshot.slideHeight : DEFAULT_SLIDE.height;
  const zoom = snapshot.zoom > 0 ? snapshot.zoom : 1;
  return {
    width: Math.round((width / EMU_PER_PX) * zoom),
    height: Math.round((height / EMU_PER_PX) * zoom),
  };
}

/** Equal sheet sizes, so a selector re-renders only when the size changes. @internal */
export function sameLoadingSheetSize(a: LoadingSheetSize, b: LoadingSheetSize): boolean {
  return a.width === b.width && a.height === b.height;
}

/** The longest `yieldToPaint` waits for a frame. */
const PAINT_WAIT_MS = 100;

/**
 * Resolves after the browser paints the next frame, or after {@link PAINT_WAIT_MS}. `load` awaits it between publishing the
 * `loading` state and the parse, which blocks the main thread for a large deck, so the
 * loading screen paints first. The spinner turns by a CSS transform, which the compositor
 * runs while the main thread is busy. Outside a browser it resolves at once.
 */
export function yieldToPaint(): Promise<void> {
  if (typeof requestAnimationFrame !== 'function' || typeof document === 'undefined')
    return Promise.resolve();
  // `requestAnimationFrame` runs before the paint; the task after it runs after the paint. A
  // hidden page runs no frames, so a timer ends the wait there: the open never waits for the
  // tab to show.
  return new Promise((resolve) => {
    const done = () => setTimeout(resolve, 0);
    const timer = setTimeout(done, PAINT_WAIT_MS);
    requestAnimationFrame(() => (clearTimeout(timer), done()));
  });
}

/**
 * The reason group of each open code (`EditorSnapshot.error`), after DOCX Editor's open codes:
 * a limit has a code of its own (`xml-depth-limit`), as DOCX Editor's `open.xml-depth-limit`.
 */
const OPEN_FAILURE_REASONS: Readonly<Record<string, DocumentOpenFailureReason>> = Object.assign(
  Object.create(null) as Record<string, DocumentOpenFailureReason>,
  {
    zip: 'damaged',
    xml: 'damaged',
    'no-content-types': 'damaged',
    'not-a-presentation': 'not-pptx',
    empty: 'not-pptx',
    'too-many-parts': 'too-large',
    'zip-entry-limit': 'too-large',
    'zip-size-limit': 'too-large',
    'xml-size-limit': 'too-large',
    'xml-depth-limit': 'too-large',
    'xml-element-limit': 'too-large',
    'xml-dtd-forbidden': 'blocked',
    failed: 'unknown',
  }
);

/**
 * The refusal a person sees for the open code `code`, after DOCX Editor's
 * `documentOpenFailure`: frozen, and a new object for each call, so a second refusal of the
 * same file shows again. @internal
 */
export function documentOpenFailure(code: string): DocumentOpenFailure {
  return Object.freeze({ code, reason: OPEN_FAILURE_REASONS[code] ?? 'unknown' });
}

/** The open code of each zip and XML reader limit, after DOCX Editor's `openDiagnosticCode`. */
const LIMIT_CODES: Readonly<Record<string, Readonly<Record<string, string>>>> = {
  zip: { 'too-many-entries': 'zip-entry-limit', 'too-large': 'zip-size-limit' },
  xml: {
    'too-large': 'xml-size-limit',
    'too-deep': 'xml-depth-limit',
    'too-many-elements': 'xml-element-limit',
    'dtd-forbidden': 'xml-dtd-forbidden',
    'entity-forbidden': 'xml-dtd-forbidden',
  },
};

/** The open code of an engine rejection: its reason, or the limit it crossed. @internal */
export function openRejectionCode(error: {
  readonly reason: string;
  readonly detail?: unknown;
}): string {
  const limits = LIMIT_CODES[error.reason];
  const limit = limits && typeof error.detail === 'string' ? limits[error.detail] : undefined;
  return limit ?? error.reason;
}

/**
 * The open code of an open that threw, as DOCX Editor's `open.failed`. It carries no engine
 * text: an engine message can quote the file. @internal
 */
export const OPEN_FAILED = 'failed';

/** The version of `@pptx-editor.dev/core` that a build inlines; `source` when run from source. */
declare const __PPTX_EDITOR_VERSION__: string | undefined;
const EDITOR_VERSION =
  typeof __PPTX_EDITOR_VERSION__ === 'string' ? __PPTX_EDITOR_VERSION__ : 'source';

/** What `openFailureReport` adds besides the failure. @public */
export interface OpenFailureReportContext {
  /** The browser, such as `navigator.userAgent`. Defaults to it in a browser. */
  readonly userAgent?: string;
  /** When the open failed. Defaults to now. */
  readonly time?: Date;
}

/**
 * Plain text that describes a refused open, for a person to copy and send with a support
 * request: the code, the reason, the editor version, the browser, and the time.
 * It holds no file content and no file name. The text is in English, for the support team that
 * reads it, whatever the editor's locale. **Copy details** in `PptxEditor.OpenFailure` copies
 * it. @public
 */
export function openFailureReport(
  failure: DocumentOpenFailure,
  context: OpenFailureReportContext = {}
): string {
  const agent =
    context.userAgent ?? (typeof navigator === 'undefined' ? undefined : navigator.userAgent);
  return [
    'PPTX Editor could not open a file.',
    `Code: ${failure.code}`,
    `Reason: ${failure.reason}`,
    `Editor version: ${EDITOR_VERSION}`,
    ...(agent ? [`Browser: ${agent}`] : []),
    `Time: ${(context.time ?? new Date()).toISOString()}`,
  ].join('\n');
}

/** The i18n key of the message for each open failure reason. @internal */
export const OPEN_FAILURE_MESSAGE_KEYS = {
  damaged: 'openFailure.reasons.damaged',
  'not-pptx': 'openFailure.reasons.notPptx',
  'too-large': 'openFailure.reasons.tooLarge',
  blocked: 'openFailure.reasons.blocked',
  unknown: 'openFailure.reasons.unknown',
} as const satisfies Record<DocumentOpenFailureReason, string>;

/** The result of **Copy details** for one refusal: the text, and whether it is copied. @internal */
export interface OpenFailureCopy {
  readonly failure: DocumentOpenFailure;
  /** The report, built once, so the text a person copies by hand is the text tried. */
  readonly report: string;
  /** `false` when the page has no clipboard access: the notice shows `report` then. */
  readonly copied: boolean;
}

/**
 * Copy `openFailureReport(failure)` to the clipboard. Shared by the React and Vue notices.
 * @internal
 */
export async function copyOpenFailureReport(
  failure: DocumentOpenFailure
): Promise<OpenFailureCopy> {
  const report = openFailureReport(failure);
  const clipboard = typeof navigator === 'undefined' ? undefined : navigator.clipboard;
  if (!clipboard?.writeText) return { failure, report, copied: false };
  try {
    await clipboard.writeText(report);
    return { failure, report, copied: true };
  } catch {
    return { failure, report, copied: false };
  }
}

/**
 * Wait this long between clearing a copy result and showing the next one, so a screen reader
 * reads the same result again. @internal
 */
export const COPY_RESULT_GAP_MS = 100;

const failureKeys = new WeakMap<DocumentOpenFailure, number>();
let nextFailureKey = 0;

/**
 * A key for each refusal: the notice renders again for a second refusal of the same file, so
 * it shows and reads again. @internal
 */
export function openFailureKey(failure: DocumentOpenFailure): number {
  let key = failureKeys.get(failure);
  if (key === undefined) failureKeys.set(failure, (key = nextFailureKey += 1));
  return key;
}

/**
 * The refused open that an open-failure part shows: none while the next file opens (the
 * loading screen shows then), and none once a deck opened. @internal
 */
export function shownOpenFailure(snapshot: EditorSnapshot): DocumentOpenFailure | null {
  return snapshot.status === 'loading' ? null : (snapshot.openFailure ?? null);
}
