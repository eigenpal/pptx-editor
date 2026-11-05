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

/**
 * Resolves after the browser paints the next frame. `load` awaits it between publishing the
 * `loading` state and the parse, which blocks the main thread for a large deck, so the
 * loading screen paints first. The spinner turns by a CSS transform, which the compositor
 * runs while the main thread is busy. Outside a browser it resolves at once.
 */
export function yieldToPaint(): Promise<void> {
  if (typeof requestAnimationFrame !== 'function' || typeof document === 'undefined')
    return Promise.resolve();
  // `requestAnimationFrame` runs before the paint; the task after it runs after the paint.
  return new Promise((resolve) => requestAnimationFrame(() => setTimeout(resolve, 0)));
}

/** The reason group of each engine open code (`EditorSnapshot.error`). */
const OPEN_FAILURE_REASONS: Readonly<Record<string, DocumentOpenFailureReason>> = Object.assign(
  Object.create(null) as Record<string, DocumentOpenFailureReason>,
  {
    zip: 'damaged',
    xml: 'damaged',
    'no-content-types': 'damaged',
    'not-a-presentation': 'not-pptx',
    empty: 'not-pptx',
    'too-many-parts': 'too-large',
  }
);

/**
 * The refusal a person sees for the engine's open code `code`, after DOCX Editor's
 * `documentOpenFailure`: frozen, and a new object for each call, so a second refusal of the
 * same file shows again. @internal
 */
export function documentOpenFailure(code: string): DocumentOpenFailure {
  return Object.freeze({ code, reason: OPEN_FAILURE_REASONS[code] ?? 'unknown' });
}

/** The i18n key of the message for each open failure reason. @internal */
export const OPEN_FAILURE_MESSAGE_KEYS = {
  damaged: 'openFailure.reasons.damaged',
  'not-pptx': 'openFailure.reasons.notPptx',
  'too-large': 'openFailure.reasons.tooLarge',
  unknown: 'openFailure.reasons.unknown',
} as const satisfies Record<DocumentOpenFailureReason, string>;

/**
 * The refused open that an open-failure part shows: none while the next file opens (the
 * loading screen shows then), and none once a deck opened. @internal
 */
export function shownOpenFailure(snapshot: EditorSnapshot): DocumentOpenFailure | null {
  return snapshot.status === 'loading' ? null : (snapshot.openFailure ?? null);
}
