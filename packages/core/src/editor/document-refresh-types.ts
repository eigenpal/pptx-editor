// The document refresh API types, copied from DOCX Editor's
// packages/core/src/editor/document-refresh-types.ts at b63cbea4a with its comments. PPTX
// Editor differences: `RefreshLocation` names a slide, a shape, and a text range; `RefreshChange`
// adds `computed`; `RefreshHighlightOptions` adds `surfaces`; the diagnostic codes add the slide
// and shape codes. Every DOCX Editor member stays, so host code compiles against both.

/** Opacity fade settings. Reduced motion caps fades at 125ms. @public */
export interface AnchorHighlightAnimation {
  /** Entrance fade, in milliseconds: 0 through 10000. Default: 180. */
  readonly durationMs?: number;
  /** Exit fade, in milliseconds: 0 through 10000. Default: `durationMs`. */
  readonly exitDurationMs?: number;
  /** A CSS easing function. CSS variables are refused. Default: `--doc-motion-ease-out`. */
  readonly easing?: string;
}

/** Presentation of a temporary highlight, as DOCX Editor's anchor highlights. @public */
export interface AnchorHighlightOptions {
  /** `true` (default) fades; `false` changes at once; an object sets the timing. */
  readonly animation?: boolean | AnchorHighlightAnimation;
  /** CSS border color. Default: `color`. */
  readonly borderColor?: string;
  /** Corner radius in CSS pixels at 100% zoom. Default: 6. */
  readonly borderRadius?: number;
  /** `'solid'` (default), `'dashed'`, or `'dotted'`. */
  readonly borderStyle?: 'solid' | 'dashed' | 'dotted';
  /** Border width in CSS pixels at 100% zoom. Default: 0 for text, 2 for shapes and slides. */
  readonly borderWidth?: number;
  /** CSS classes for shadows, outlines, or patterns. */
  readonly className?: string;
  /** CSS color or variable. Default: `var(--doc-refresh-highlight-color)`. */
  readonly color?: string;
  /** Fill opacity between 0 and 1. Default: 0.14. */
  readonly opacity?: number;
  /** Extra space on each edge, in CSS pixels at 100% zoom. Default: 4. */
  readonly padding?: number;
  /** Milliseconds before dismissal; `null` keeps it until you clear it. Default: 3000. */
  readonly timeoutMs?: number | null;
}

/** Options of an explicit highlight dismissal. @public */
export interface ClearAnchorHighlightOptions {
  /** `false` removes at once; an object sets the exit timing. Default: the last highlight's. */
  readonly animation?: boolean | Pick<AnchorHighlightAnimation, 'exitDurationMs' | 'easing'>;
}

/** Where a revealed location lands. @public */
export type ScrollPlacement = 'start' | 'center' | 'centerIfNeeded' | 'nearest';

/** Options of an explicit scroll to a location. @public */
export interface ScrollToAnchorOptions {
  /** Target alignment. Default: `'center'`. */
  readonly block?: ScrollPlacement;
  /** `'smooth'` for an explicit animated action. Default: `'instant'`. */
  readonly behavior?: 'instant' | 'smooth';
  /** Edge padding for `'start'` and `'nearest'`, in CSS pixels. Default: 24. */
  readonly offsetPx?: number;
}

/**
 * A location in the returned presentation: a slide, a shape on it, or a text range in one
 * paragraph of that shape. Each level is valid alone.
 *
 * @public
 */
export interface RefreshLocation {
  /** `p:sldId/@id` of the slide: 256 through 2147483647. Use this or `slideIndex`, not both. */
  readonly slideId?: number;
  /** Zero-based slide index in presentation order, hidden slides included. */
  readonly slideIndex?: number;
  /** `p:cNvPr/@id` of a shape on the slide. Omit to locate the whole slide. */
  readonly shapeId?: number;
  /** Zero-based paragraph index in the shape's text body. Requires `shapeId`. */
  readonly paragraphIndex?: number;
  /** UTF-16 offsets in that paragraph. Require `paragraphIndex`. */
  readonly start?: number;
  readonly end?: number;
  /** Exact returned text at these offsets. Required when offsets are present. */
  readonly text?: string;
  /** The notes slide of the slide instead of the slide itself. Default: false. */
  readonly notes?: boolean;
}

/** One processor change. Keep its id stable across cumulative results. @public */
export interface RefreshChangeInput {
  readonly id: string;
  /** Processor-provided review summary, including changes without a location. Maximum: 1000 characters. */
  readonly description?: string;
  readonly location?: RefreshLocation;
  /** Deletions without a surviving range have no highlight. */
  readonly unavailableReason?: 'deleted' | 'unavailable';
}

/** What a computed change changed. @public */
export type RefreshChangeAspect =
  | 'position'
  | 'size'
  | 'rotation'
  | 'geometry'
  | 'fill'
  | 'line'
  | 'effects'
  | 'text'
  | 'textFormat'
  | 'picture'
  | 'table'
  | 'chart'
  | 'diagram'
  | 'transition'
  | 'animation'
  | 'background'
  | 'layout'
  | 'other';

/**
 * A change the review module computed by comparing the captured and returned presentations, in
 * fields that an adapter can describe in any language.
 *
 * @public
 */
export interface RefreshComputedChange {
  readonly scope: 'slide' | 'shape' | 'text' | 'notes' | 'layout' | 'master' | 'theme' | 'package';
  readonly action: 'added' | 'removed' | 'moved' | 'changed';
  /** For `changed`: what changed. */
  readonly aspects?: readonly RefreshChangeAspect[];
  /** For `layout`, `master`, and `theme`: zero-based indexes of the slides that show the change. */
  readonly affectedSlides?: readonly number[];
  /** For `removed` and `moved` slides: the zero-based slide index in the captured presentation. */
  readonly previousSlideIndex?: number;
}

/** A computed change in the form `update.changes` takes. @public */
export type ComputedRefreshChange = RefreshChangeInput & {
  readonly computed: RefreshComputedChange;
};

/** Why a change cannot receive a highlight. @public */
export type RefreshDiagnosticCode =
  | 'missing-location'
  | 'deleted'
  | 'invalid-selector'
  | 'invalid-paragraph-id'
  | 'paragraph-not-found'
  | 'ambiguous-paragraph-id'
  | 'invalid-offsets'
  | 'text-mismatch'
  | 'not-rendered'
  | 'unsupported-story'
  | 'multiple-paragraphs'
  | 'invalid-slide-id'
  | 'slide-not-found'
  | 'shape-not-found'
  | 'ambiguous-shape-id';

/** One validated change for an accepted result. @public */
export interface RefreshChange {
  readonly id: string;
  /** Processor-provided summary. Computed changes have none: describe them from `computed`. */
  readonly description?: string;
  /** Identity of the accepted submission and processor sequence. */
  readonly resultId: string;
  /** True for a new change id, changed text, or changed description in an existing change id. */
  readonly isNew: boolean;
  readonly status: 'available' | 'deleted' | 'unavailable' | 'invalid' | 'unsupported-story';
  readonly location?: RefreshLocation;
  /** Why this change cannot receive a highlight. Includes the requested location when supplied. */
  readonly diagnostic?: {
    readonly code: RefreshDiagnosticCode;
    readonly location?: RefreshLocation;
  };
  /** Present when the review module computed this change. Absent for processor changes. */
  readonly computed?: RefreshComputedChange;
}

/** A document captured for external processing. Treat this object as an opaque token. @public */
export interface RefreshSubmission {
  readonly id: string;
  readonly bytes: ArrayBuffer;
}

/**
 * How `applyUpdate()` puts returned bytes on screen.
 *
 * - `'reload'` (default): show the returned file as a reload does. Resets selection and undo
 *   history. Refused in collaboration sessions.
 * - `'edit'`: apply the difference as ONE undoable edit on top of the current history. Undo
 *   returns to the captured submission (or the previous accepted result), and redo applies the
 *   result again. Works in collaboration sessions, where the update reaches other participants
 *   as an ordinary edit. Refuses with `edit-unsupported` when the difference cannot apply as an
 *   edit; the document and its history stay unchanged, so outside a collaboration session you
 *   can retry with `'reload'`. Refuses with `edit-not-permitted` in viewing mode.
 *
 * Both modes keep the old paint on screen until the new one is ready, and never show the
 * loading screen.
 *
 * @public
 */
export type RefreshUpdateMode = 'reload' | 'edit';

/** Cumulative output for one submission. Assign sequence at the processor, before delivery. @public */
export interface RefreshUpdate {
  readonly submission: RefreshSubmission;
  readonly sequence: number;
  readonly bytes: ArrayBuffer | Uint8Array;
  /**
   * `'edit'` applies the result as one undoable edit; `'reload'` (default) reloads the document.
   * See {@link RefreshUpdateMode}.
   *
   * @example
   * ```ts
   * let result = await refresh.applyUpdate({ submission, sequence: 1, bytes, mode: 'edit' });
   * if (!result.ok && result.code === 'edit-unsupported') {
   *   result = await refresh.applyUpdate({ submission, sequence: 1, bytes });
   * }
   * ```
   */
  readonly mode?: RefreshUpdateMode;
  /** Omit to compute changes when `reviewModule()` is registered. */
  readonly changes?: readonly RefreshChangeInput[];
  /** Stable failed operation ids. The controller reports these and never retries processor work. */
  readonly failures?: readonly string[];
}

/**
 * Stable refusal codes for external document replacement.
 *
 * - `local-edits`: the document changed after capture, even through undo. With `mode: 'edit'`
 *   in a collaboration session, another participant's edit counts too.
 * - `collaboration`: a reload while a collaboration session is attached.
 * - `unavailable`: no mounted document, the document is still opening, or the collaboration
 *   session cannot accept edits now. A later retry can succeed.
 * - `edit-not-permitted`: the editor does not permit an update with `mode: 'edit'`. A retry
 *   does not help until that changes.
 * - `edit-unsupported`: the result cannot apply as one undoable edit, or the collaboration
 *   room refused it. The document and its history are unchanged.
 *
 * The document refresh guide describes every code and the usual response.
 *
 * @public
 */
export type RefreshFailureCode =
  | 'unavailable'
  | 'collaboration'
  | 'busy'
  | 'cancelled'
  | 'superseded'
  | 'document-changed'
  | 'local-edits'
  | 'out-of-order'
  | 'invalid-result'
  | 'invalid-document'
  | 'input-failed'
  | 'load-failed'
  | 'recovery-failed'
  | 'edit-not-permitted'
  | 'edit-unsupported';

/**
 * The specific cause behind a refusal code, for an update with `mode: 'edit'`.
 *
 * With code `edit-not-permitted`:
 * - `viewing`: the editor is in viewing mode.
 * - `protected`: not produced; PresentationML has no enforced edit protection.
 *
 * With code `unavailable`:
 * - `collaboration-unavailable`: the collaboration session cannot accept edits now.
 *
 * With code `edit-unsupported`:
 * - `different-main-part`: the returned file names a different main presentation part.
 * - `unsupported-difference`: the difference cannot be expressed as an edit of the open
 *   presentation.
 * - `collaboration-refused`: the collaboration room refused the edit.
 *
 * @public
 */
export type RefreshRefusalReason =
  | 'viewing'
  | 'protected'
  | 'collaboration-unavailable'
  | 'different-main-part'
  | 'unsupported-difference'
  | 'collaboration-refused';

/** Fields every refused {@link RefreshResult} carries. @public */
export interface RefreshRefusalFields {
  readonly ok: false;
  readonly resultId: string;
  /**
   * Why the update was refused, in English, for logs and diagnostics. Never parse it: branch
   * on `code` and `reason`. `DocumentRefreshError` from `capture()` has the same pair:
   * a stable `code` and an English `message`.
   */
  readonly message: string;
  /** For `load-failed`: whether the previous document was restored. */
  readonly recovered?: boolean;
}

/**
 * A refused update. `code` is stable; `reason` names the specific cause where one exists, and
 * narrows `code`: checking `reason === 'collaboration-refused'` implies `edit-unsupported`.
 *
 * @public
 */
export type RefreshRefusal =
  | (RefreshRefusalFields & {
      readonly code: 'edit-not-permitted';
      readonly reason: 'viewing' | 'protected';
    })
  | (RefreshRefusalFields & {
      readonly code: 'edit-unsupported';
      readonly reason: 'different-main-part' | 'unsupported-difference' | 'collaboration-refused';
    })
  | (RefreshRefusalFields & {
      readonly code: 'unavailable';
      readonly reason?: 'collaboration-unavailable';
    })
  | (RefreshRefusalFields & {
      readonly code: Exclude<
        RefreshFailureCode,
        'edit-not-permitted' | 'edit-unsupported' | 'unavailable'
      >;
      readonly reason?: undefined;
    });

/** Completion belongs to one result, including refusals and recovery failures. @public */
export type RefreshResult =
  | {
      readonly ok: true;
      readonly resultId: string;
      /** How the result was applied: an undoable edit, or a reload. */
      readonly mode: RefreshUpdateMode;
      readonly changes: readonly RefreshChange[];
      readonly changeInformation: 'available' | 'unavailable';
      readonly failures: readonly string[];
    }
  | RefreshRefusal;

/** Observable controller state. Document editing stays available during processing. @public */
export interface DocumentRefreshState {
  readonly phase:
    'idle' | 'capturing' | 'processing' | 'refreshing' | 'recovering' | 'complete' | 'failed';
  readonly result: RefreshResult | null;
  readonly changes: readonly RefreshChange[];
  /** Whether highlights are requested for available locations. False when the whole set starts dismissal. */
  readonly highlightsVisible: boolean;
  /** Offer restore and download controls when true. */
  readonly recoveryAvailable: boolean;
}

/** Opacity fade settings. Reduced motion caps fades at 125ms. @public */
export type RefreshHighlightAnimation = AnchorHighlightAnimation;

/** Where changes flash. @public */
export type RefreshHighlightSurface = 'thumbnails' | 'canvas' | 'notes';

/** Presentation for temporary change highlights. Each call starts from these defaults. @public */
export interface RefreshHighlightOptions extends AnchorHighlightOptions {
  /** Include earlier changes from the latest cumulative result. Default: false. */
  readonly includePrevious?: boolean;
  /** Select these change IDs instead of the recent/all filter. Unknown or unavailable IDs are skipped. Empty means none. */
  readonly changeIds?: readonly string[];
  /** Where changes flash: the slide rail and sorter thumbnails, the canvas, the notes pane. Default: all. */
  readonly surfaces?: readonly RefreshHighlightSurface[];
}

/** Explicit highlight dismissal. Document edits always remove stale highlights immediately. @public */
export type ClearRefreshHighlightsOptions = ClearAnchorHighlightOptions;

/** Explicit change navigation. Does not change the default view preservation during refresh. @public */
export interface NavigateToChangeOptions extends ScrollToAnchorOptions {
  /** Select the changed shape, or put the caret at the changed range start, and focus. Default: false. */
  readonly focus?: boolean;
  /** Target alignment in the slide rail. Default: center. centerIfNeeded keeps a visible thumbnail where it is. */
  readonly block?: ScrollPlacement;
}

/**
 * External file transport stays in your application. A reload resets selection and undo
 * history; an update with `mode: 'edit'` keeps both and adds one undo step.
 * @public
 */
export interface DocumentRefresh {
  /** Finish pending input and capture bytes with their document revision. Supersedes earlier submissions. */
  capture(): Promise<RefreshSubmission>;
  /**
   * Accept cumulative output. Keep the current slide and the zoom without moving focus. With
   * `mode: 'edit'`, apply it as one undoable edit and keep selection where it maps.
   */
  applyUpdate(update: RefreshUpdate): Promise<RefreshResult>;
  /** Cancel pending processing. An accepted result stays loaded. */
  cancel(): void;
  /** End processing when no further output will arrive. */
  finish(submission: RefreshSubmission): void;
  /** Cached state, suitable for framework external-store subscriptions. */
  snapshot(): DocumentRefreshState;
  /** Observe state changes. Returns an unsubscribe function. */
  subscribe(listener: () => void): () => void;
  /** Completion notification for every `applyUpdate()` call, including a refused result. */
  onResult(listener: (result: RefreshResult) => void): () => void;
  /** Highlight available changes and return their count, including other slides. Zero means none. Always validates options. */
  highlightChanges(options?: RefreshHighlightOptions): number;
  /** Remove temporary highlights without changing document content or history. */
  clearHighlights(options?: ClearRefreshHighlightsOptions): void;
  /** Explicit navigation. Does not focus unless requested. Returns false for unavailable locations. */
  navigateToChange(id: string, options?: NavigateToChangeOptions): boolean;
  /** Return a copy for download after recovery fails. */
  recoveryBytes(): ArrayBuffer | null;
  /** Retry recovery only while the failed document session is still active. */
  recover(): Promise<boolean>;
}

/** A capture failure with a stable code. `applyUpdate()` failures use RefreshResult instead. @public */
export class DocumentRefreshError extends Error {
  /** Machine-readable reason. */
  readonly code: RefreshFailureCode;
  constructor(code: RefreshFailureCode, cause?: unknown) {
    super(`Document capture failed: ${code}.`, { cause });
    this.name = 'DocumentRefreshError';
    this.code = code;
  }
}
