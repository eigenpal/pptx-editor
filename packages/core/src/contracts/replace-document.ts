// The document handle, the change event detail, and the `replaceDocument` command, after DOCX
// Editor's `DocumentHandle`, `DocumentChange`, `ReplaceDocumentRequest`, and
// `ReplaceDocumentOutcome` (docx-editor/packages/core/src/contracts/editor.ts and
// replace-document.ts at b63cbea4a). PPTX Editor's request takes `pptx` where DOCX Editor's
// takes `docx`. DOM-free.

/**
 * The open document at one moment: which load it came from and its revision. Compare two
 * handles to tell whether anyone edited the document in between, or whether another document
 * was loaded.
 *
 * @public
 */
export interface DocumentHandle {
  /** A new value for each `load()` and each refresh that reloads. */
  readonly loadId: string;
  /** The store revision. Every commit raises it: an edit, an undo, a redo, a peer's edit. */
  readonly revision: number;
}

/**
 * What caused a `change` event. Replacement events must not start another external processing
 * request.
 *
 * @public
 */
export interface DocumentChange {
  /**
   * `'load'` for a load, `'refresh'` for a document refresh or its `replaceDocument` edit, and
   * `'recovery'` for a refresh recovery. Absent for edits, undo, and redo, also a peer's.
   */
  readonly source?: 'load' | 'refresh' | 'recovery';
  /** The store revision after the change, the same number as `getDocumentHandle().revision`. */
  readonly revision: number;
}

/**
 * Replace the whole presentation with a `.pptx` file as one undoable edit.
 *
 * @example
 * ```ts
 * const expected = editor.getDocumentHandle();
 * const pptx = await renderTemplate(editor.save()!);
 * const result = editor.exec({ type: 'replaceDocument', pptx, expected });
 * if (!result.ok) console.warn(result.replaceDocument?.code, result.reason);
 * ```
 *
 * @public
 */
export interface ReplaceDocumentRequest {
  /** The new presentation's `.pptx` bytes. */
  readonly pptx: Uint8Array | ArrayBuffer;
  /** Refuse with `staleDocument` unless the open document is still the one this handle names. */
  readonly expected?: DocumentHandle;
}

/** The `replaceDocument` command. @public */
export type ReplaceDocumentCommand = { readonly type: 'replaceDocument' } & ReplaceDocumentRequest;

/**
 * Why a `replaceDocument` command was refused. Read this code rather than the top-level
 * `CommandResult.reason`, which is coarser.
 *
 * - `invalidRequest`: the command's arguments are wrong, or it was given a history group. A
 *   replacement is always its own undo step.
 * - `invalidDocument`: the bytes are not a valid PPTX file.
 * - `staleDocument`: the open document is not the one `expected` describes.
 * - `unavailable`: no document is mounted, the document is still opening, or the editor was
 *   destroyed.
 * - `busy`: a document refresh blocks edits. A later retry can succeed.
 * - `editNotPermitted`: viewing mode does not permit the edit. A retry does not help until that
 *   changes.
 * - `editUnsupported`: the difference cannot apply as one undoable edit.
 *
 * Every refusal leaves the document and its undo history unchanged.
 *
 * @public
 */
export type ReplaceDocumentErrorCode =
  | 'invalidRequest'
  | 'invalidDocument'
  | 'staleDocument'
  | 'unavailable'
  | 'busy'
  | 'editNotPermitted'
  | 'editUnsupported';

/**
 * The specific cause behind a refusal code, where one exists.
 *
 * - `viewing`: with `editNotPermitted`. `suggesting` and `protected` stay for DOCX Editor
 *   compatibility: PresentationML has no suggesting mode or enforced protection, so they do not
 *   occur.
 * - `otherDocument`, `edited`: with `staleDocument`. Another document was loaded since
 *   `expected`, or the same document was edited, by anyone.
 * - `collaborationUnavailable`: with `unavailable`. Kept for DOCX Editor compatibility; it does
 *   not occur.
 * - `differentMainPart`: with `editUnsupported`; the file names a different main part.
 * - `unsupportedDifference`: with `editUnsupported`; the difference cannot be expressed as an
 *   edit of the open document.
 * - `collaborationRefused`: with `editUnsupported`; the store refused the edit while a
 *   collaboration session is attached.
 *
 * @public
 */
export type ReplaceDocumentRefusalReason =
  | 'viewing'
  | 'suggesting'
  | 'protected'
  | 'collaborationUnavailable'
  | 'differentMainPart'
  | 'unsupportedDifference'
  | 'collaborationRefused'
  | 'otherDocument'
  | 'edited';

/**
 * Why `replaceDocument` was refused. `code` is stable; `reason` names the cause where one exists.
 * `message` is English, for logs: branch on `code` and `reason`, never on `message`.
 *
 * @public
 */
export type ReplaceDocumentRefusal = {
  readonly ok: false;
  readonly message: string;
} & (
  | {
      readonly code: 'editNotPermitted';
      /** Only `viewing` occurs: PresentationML has no suggesting mode or enforced protection. */
      readonly reason: 'viewing' | 'suggesting' | 'protected';
    }
  | {
      readonly code: 'editUnsupported';
      readonly reason: 'differentMainPart' | 'unsupportedDifference' | 'collaborationRefused';
    }
  | { readonly code: 'staleDocument'; readonly reason: 'edited' | 'otherDocument' }
  | { readonly code: 'unavailable'; readonly reason?: 'collaborationUnavailable' }
  | { readonly code: 'invalidRequest' | 'invalidDocument' | 'busy'; readonly reason?: undefined }
);

/**
 * The outcome of `replaceDocument`. `changed` is false when the file equals the open document:
 * nothing changed, and no undo step was added.
 *
 * @public
 */
export type ReplaceDocumentOutcome =
  { readonly ok: true; readonly changed: boolean } | ReplaceDocumentRefusal;
