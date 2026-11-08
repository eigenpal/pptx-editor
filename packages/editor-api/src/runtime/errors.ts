/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/editor-api/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Adapted from DOCX Editor's `packages/editor-api/src/runtime/errors.ts` (commit 5993bd223945).
//
// One error class, a stable code per refusal, and one fixed sentence per code. Nothing from the
// engine or the file reaches a message: host details only select a public diagnostic, whose
// fixed hint becomes `debugInfo.innerError`. The edit refusals of the engine have their
// diagnostics in `edit-diagnostics.ts`.

import { argumentHint } from './argument-hints.ts';
import { CODE_DIAGNOSTICS, EDIT_DIAGNOSTICS } from './edit-diagnostics.ts';

/**
 * What went wrong, as a value a consumer may branch on.
 *
 * @public
 */
export type PptxEditorErrorCode =
  /** A property was read before a `load(...)` for it completed in a `sync()`. */
  | 'PropertyNotLoaded'
  /** A `ClientResult` value was read before the sync that fills it. */
  | 'ValueNotLoaded'
  /**
   * The object cannot be addressed: its run ended and nothing tracked it, the call that would
   * have created it was in a refused sync, or the presentation no longer holds it (an edit
   * removed it). A lookup from a refused sync is sent again, so its proxy stays usable.
   */
  | 'InvalidObjectPath'
  /** The object still belongs to a run that has not finished, so it cannot be handed over. */
  | 'ObjectInUse'
  /** An argument or load option this API does not accept. */
  | 'InvalidArgument'
  /** Opening the presentation exceeded a bounded reader limit. */
  | 'ResourceLimitExceeded'
  /** The collection has no such item: `getItem` with an unknown id, `getItemAt` out of range. */
  | 'ItemNotFound'
  /**
   * The call is outside the supported subset for this host or content, and nothing changed.
   * Examples: a server runtime has no selection; a shape has no text frame.
   */
  | 'NotSupported'
  /** The member is one this API does not implement, and nothing changed. */
  | 'NotImplemented'
  /** The request context's `run` has finished, so it can no longer be used. */
  | 'InvalidRequestContext'
  /** The runtime was disposed. Every later operation fails this way. */
  | 'RuntimeDisposed'
  /**
   * The batch was built from an older read, so nothing was applied: another writer changed the
   * presentation after this context read it. Read again before retrying.
   */
  | 'StaleDocument'
  /** The host is live but holds no presentation right now: an editor between documents. */
  | 'DocumentUnavailable'
  /** The presentation refused the change, or answered something this runtime cannot use. */
  | 'GeneralException';

const MESSAGES: Readonly<Record<PptxEditorErrorCode, string>> = Object.freeze({
  PropertyNotLoaded:
    'the property has not been loaded. Call load(...) and await context.sync() before reading it.',
  ValueNotLoaded: 'the result has not been filled in yet. Await context.sync() before reading it.',
  InvalidObjectPath:
    'the object cannot be addressed. Keep proxies within their run, or use ' +
    'context.trackedObjects.add(...) before the run ends.',
  ObjectInUse:
    'the object still belongs to a run that has not finished. Await that run before passing the ' +
    'object to another one.',
  InvalidArgument: 'the argument is not one this API accepts.',
  ResourceLimitExceeded: 'the presentation exceeds a supported resource limit.',
  ItemNotFound: 'the collection has no such item.',
  NotSupported:
    'the operation or argument is outside the supported subset for this host or content.',
  NotImplemented: 'this API does not implement that operation for this target.',
  InvalidRequestContext: 'the request context has finished. Start another run to continue.',
  RuntimeDisposed: 'the runtime has been disposed.',
  StaleDocument:
    'the presentation changed after this context read it, so nothing was applied. Read again ' +
    'before you decide.',
  DocumentUnavailable: 'the presentation is not available right now.',
  GeneralException: 'the presentation could not complete the request.',
});

/** A public reason for a refusal: a stable code and a fixed hint. */
export interface PptxEditorDiagnostic {
  readonly code: string;
  readonly hint: string;
}

/** The diagnostic of an argument refusal that names what the API accepts there. */
const ARGUMENT_NOT_ACCEPTED = 'argument.not-accepted';

/**
 * The public diagnostics, by the host detail that selects them. A detail that is not here
 * gives no `innerError`: an unknown engine reason never becomes public API.
 */
const DIAGNOSTICS: ReadonlyMap<string, PptxEditorDiagnostic> = new Map(
  (
    [
      ['gone', 'edit.stale-target', 'An edit removed the object. Read the presentation again.'],
      [
        'null-object',
        'edit.null-object',
        'The object is a null object: its lookup found nothing. Check isNullObject first.',
      ],
      [
        'refused-sync',
        'sync.object-from-refused-sync',
        'This object came from a sync that was refused, so it was never created. Create it again.',
      ],
      ['disabled.viewing', 'edit.read-only', 'The editor is open for viewing. Switch to editing.'],
      ['selection', 'host.no-selection', 'Only a runtime over a live editor has a selection.'],
      [
        'no-text',
        'edit.no-text-frame',
        'The shape has no text frame. Check shape.type first, and read text only of a shape that ' +
          'holds text.',
      ],
      [
        'not-group',
        'edit.not-a-group',
        'The shape is not a group. Check shape.type first, and use a group shape.',
      ],
      [
        'group-count',
        'edit.group-count',
        'A group needs two or more different shapes. Remove repeated shapes or add more.',
      ],
      [
        'disabled.groupPlaceholder',
        'edit.group-placeholder',
        'A placeholder cannot be in a group. Group shapes that are not placeholders.',
      ],
      [
        'group-parent',
        'edit.group-parent',
        'The shapes are not in the same parent. Use shapes of one slide, or of one group.',
      ],
      [
        'group-owner',
        'edit.group-off-slide',
        'Only shapes on a slide can be grouped or ungrouped. Use a slide, not a layout or master.',
      ],
      [
        'group-line',
        'edit.group-line',
        'A group has no outline of its own. Set the line format of each shape in the group.',
      ],
      [
        'disabled.lastSlide',
        'edit.last-slide',
        'A presentation keeps at least one slide. Add a slide before you delete this one.',
      ],
      ['not-table', 'edit.not-a-table', 'The shape is not a table. Use a table shape.'],
      [
        'table-style-undefined',
        'edit.table-style-undefined',
        'The presentation does not define that table style. Choose a style that the ' +
          'presentation defines, or MediumStyle2Accent1, which is built in.',
      ],
      [
        'table-style-unknown',
        'read.table-style-unknown',
        'The table uses a style that has no TableStyle value. Read the other settings, or set ' +
          'a style first.',
      ],
      [
        'no-preset-geometry',
        'edit.no-preset-geometry',
        'The shape has no preset geometry, so it has no adjustments. Use a shape with a preset ' +
          'geometry, not a freeform, a group, or a table.',
      ],
      [
        'merge-overlap',
        'edit.merge-overlap',
        'The new span cuts another merged cell. Resize or split that cell first.',
      ],
      [
        'duplicate-id',
        'edit.duplicate-shape-id',
        'Another shape on the slide has the same id. A write could reach the wrong shape.',
      ],
      [
        'custom-xml',
        'edit.invalid-custom-xml',
        'The text is not one well-formed XML element, declares a DTD or an entity, declares an ' +
          'encoding other than UTF-8, or is longer than 8 MiB. Write plain XML with one root.',
      ],
      [
        'custom-xml-id',
        'edit.custom-xml-id-taken',
        'Another custom XML part of the presentation has that id. Use a new GUID.',
      ],
      [
        'custom-properties-root',
        'read.custom-properties-invalid',
        'The custom properties part of the file is not a custom properties part. Remove it first.',
      ],
      [
        'import-base64',
        'import.invalid-base64',
        'The file is not Base64 text, or it is larger than 100 MiB. Send the PPTX bytes as Base64.',
      ],
      [
        'import-package',
        'import.not-a-presentation',
        'The file is not a presentation, or it exceeds a limit of the bounded reader.',
      ],
      [
        'import-slide',
        'import.slide-not-found',
        'The source presentation has no slide with that id. Use the ids of its Slide objects.',
      ],
      [
        'import-target',
        'import.target-not-found',
        'The presentation has no slide with the targetSlideId. Read the slide ids again.',
      ],
      ['save', 'host.no-save', 'This runtime borrows an editor, which saves its own document.'],
      [
        'images',
        'host.no-images',
        'This runtime draws no pictures. Use a runtime over a live editor in a browser, or ' +
          'create the server runtime with a rasterize option.',
      ],
      [
        'image-owner',
        'image.not-on-slide',
        'Only a slide or a shape on a slide has a picture. Use a shape of a slide, not of a ' +
          'layout or a master.',
      ],
      [
        'image-size',
        'image.invalid-size',
        'The width and the height of a picture must be positive numbers of pixels.',
      ],
      [
        'image-failed',
        'image.failed',
        'The editor could not draw the picture: the shape is not drawn (for example, it is ' +
          'hidden), the browser has no canvas, or the rasterizer failed or answered something ' +
          'other than a PNG of the asked size. Read the shape, then try again.',
      ],
      [
        'address',
        ARGUMENT_NOT_ACCEPTED,
        'Use an http, https, mailto, tel, or ftp address of up to 2048 characters.',
      ],
      [
        'binding-shape',
        'binding.shape-deleted',
        'An edit deleted the shape of the binding. Delete the binding, or bind another shape ' +
          'with the same id.',
      ],
      [
        'binding-selection',
        'binding.selection',
        'The selection is not one shape. Select exactly one shape on a slide, then add the ' +
          'binding.',
      ],
      [
        'binding-limit',
        'binding.limit',
        'The runtime holds 10000 bindings. Delete a binding before you add another.',
      ],
      [
        'read-after-commit',
        'edit.read-after-commit',
        'The writes of this sync were applied, but a read after them failed. Read again, and ' +
          'do not apply the same edit again.',
      ],
    ] as const
  )
    .map(([detail, code, hint]): [string, PptxEditorDiagnostic] => [
      detail,
      Object.freeze({ code, hint }),
    ])
    .concat(EDIT_DIAGNOSTICS.map(([detail, code, hint]) => [detail, Object.freeze({ code, hint })]))
);

/**
 * The diagnostic of a host refusal: the one its detail selects, else the fixed one of its host
 * code, so every refusal of the engine carries an `innerError`.
 */
export function hostDiagnosticOf(
  code: string,
  detail: string | undefined
): PptxEditorDiagnostic | undefined {
  const own = diagnosticOf(detail);
  if (own) return own;
  // An argument the engine refused: the hint says what to pass (`argument-hints.ts`).
  const hint = code === 'invalid-argument' ? argumentHint(detail) : undefined;
  if (hint) return Object.freeze({ code: ARGUMENT_NOT_ACCEPTED, hint });
  const fallback = Object.prototype.hasOwnProperty.call(CODE_DIAGNOSTICS, code)
    ? CODE_DIAGNOSTICS[code]
    : undefined;
  return fallback ? Object.freeze({ code: fallback[0], hint: fallback[1] }) : undefined;
}

/**
 * A slide to import that the import cannot copy, with the ids of those slides when the engine
 * names them. The ids are numbers only, so nothing else from the file reaches the hint.
 */
const NOT_COPYABLE =
  /^disabled\.slideNotCopyable:import-(comments|control)(?::(\d{1,10}(?:,\d{1,10}){0,19}))?$/;

function notCopyable(detail: string): PptxEditorDiagnostic | undefined {
  const m = NOT_COPYABLE.exec(detail);
  if (!m) return undefined;
  const ids = m[2]?.split(',') ?? [];
  const one = ids.length <= 1;
  const slides =
    ids.length === 0
      ? 'A slide to import'
      : `${one ? 'Slide' : 'Slides'} ${ids.join(', ')} of the source`;
  const what =
    m[1] === 'comments'
      ? 'threaded comments, which an import cannot copy yet'
      : 'an embedded control (p:control)';
  return Object.freeze({
    code: m[1] === 'comments' ? 'import.slide-has-comments' : 'import.slide-has-controls',
    hint: `${slides} ${one ? 'has' : 'have'} ${what}. Leave ${one ? 'it' : 'them'} out with sourceSlideIds.`,
  });
}

/** The public diagnostic a host detail selects, if any. */
export function diagnosticOf(detail: string | undefined): PptxEditorDiagnostic | undefined {
  if (detail === undefined) return undefined;
  return DIAGNOSTICS.get(detail) ?? notCopyable(detail) ?? DIAGNOSTICS.get(detail.split(':')[0]!);
}

const REMOVED_OBJECT_MESSAGE =
  'the object no longer exists in the presentation. An edit removed it. Read the presentation ' +
  'again to get a fresh object.';

const NULL_OBJECT_MESSAGE =
  'the object is a null object, so it cannot be written or called. Check isNullObject after a ' +
  'sync before you use an object from a get...OrNullObject method.';

function messageOf(
  code: PptxEditorErrorCode,
  diagnostic: PptxEditorDiagnostic | undefined
): string {
  if (code === 'InvalidArgument' && diagnostic?.code === ARGUMENT_NOT_ACCEPTED)
    return `${MESSAGES[code]} ${diagnostic.hint}`;
  if (code !== 'InvalidObjectPath') return MESSAGES[code];
  if (diagnostic?.code === 'edit.stale-target') return REMOVED_OBJECT_MESSAGE;
  if (diagnostic?.code === 'edit.null-object') return NULL_OBJECT_MESSAGE;
  return MESSAGES[code];
}

/**
 * The fields a {@link PptxEditorError} is constructed from.
 *
 * @public
 */
export interface PptxEditorErrorInit {
  /** Which refusal this is. The stable thing to branch on. */
  readonly code: PptxEditorErrorCode;
  /** The consumer-facing path of the object or property involved, never a handle. */
  readonly target?: string;
  /** The revision the context had read at, for `StaleDocument`. */
  readonly expectedRevision?: number;
  /** The revision the presentation was actually at, for `StaleDocument`. */
  readonly actualRevision?: number;
  /** The public reason, when the engine gave one this API knows. */
  readonly diagnostic?: PptxEditorDiagnostic;
  /** The `context.trace` messages queued before the call that failed. */
  readonly traceMessages?: readonly string[];
}

/**
 * Why the engine refused, as a public diagnostic code with its fixed hint.
 *
 * @public
 */
export interface PptxEditorInnerError {
  /** A stable public diagnostic code, such as `edit.stale-target`. */
  readonly code: string;
  /** The fixed hint for the code: the cause, then the action. Never presentation content. */
  readonly message: string;
}

/**
 * Debug details in the shape of Office.js `OfficeExtension.DebugInfo`.
 *
 * @public
 */
export interface PptxEditorDebugInfo {
  /** The same code as {@link PptxEditorError.code}. */
  readonly code: PptxEditorErrorCode;
  /** The same message as the error's own. */
  readonly message: string;
  /** The consumer-facing path of the object or property involved, when there is one. */
  readonly errorLocation?: string;
  /** Why the engine refused, when it knows a specific reason. */
  readonly innerError?: PptxEditorInnerError;
}

/**
 * Every refusal this runtime throws.
 *
 * Branch on {@link PptxEditorError.code}, never on the message. Codes are stable public API.
 * Nothing from the engine or the file appears in the message: a consumer gets a stable code, a
 * fixed sentence, and `target`, the consumer-facing path they wrote themselves.
 *
 * @example
 * ```ts
 * try {
 *   await context.sync();
 * } catch (error) {
 *   if (isPptxEditorError(error) && error.code === 'StaleDocument') {
 *     // Read again: another writer changed the presentation first.
 *   }
 * }
 * ```
 *
 * @public
 */
export class PptxEditorError extends Error {
  /** Which refusal this is. Stable across versions; the thing to branch on. */
  readonly code: PptxEditorErrorCode;
  /** Consumer-facing path of the object or property involved, when there is one to name. */
  readonly target?: string;
  /** For `StaleDocument`: the revision the context had last read at. */
  readonly expectedRevision?: number;
  /** For `StaleDocument`: the revision the presentation was actually at. */
  readonly actualRevision?: number;
  /** Debug details, as Office.js `OfficeExtension.Error.debugInfo` has them. */
  readonly debugInfo: PptxEditorDebugInfo;
  /**
   * The messages of `context.trace` calls queued before the call that failed, in order, as
   * Office.js `OfficeExtension.Error.traceMessages`. Empty when the sync had none.
   */
  readonly traceMessages: readonly string[];

  constructor(init: PptxEditorErrorInit) {
    const base = messageOf(init.code, init.diagnostic);
    const message = init.target ? `${base} (${init.target})` : base;
    super(message);
    this.name = 'PptxEditorError';
    this.code = init.code;
    this.traceMessages = Object.freeze([...(init.traceMessages ?? [])]);
    this.debugInfo = Object.freeze({
      code: init.code,
      message,
      ...(init.target === undefined ? {} : { errorLocation: init.target }),
      ...(init.diagnostic === undefined
        ? {}
        : {
            innerError: Object.freeze({
              code: init.diagnostic.code,
              message: init.diagnostic.hint,
            }),
          }),
    });
    if (init.target !== undefined) this.target = init.target;
    if (init.expectedRevision !== undefined) this.expectedRevision = init.expectedRevision;
    if (init.actualRevision !== undefined) this.actualRevision = init.actualRevision;
  }
}

/**
 * Whether a caught value is one of ours. By `name` as well as by `instanceof`: a consumer can
 * end up with two copies of this module, and `instanceof` fails across them.
 *
 * @public
 */
export function isPptxEditorError(value: unknown): value is PptxEditorError {
  if (value instanceof PptxEditorError) return true;
  return (
    typeof value === 'object' &&
    value !== null &&
    (value as { name?: unknown }).name === 'PptxEditorError' &&
    typeof (value as { code?: unknown }).code === 'string'
  );
}

/** The message this runtime uses for a code, for tests that assert nothing else is thrown. */
export function messageFor(code: PptxEditorErrorCode): string {
  return MESSAGES[code];
}

export function fail(init: PptxEditorErrorInit): never {
  throw new PptxEditorError(init);
}

/** The same refusal with the trace messages of its sync. Any other value comes back as it was. */
export function withTraceMessages(error: unknown, messages: readonly string[]): unknown {
  if (!(error instanceof PptxEditorError) || messages.length === 0) return error;
  const inner = error.debugInfo.innerError;
  return new PptxEditorError({
    code: error.code,
    ...(error.target === undefined ? {} : { target: error.target }),
    ...(error.expectedRevision === undefined ? {} : { expectedRevision: error.expectedRevision }),
    ...(error.actualRevision === undefined ? {} : { actualRevision: error.actualRevision }),
    ...(inner === undefined ? {} : { diagnostic: { code: inner.code, hint: inner.message } }),
    traceMessages: messages,
  });
}

/**
 * Refuse an argument with `InvalidArgument`, naming what the API accepts there, such as
 * `a string, a finite number, a boolean, or a valid Date`. The message and the diagnostic
 * `argument.not-accepted` both carry it.
 */
export function failArgument(target: string | undefined, accepted: string, guess?: string): never {
  const hint =
    guess === undefined ? `Use ${accepted}.` : `Did you mean '${guess}'? Use ${accepted}.`;
  throw new PptxEditorError({
    code: 'InvalidArgument',
    ...(target === undefined ? {} : { target }),
    diagnostic: Object.freeze({ code: ARGUMENT_NOT_ACCEPTED, hint }),
  });
}

/** Longer values get no suggestion: they are no typo of a name. */
const MAX_GUESSED = 64;

/**
 * The edit distance between `a` and `b`, ignoring case: insertions, deletions, substitutions,
 * and swaps of two neighbors each count one (optimal string alignment).
 */
function distance(a: string, b: string): number {
  const x = a.toLowerCase();
  const y = b.toLowerCase();
  let before: number[] = [];
  let row = Array.from({ length: y.length + 1 }, (_, j) => j);
  for (let i = 1; i <= x.length; i += 1) {
    const next = [i];
    for (let j = 1; j <= y.length; j += 1) {
      let d = Math.min(
        row[j]! + 1,
        next[j - 1]! + 1,
        row[j - 1]! + (x[i - 1] === y[j - 1] ? 0 : 1)
      );
      if (i > 1 && j > 1 && x[i - 1] === y[j - 2] && x[i - 2] === y[j - 1])
        d = Math.min(d, before[j - 2]! + 1);
      next[j] = d;
    }
    before = row;
    row = next;
  }
  return row[y.length]!;
}

/** The words of a name: `ArabicNumeralPeriod` is `arabic`, `numeral`, `period`. */
function words(name: string): string[] {
  return name
    .replace(/([a-z0-9])([A-Z])/g, '$1 $2')
    .replace(/([A-Z]+)([A-Z][a-z])/g, '$1 $2')
    .toLowerCase()
    .split(/[^a-z0-9]+/)
    .filter((w) => w.length > 0);
}

/** Two words match when equal, when one starts the other (3 or more letters), or by one typo. */
function sameWord(a: string, b: string): boolean {
  if (a === b) return true;
  if ((a.length >= 3 && b.startsWith(a)) || (b.length >= 3 && a.startsWith(b))) return true;
  return Math.min(a.length, b.length) >= 4 && distance(a, b) <= 1;
}

/**
 * The candidate nearest to `value`, or undefined when none is near. A candidate is always one
 * of `candidates`, so the suggestion is a value the caller may use.
 *
 * First a typo: the candidate nearest by edit distance, when that is at most a third of its
 * length, and at least 1 and at most 3 edits. Else the words: a candidate that holds every word
 * of `value` (`ArabicPeriod` is in `ArabicNumeralPeriod`; `Rect` starts `Rectangle`), or all
 * but one of three or more. Among those, the one with the fewest other words wins, then the
 * nearest by edit distance, then the first.
 */
export function nearest(value: unknown, candidates: readonly string[]): string | undefined {
  if (typeof value !== 'string' || value.length === 0 || value.length > MAX_GUESSED)
    return undefined;
  let best: string | undefined;
  let bestDistance = Infinity;
  for (const candidate of candidates) {
    const d = distance(value, candidate);
    if (d < bestDistance) {
      best = candidate;
      bestDistance = d;
    }
  }
  if (best === undefined) return undefined;
  const limit = Math.min(3, Math.max(1, Math.floor(best.length / 3)));
  if (bestDistance <= limit) return best;
  const wanted = words(value);
  if (wanted.length === 0) return undefined;
  let chosen: { candidate: string; missing: number; extra: number; d: number } | undefined;
  for (const candidate of candidates) {
    const have = words(candidate);
    const used = new Set<number>();
    let missing = 0;
    for (const w of wanted) {
      const at = have.findIndex((h, i) => !used.has(i) && sameWord(w, h));
      if (at < 0) missing += 1;
      else used.add(at);
    }
    if (missing > (wanted.length >= 3 ? 1 : 0)) continue;
    const score = {
      candidate,
      missing,
      extra: have.length - used.size,
      d: distance(value, candidate),
    };
    if (
      !chosen ||
      score.missing < chosen.missing ||
      (score.missing === chosen.missing &&
        (score.extra < chosen.extra || (score.extra === chosen.extra && score.d < chosen.d)))
    )
      chosen = score;
  }
  return chosen?.candidate;
}

/** Up to this many accepted values are listed in full. */
const LISTED_VALUES = 12;

/**
 * `one of A, B, or C`: the accepted values of an enum, in full when there are few, else the
 * count and the first few.
 */
export function oneOf(values: readonly string[]): string {
  const quoted = values.map((v) => `'${v}'`);
  if (quoted.length === 1) return quoted[0]!;
  if (quoted.length > LISTED_VALUES)
    return `one of the ${quoted.length} values of the enum, such as ${quoted.slice(0, 3).join(', ')}`;
  return `one of ${quoted.slice(0, -1).join(', ')}${quoted.length > 2 ? ',' : ''} or ${quoted.at(-1)}`;
}
