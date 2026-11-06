// Field logic of the packaged dialogs that both adapters share: what a dialog shows for a
// document value, and how its draft becomes a command. Framework-free, so the React and Vue
// dialogs agree. Every export is adapter glue, not a stable API: `@internal`.

import type {
  ColorChoice,
  Editor,
  EditorCommand,
  SlideBackground,
  SlideBackgroundState,
} from '../contracts/index.ts';
import type { Translate } from '@pptx-editor.dev/i18n';
import { translateKeyOrText } from './adapter-support.ts';
import type { DialogCommitResult, DialogErrors } from './dialog-draft.ts';

/**
 * The draft of the Background dialog. `background` is the fill to apply: a fill, `null` for
 * the layout's background, or `undefined` to keep the slide's background as it is (a picture
 * or a pattern the snapshot cannot hand back). @internal
 */
export interface BackgroundDraftFields {
  readonly background: SlideBackground | null | undefined;
  readonly allSlides: boolean;
}

/**
 * The draft a Background dialog opens with: the slide's own solid, gradient, or pattern fill, `null`
 * when it follows the layout, or `undefined` for a fill the dialog cannot hand back. @internal
 */
export function backgroundDraftFrom(state: SlideBackgroundState): BackgroundDraftFields {
  if (state.inherited) return { background: null, allSlides: false };
  if (state.kind === 'solid' && state.color)
    return { background: { kind: 'solid', color: state.color }, allSlides: false };
  if (state.kind === 'gradient' && state.gradient)
    return { background: { kind: 'gradient', gradient: state.gradient }, allSlides: false };
  if (state.kind === 'pattern' && state.pattern)
    return { background: { kind: 'pattern', pattern: state.pattern }, allSlides: false };
  return { background: undefined, allSlides: false };
}

/**
 * Whether two background fills are the same choice, for the pressed swatch of the current
 * fill. Pictures compare by identity: their bytes are not compared. @internal
 */
export function sameBackground(
  a: SlideBackground | null | undefined,
  b: SlideBackground | null | undefined
): boolean {
  if (a === b) return true;
  if (!a || !b || a.kind !== b.kind || a.kind === 'picture') return false;
  return JSON.stringify(a) === JSON.stringify(b);
}

/** The solid color a draft shows as selected in the swatches, or null. @internal */
export function backgroundDraftColor(fields: BackgroundDraftFields): ColorChoice | null {
  return fields.background?.kind === 'solid' ? fields.background.color : null;
}

/**
 * Errors that block the Background draft. "All slides" needs a fill the user picked (or the
 * layout's background): the opening fill comes from the snapshot, which keeps only RGB and
 * hex gradient stops, so rewriting slides from it would drop theme colors, transforms, alpha,
 * and pictures. @internal
 */
export function validateBackgroundDraft(
  fields: BackgroundDraftFields,
  initial: BackgroundDraftFields,
  picked: boolean = fields.background !== initial.background
): DialogErrors<BackgroundDraftFields> {
  // Unpicked is the opening value itself: a swatch the user picks, even one equal to the
  // slide's fill, is built from the swatch, not from the snapshot.
  const unpicked = fields.background === undefined || (fields.background !== null && !picked);
  return fields.allSlides && unpicked ? { allSlides: 'background.pickFirst' } : {};
}

/**
 * The command a Background draft applies, or null when the user picked nothing. A pick (a
 * swatch, a picture, or Reset) applies even when it equals the opening fill: another editor
 * may have changed the slide since the dialog opened. The opening fill itself never applies:
 * it is the lossy snapshot. `picked` says whether the user picked; it defaults to a value
 * other than the opening one (by identity), which cannot see a Reset on a slide that already
 * followed the layout, so the dialogs pass it. @internal
 */
export function backgroundDraftCommand(
  fields: BackgroundDraftFields,
  initial: BackgroundDraftFields,
  picked: boolean = fields.background !== initial.background
): EditorCommand | null {
  if (fields.background === undefined) return null;
  if (!picked && (!fields.allSlides || fields.background !== null)) return null;
  return {
    type: 'setSlideBackground',
    background: fields.background,
    ...(fields.allSlides ? { allSlides: true } : {}),
  };
}

/**
 * Run a dialog's command as a draft commit. A refusal keeps the dialog open with the engine's
 * reason; a command with nothing to do counts as done. @internal
 */
export function commitDialogCommand(
  editor: Pick<Editor, 'exec'> | null,
  command: EditorCommand | null
): DialogCommitResult {
  if (!command) return { ok: true };
  if (!editor) return { ok: false, reason: 'disabled.noEditor' };
  const result = editor.exec(command);
  return result.ok || result.noop ? { ok: true } : { ok: false, reason: result.reason };
}

/**
 * A draft's errors as the user reads them. The draft keeps translation keys and engine text;
 * the state a dialog hands to its parts and hooks carries the translated messages, as DOCX
 * Editor's dialogs do. @internal
 */
export function translateDialogErrors<Fields extends object>(
  t: Translate,
  errors: DialogErrors<Fields>
): DialogErrors<Fields> {
  const out: Record<string, string> = {};
  for (const [name, message] of Object.entries(errors))
    if (typeof message === 'string') out[name] = translateKeyOrText(t, message);
  return out as DialogErrors<Fields>;
}

/**
 * The choice an arrow key moves to in a radio group (WAI-ARIA radio group pattern): Right and
 * Down go to the next choice, Left and Up to the previous one, wrapping. Null for other keys.
 * @internal
 */
export function radioGroupStep<T>(choices: readonly T[], current: T, key: string): T | null {
  const step =
    key === 'ArrowRight' || key === 'ArrowDown'
      ? 1
      : key === 'ArrowLeft' || key === 'ArrowUp'
        ? -1
        : 0;
  if (step === 0 || choices.length === 0) return null;
  const at = Math.max(0, choices.indexOf(current));
  return choices[(at + step + choices.length) % choices.length] ?? null;
}
