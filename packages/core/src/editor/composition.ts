// The state of one IME composition in the surface's input, and where its text goes in.
//
// A composition changes the model only when it ends. Its place is fixed when it starts: the
// caret's paragraph and offset, and the part of the input's value that the system selected
// for it (a keyboard that composes a word again selects that word). Every change it makes
// applies there, and nowhere else. When the model caret leaves that place (a click or a tap,
// an arrow key, the API, a slide change), the surface makes the browser end the composition
// (`surface-input.ts`), and the browser's final text goes to the place. An edit that keeps the
// caret in the paragraph (a peer's, or the API's) moves the place with the caret.
//
// The transitions are pure; `applyComposition` writes through the editor's text input.

import type { EditorSelection } from '../contracts/index.ts';
import {
  compositionEdit,
  fieldEdit,
  type ContextEdit,
  type ReplacementPlace,
  type ReplacementTarget,
  type TypingContext,
} from './replacement-text.ts';

/** Where a composition started. */
export interface CompositionPlace {
  /** The input's context when the composition started: its paragraph key, text, and caret. */
  readonly context: TypingContext | null;
  /** The input's value then: the context text, or empty when the caret had no context. */
  readonly value: string;
  /** The input selection that the system made for the composition, in `value`. */
  readonly from: number;
  readonly to: number;
}

/**
 * Idle; composing at a place (`seen`: the context at the place, moved by edits since the
 * start); or ending: the surface asked the browser to end the composition, and waits for
 * its final text.
 */
export type CompositionState =
  | { readonly phase: 'idle' }
  | {
      readonly phase: 'composing' | 'ending';
      readonly place: CompositionPlace;
      readonly seen: TypingContext | null;
      /** The input's text width when the composition started, for the overlay. */
      readonly width: number;
    };

export const IDLE: CompositionState = { phase: 'idle' };

/** The state when a composition starts, from the input as the system left it. */
export function startComposition(
  context: TypingContext | null,
  value: string,
  from: number,
  to: number,
  width: number
): CompositionState {
  return { phase: 'composing', place: { context, value, from, to }, seen: context, width };
}

/**
 * What a new context during a composition means. `same`: the caret is still at the place.
 * `follow`: an edit moved the caret in the place's paragraph, and the place moves with it.
 * `left`: the caret left the place, so the composition must end.
 */
export function contextChange(
  state: CompositionState,
  now: TypingContext | null
): 'same' | 'follow' | 'left' {
  if (state.phase === 'idle') return 'same';
  const { seen } = state;
  if (now?.key === seen?.key && now?.text === seen?.text) return 'same';
  if (now && seen && now.key === seen.key && now.revision !== seen.revision) return 'follow';
  return 'left';
}

/** The state after an edit moved the place (`contextChange` gave `follow`). */
export function followEdit(state: CompositionState, now: TypingContext): CompositionState {
  return state.phase === 'idle' ? state : { ...state, seen: now };
}

/** The state while the surface waits for the browser's final text. */
export function ending(state: CompositionState): CompositionState {
  return state.phase === 'composing' ? { ...state, phase: 'ending' } : state;
}

/**
 * Where the composition's text goes, for the caret context `now`. At the caret, while the
 * caret is the place (`typed`: added characters whose target an edit moved still go in at
 * the caret). At the place's own caret (`at`) when the caret left it. Null: no text caret.
 */
export function placement(
  state: CompositionState,
  now: TypingContext | null
): ReplacementPlace | null {
  if (state.phase === 'idle' || !state.seen) return null;
  const { seen } = state;
  const edited = now?.revision !== seen.revision;
  if (now?.key === seen.key && (now.text === seen.text || edited)) return { typed: true };
  return seen.selection ? { at: seen.selection } : null;
}

/**
 * The change the composition makes to its place's context: from the value when it started
 * to the final `value`, with the caret after the committed `data`. When the field does not
 * hold `data` at the caret, `data` replaces what the system selected for the composition
 * when it started (nothing, for a composition at the caret).
 */
export function finalEdit(
  place: CompositionPlace,
  value: string,
  caret: number,
  data: string
): ContextEdit | null {
  const key = place.context?.key ?? null;
  if (key === null) return null;
  const edit =
    place.value && (data || value !== place.value)
      ? compositionEdit(place.value, value, caret, data, key)
      : null;
  if (edit) return edit;
  const { value: from, from: start, to } = place;
  const target = { text: from.slice(start, to), after: from.length - to, between: from.slice(to) };
  return { target: { ...target, paragraph: key }, text: data };
}

/**
 * The change the composition made so far, for an end with no final text from the browser
 * (`finalEdit` with the field's own text). Null when nothing changed.
 */
export function partialEdit(place: CompositionPlace, value: string): ContextEdit | null {
  const key = place.context?.key;
  if (key === undefined || value === place.value) return null;
  return fieldEdit(place.value, value, key);
}

/** What `applyComposition` asks of the editor. */
export interface CompositionHost {
  textInput(kind: 'text', data: string): void;
  replaceText?(
    target: ReplacementTarget | null,
    replacement: string,
    place?: ReplacementPlace
  ): boolean | void;
  refuse(reason: string, detail?: string): void;
}

/**
 * Apply a composition's change at `where` (`placement`); whether it went in. When the place's
 * caret no longer holds the change, added text goes in at the start of the place's
 * paragraph, and a change to existing text is refused; with no paragraph, the text is refused.
 * A refusal publishes its reason.
 */
export function applyComposition(
  host: CompositionHost,
  where: ReplacementPlace | null,
  edit: ContextEdit
): boolean {
  if (edit.target.text === '' && edit.text === '') return true; // Nothing changed.
  if (!where || !host.replaceText) {
    host.refuse('errors.inputUnsupported', 'insertCompositionText');
    return false;
  }
  // Plain typing at the caret: the editor's own text input (AutoCorrect, list starts).
  if (!where.at && edit.target.text === '' && edit.target.after === 0) {
    host.textInput('text', edit.text);
    return true;
  }
  if (host.replaceText(edit.target, edit.text, where) !== false) return true;
  const at = where.at;
  if (!at || at.kind !== 'text' || edit.target.text !== '' || !edit.text) return false;
  const start = { paragraph: at.focus.paragraph, offset: 0 };
  const target = { text: '', after: 0, between: '', paragraph: edit.target.paragraph };
  return host.replaceText(target, edit.text, { at: atPoint(at, start) }) !== false;
}

const atPoint = (
  sel: EditorSelection & { kind: 'text' },
  point: { paragraph: number; offset: number }
): EditorSelection => ({ ...sel, anchor: point, focus: point });
