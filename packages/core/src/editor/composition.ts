// The state of one IME composition in the surface's input, and where its text goes in.
//
// A composition changes the model only when it ends. Its place is fixed when it starts: the
// caret's paragraph (by element id), the paragraph's exact text before the caret, and the
// part of the input's value that the system selected for it (a keyboard that composes a word
// again selects that word). Every change it makes applies there, and nowhere else
// (`applyComposition`). When the model caret leaves that place (a click or a tap, an arrow
// key, the API, a slide change), the surface makes the browser end the composition
// (`surface-input.ts`), and the browser's final text goes to the place. An edit after the
// place, or in another paragraph, keeps the place; any other edit ends the composition, and
// its text is refused (`paragraph-place.ts`).
//
// The transitions are pure; `applyComposition` writes through the editor.

import {
  compositionEdit,
  fieldEdit,
  type ContextEdit,
  type ReplacementTarget,
  type TextPlace,
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
  /** The store revision then, for a composition with no text caret (a selected range). */
  readonly revision?: number;
}

/**
 * Idle; composing at a place (`seen`: the context at the place); or ending: the surface asked
 * the browser to end the composition, and waits for its final text.
 */
export type CompositionState =
  | { readonly phase: 'idle' }
  | {
      readonly phase: 'composing' | 'ending';
      readonly place: CompositionPlace;
      readonly seen: TypingContext | null;
      /** The input's text width when the composition started, for the overlay. */
      readonly width: number;
      /** An edit before or over the place ended it: the composition's text is refused. */
      readonly lost?: true;
    };

export const IDLE: CompositionState = { phase: 'idle' };

/** The state when a composition starts, from the input as the system left it. */
export function startComposition(
  context: TypingContext | null,
  value: string,
  from: number,
  to: number,
  width: number,
  revision?: number
): CompositionState {
  const place = { context, value, from, to, ...(revision === undefined ? {} : { revision }) };
  return { phase: 'composing', place, seen: context, width };
}

/**
 * What a new context during a composition means. `follow`: the model changed (a peer's
 * edit, or the API's), so the place must be checked again (`keepPlace`); the caret that the
 * edit left is not the place. `same`: the caret is still at the place. `left`: the caret
 * moved with no edit (a click, a tap, a key, the API), so the composition must end.
 */
export function contextChange(
  state: CompositionState,
  now: TypingContext | null
): 'same' | 'follow' | 'left' {
  if (state.phase === 'idle') return 'same';
  const { seen } = state;
  if (now && seen && now.revision !== seen.revision) return 'follow';
  return now?.key === seen?.key && now?.text === seen?.text ? 'same' : 'left';
}

/** The state after an edit kept the place (`contextChange` gave `follow`). */
export function followEdit(state: CompositionState, now: TypingContext): CompositionState {
  return state.phase === 'idle' ? state : { ...state, seen: now };
}

/** The state while the surface waits for the browser's final text. */
export function ending(state: CompositionState): CompositionState {
  return state.phase === 'composing' ? { ...state, phase: 'ending' } : state;
}

/** The state after an edit took the place away (`keepPlace`): its text will be refused. */
export function lose(state: CompositionState): CompositionState {
  return state.phase === 'idle' ? state : { ...state, lost: true };
}

/**
 * The composition's place: its caret, its paragraph's element id, and the paragraph's exact
 * text before the caret. Null with no text caret.
 */
export function placement(state: CompositionState): TextPlace | null {
  if (state.phase === 'idle' || !state.seen) return null;
  const { selection: at, paragraphId: paragraph, before } = state.seen;
  return at && paragraph !== undefined && before !== undefined ? { at, paragraph, before } : null;
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
  /** Write a change at a composition's place, which must still hold; whether it applied. */
  compose?(place: TextPlace, target: ReplacementTarget, replacement: string): boolean;
  refuse(reason: string, detail?: string): void;
}

/**
 * Apply a composition's change at its place `where` (`placement`); whether it went in. The
 * editor checks the place and writes there (`compose`), or refuses the change with its
 * reason. Never anywhere else.
 */
export function applyComposition(
  host: CompositionHost,
  where: TextPlace | null,
  edit: ContextEdit
): boolean {
  if (edit.target.text === '' && edit.text === '') return true; // Nothing changed.
  if (!where || !host.compose) {
    host.refuse('errors.inputUnsupported', 'insertCompositionText');
    return false;
  }
  return host.compose(where, edit.target, edit.text);
}
