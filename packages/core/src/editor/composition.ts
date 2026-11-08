// IME composition in the slide surface: the state of one composition, and where its text goes.
//
// Four modules hold the rules:
// - `surface-input.ts`: the browser events. It starts, keeps, and ends compositions.
// - `composition.ts` (this module): the state and its pure transitions.
// - `replacement-text.ts`: the editor side. It checks a place and writes there.
// - `paragraph-place.ts`: what a place is, and the check that it still holds.
//
// The invariant: text is written only at a place that the check proves. A composition at a
// caret has a `TextPlace`: the caret, its paragraph's element id, and the paragraph's exact
// text before the caret. A composition over a selected range has a `SelectionPlace`: the
// range, its paragraphs' element ids, and its exact text. Every write goes through one checked
// write (`compose` or `composeSelection`), which refuses with `errors.compositionChanged`
// when the place no longer holds. A write never goes to the paragraph start, to another
// paragraph, or to the caret instead.
//
// A composition ends when:
// - the browser ends it (`compositionend`): its text goes in at its place;
// - the caret or the selection moves locally (a click, a tap, a key, the API), or the editor
//   is about to change what it applies to (an undo, a slide change): the surface makes the
//   browser end it, and its text goes in at its place;
// - a local or API edit changes the model: the same, and the caret stays where the edit
//   left it;
// - a peer's transaction leaves the place unproven (an edit before or over it, a split before
//   it, a merge, a delete): the composition is lost, and its text is refused.
// A peer's transaction that keeps the place (an edit after it, or in another paragraph) keeps
// the composition.
//
//   idle --compositionstart--> composing --compositionend--> idle (write at the place)
//   composing --peer edit, place holds--> composing
//   composing --peer edit, place gone--> lost --end--> idle (refused)
//   composing --local caret move or edit, undo, slide change--> ending --end--> idle (write)
//
// Known limit: with no collaboration session there is no caret that moves with its
// characters. A store edit there (a host's transaction) that puts the same text before the
// place as was there passes the check, and the text goes in after that text. For example, a
// host that turns "Hello|" into "Hello Hello" gets the composition after the first "Hello".
// Text alone cannot tell such an edit from one after the place.
//
// The transitions are pure; `applyComposition` writes through the editor.

import {
  compositionEdit,
  fieldEdit,
  type ContextEdit,
  type ReplacementTarget,
  type SelectionPlace,
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
  /** For a composition with no text caret (a selected range, a shape): the selection then. */
  readonly selection?: SelectionPlace;
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
      /** A peer's edit took the place away: the composition's text is refused. */
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
  selection?: SelectionPlace
): CompositionState {
  const place = { context, value, from, to, ...(selection ? { selection } : {}) };
  return { phase: 'composing', place, seen: context, width };
}

/**
 * What a new context during a composition at a caret means. `remote`: a peer's transaction
 * changed the model, so the place must be checked again (`keepPlace`). `same`: the caret is
 * still at the place. `left`: the caret moved, or a local or API edit changed the model, so
 * the composition must end. No context, before or now, is `left` too.
 */
export function contextChange(
  state: CompositionState,
  now: TypingContext | null,
  remote: boolean
): 'same' | 'remote' | 'left' {
  if (state.phase === 'idle') return 'same';
  const { seen } = state;
  if (!now || !seen) return 'left';
  if (now.revision !== seen.revision) return remote ? 'remote' : 'left';
  return now.key === seen.key && now.text === seen.text ? 'same' : 'left';
}

/** The state after a peer's edit kept the place (`contextChange` gave `remote`). */
export function placeKept(state: CompositionState, now: TypingContext): CompositionState {
  return state.phase === 'idle' ? state : { ...state, seen: now };
}

/** The state while the surface waits for the browser's final text. */
export function ending(state: CompositionState): CompositionState {
  return state.phase === 'composing' ? { ...state, phase: 'ending' } : state;
}

/** The state after a peer's edit took the place away (`keepPlace`): its text is refused. */
export function placeLost(state: CompositionState): CompositionState {
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
