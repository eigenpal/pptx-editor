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
      /** A composition that the editor ended, which this one may continue (`continues`). */
      readonly resumes: EndedComposition | null;
    };

export const IDLE: CompositionState = { phase: 'idle' };

/** The state when a composition starts, from the input as the system left it. */
export function startComposition(
  context: TypingContext | null,
  value: string,
  from: number,
  to: number,
  width: number,
  resumes: EndedComposition | null = null
): CompositionState {
  const place = { context, value, from, to };
  return { phase: 'composing', place, seen: context, width, resumes };
}

/**
 * What a new context during a composition means. `same`: the caret is still at the place.
 * `follow`: an edit (a peer's, or the API's) moved the caret and kept it in the place's
 * paragraph, by its element id: a peer's Enter above, a row inserted above the cell, or text
 * deleted before the caret. The place moves with the caret. `moved`: an edit took the caret
 * out of the place's paragraph; the place stays where its paragraph is, if that paragraph
 * still holds it (`followPlace`). `left`: the caret moved with no edit (a click, a tap, a
 * key, the API), so the composition must end.
 */
export function contextChange(
  state: CompositionState,
  now: TypingContext | null
): 'same' | 'follow' | 'moved' | 'left' {
  if (state.phase === 'idle') return 'same';
  const { seen } = state;
  if (now?.key === seen?.key && now?.text === seen?.text) return 'same';
  if (!now || !seen || now.revision === seen.revision) return 'left';
  const same = now.paragraphId ?? now.key;
  return same === (seen.paragraphId ?? seen.key) ? 'follow' : 'moved';
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
 * the caret). At the place's own caret (`at`) when the caret left it. Either way the place
 * must still hold its paragraph and the text before it (`expect`), or the write is refused.
 * Null: no text caret.
 */
export function placement(
  state: CompositionState,
  now: TypingContext | null
): ReplacementPlace | null {
  if (state.phase === 'idle' || !state.seen) return null;
  const { seen } = state;
  const expect = expectAt(seen);
  if (now?.key === seen.key && now.text === seen.text) return { typed: true, expect };
  return seen.selection ? { at: seen.selection, expect } : null;
}

/** What a place must still hold: its paragraph and the text before its caret. */
function expectAt(seen: TypingContext): ReplacementPlace['expect'] {
  return seen.paragraphId === undefined
    ? undefined
    : { paragraph: seen.paragraphId, text: seen.text };
}

/**
 * A composition that the editor ended (`surface-input.ts`), and what it committed: the
 * caret after its text, and its text before that caret (`target`). Input with no key first
 * that continues that text (`continues`) is the input method's rest of it, which it sent
 * before it learned of the end, and goes there.
 */
export interface EndedComposition {
  readonly at: Extract<EditorSelection, { kind: 'text' }>;
  readonly expect: NonNullable<ReplacementPlace['expect']>;
  readonly target: ReplacementTarget;
}

/**
 * The ended composition `state`, after its text `value` (the input's final value) went in
 * at its place. Null when the place moved with an edit or the input changed outside it.
 */
export function endedComposition(state: CompositionState, value: string): EndedComposition | null {
  if (state.phase === 'idle' || !state.seen) return null;
  const { place, seen } = state;
  const sel = seen.selection;
  const paragraph = seen.paragraphId;
  if (sel?.kind !== 'text' || paragraph === undefined || seen.text !== place.value) return null;
  const tail = place.value.length - place.to;
  const kept =
    value.slice(0, place.from) === place.value.slice(0, place.from) &&
    value.slice(value.length - tail) === place.value.slice(place.to);
  if (!kept) return null;
  const point = { ...sel.focus, offset: sel.focus.offset + value.length - place.value.length };
  const text = value.slice(place.from, value.length - tail);
  const between = value.slice(value.length - tail);
  return {
    at: { ...sel, anchor: point, focus: point },
    expect: { paragraph, text: value },
    target: { text, after: tail, between, paragraph: seen.key },
  };
}

/** The ended composition `late` after `text` replaced its text at its place (`continues`). */
export function resumed(late: EndedComposition, text: string): EndedComposition {
  const { at, expect, target } = late;
  const point = { ...at.focus, offset: at.focus.offset + text.length - target.text.length };
  const cut = expect.text.length - target.after - target.text.length;
  const after = expect.text.slice(0, cut) + text + (target.between ?? '');
  return {
    at: { ...at, anchor: point, focus: point },
    expect: { ...expect, text: after },
    target: { ...target, text },
  };
}

/**
 * Whether `data` continues the text an ended composition committed, `committed`: it adds to
 * it, takes from its end, or changes one character or swaps two (a correction). Content,
 * never time, decides.
 */
export function continues(committed: string, data: string): boolean {
  if (!committed || !data) return false;
  if (data.startsWith(committed) || committed.startsWith(data)) return true;
  return Math.abs(committed.length - data.length) <= 1 && editDistance(committed, data) <= 1;
}

/** The optimal string alignment distance: inserts, deletes, changes, and swaps of two. */
function editDistance(a: string, b: string): number {
  const d = Array.from({ length: a.length + 1 }, (_, i) => [i, ...Array<number>(b.length).fill(0)]);
  for (let j = 1; j <= b.length; j++) d[0]![j] = j;
  for (let i = 1; i <= a.length; i++) {
    for (let j = 1; j <= b.length; j++) {
      const cost = a[i - 1] === b[j - 1] ? 0 : 1;
      let best = Math.min(d[i - 1]![j]! + 1, d[i]![j - 1]! + 1, d[i - 1]![j - 1]! + cost);
      if (i > 1 && j > 1 && a[i - 1] === b[j - 2] && a[i - 2] === b[j - 1])
        best = Math.min(best, d[i - 2]![j - 2]! + 1);
      d[i]![j] = best;
    }
  }
  return d[a.length]![b.length]!;
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
  const { at, expect } = where;
  if (!at || at.kind !== 'text' || !expect || edit.target.text !== '' || !edit.text) return false;
  // Only in the place's own paragraph, found by its element id.
  const start = { paragraph: at.focus.paragraph, offset: 0 };
  const target = { text: '', after: 0, between: '', paragraph: null };
  const there = { at: atPoint(at, start), expect: { paragraph: expect.paragraph, text: '' } };
  return host.replaceText(target, edit.text, there) !== false;
}

const atPoint = (
  sel: EditorSelection & { kind: 'text' },
  point: { paragraph: number; offset: number }
): EditorSelection => ({ ...sel, anchor: point, focus: point });
