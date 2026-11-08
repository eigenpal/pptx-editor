// Adapted from docx-editor/packages/core/src/editor/surface-input.ts at f189d6503 (EigenPal).
// The model side of `applyReplacementText`: a system text change (autocorrect, a capital,
// a curly quote, dictation, the double-space period of macOS) arrives as
// `insertReplacementText`, and the text it targets is matched against the model text before
// the caret. On agreement the replacement replaces those characters. A target that the model
// text does not hold, or that names another paragraph, is never matched: the input is
// refused, so no text the system did not target changes and no word goes in twice. A
// replacement with no target goes in at the caret, unless it is one word right after a
// word (`correctsWord`): such a replacement can correct that word, so it is refused.
//
// PPTX Editor types into a hidden textarea, not into painted pages, so the target comes from
// the textarea's selection over the text that the surface mirrors into it (`TypingContext`):
// the caret's paragraph up to the caret. A target can end before the caret (`after`), as when
// the system corrects a word after the space that ends it. The paragraph check compares the
// mirror's paragraph key with the caret's. DOM-free: `surface-input.ts` reads the event.

import { segmentGraphemes } from '@eigenpal/ooxml-core';
import type { EditorSelection } from '../contracts/index.ts';
import { paragraphText, paragraphsOf, slideNumberOf } from '../automation/commands/text-model.ts';
import type { DocumentSession } from './document-session.ts';
import { selectionBody } from './table-cells.ts';
import { deliveringChange } from '../store/store/delivering.ts';
import {
  holdsSelection,
  placeCaret,
  selectionPlace,
  type SelectionPlace,
  type TextPlace,
} from './paragraph-place.ts';

export type { SelectionPlace, TextPlace } from './paragraph-place.ts';

/** The text the hidden input holds for the system: the caret paragraph before the caret. */
export interface TypingContext {
  /** The paragraph the text comes from (`paragraphKey`). */
  readonly key: string;
  readonly text: string;
  /** The caret the text ends at. */
  readonly selection?: EditorSelection;
  /** The store revision the text was read at. */
  readonly revision?: number;
  /** The element id of the caret's paragraph, which edits do not change (`paragraph-place.ts`). */
  readonly paragraphId?: string;
  /** The paragraph's exact text before the caret: the place of a composition (`TextPlace`). */
  readonly before?: string;
}

/** What one system replacement targeted. */
export interface ReplacementTarget {
  /** The replaced text. */
  readonly text: string;
  /** How many characters lie between the end of the target and the caret. */
  readonly after: number;
  /** The text between the end of the target and the caret, when the event says. */
  readonly between?: string;
  /** The paragraph key of the target, or null when the event does not say. */
  readonly paragraph: string | null;
}

/** The most characters of context the hidden input holds. */
export const CONTEXT_LIMIT = 216;
/** How far past the cut the context looks for a word start: it keeps at least 200. */
const WORD_SLACK = 16;

/** The key of the paragraph a caret is in: shape, table cell, and paragraph index. */
export function paragraphKey(
  part: string,
  nodeId: string,
  cell: { readonly row: number; readonly column: number } | undefined,
  paragraph: number
): string {
  return JSON.stringify([part, nodeId, cell ? [cell.row, cell.column] : null, paragraph]);
}

/**
 * The context for the system: the paragraph text before the caret, from its last line break,
 * at most `CONTEXT_LIMIT` characters. A long text is cut at a word start near the limit, else
 * at a grapheme boundary, so the context keeps at least `CONTEXT_LIMIT - WORD_SLACK`.
 */
export function contextBefore(paragraphText: string, offset: number): string {
  const before = paragraphText.slice(0, offset);
  // A line break reads as `\n` in the model text (`paragraphText`).
  const lineStart = before.lastIndexOf('\n') + 1;
  let from = Math.max(lineStart, before.length - CONTEXT_LIMIT);
  if (from > lineStart) {
    const space = before.slice(from, from + WORD_SLACK).search(/\s/u);
    from = space >= 0 ? from + space + 1 : graphemeStart(before, from);
  }
  return before.slice(from);
}

/** The first grapheme boundary at or after `at` in `text`. */
function graphemeStart(text: string, at: number): number {
  const base = Math.max(0, at - 32);
  const boundary = segmentGraphemes(text.slice(base, at + 32))
    .map((g) => base + g.utf16From)
    .find((offset) => offset >= at);
  return boundary ?? at;
}

/**
 * The model text of the paragraph a collapsed text caret is in, and its key; null for a
 * range, a diagram node, or a caret with no paragraph. A read: it commits nothing.
 */
export function caretParagraph(
  session: DocumentSession,
  part: string,
  sel: EditorSelection
): {
  readonly key: string;
  readonly text: string;
  readonly offset: number;
  readonly id: string;
} | null {
  if (sel.kind !== 'text' || sel.diagramNode !== undefined) return null;
  const { anchor, focus } = sel;
  if (anchor.paragraph !== focus.paragraph || anchor.offset !== focus.offset) return null;
  const batch = session.probeBatch();
  const body = selectionBody(batch, part, sel);
  const p = body && paragraphsOf(body)[focus.paragraph];
  if (!p) return null;
  return {
    key: paragraphKey(part, sel.nodeId, sel.cell, focus.paragraph),
    text: paragraphText(p, slideNumberOf(batch, part)),
    offset: focus.offset,
    id: p.id,
  };
}

/**
 * The range `[from, to)` of the caret's paragraph that a target replaces, or null when the
 * target does not match the model text before the caret. `text` is the caret's paragraph,
 * `caret` its offset, and `key` its paragraph key.
 */
export function replacementRange(
  text: string,
  caret: number,
  key: string,
  target: ReplacementTarget
): { readonly from: number; readonly to: number } | null {
  if (target.paragraph !== null && target.paragraph !== key) return null;
  if (target.after < 0) return null;
  const to = caret - target.after;
  const from = to - target.text.length;
  if (from < 0) return null;
  // The text up to the caret matches too, so an empty target (an insertion) is placed only
  // where the model still holds what the event saw.
  if (target.between !== undefined && text.slice(to, caret) !== target.between) return null;
  return text.slice(from, to) === target.text ? { from, to } : null;
}

/**
 * Whether a replacement with no target could correct the word before the caret: the
 * replacement is one word with a letter in it, and the context ends with a word, or with a
 * word and one space. Text with white space in it (dictation of a phrase), a number, an
 * emoji, or a sign is no such correction, and goes in at the caret.
 *
 * The rule also refuses a real one-word insert with no target right after a word (a
 * dictated "five"). That is the safe side: a refused word is visible and the user types it
 * again, but a correction that goes in at the caret repeats the word it corrects.
 */
export function correctsWord(context: string, replacement: string): boolean {
  return (
    /^[\p{L}\p{N}][\p{L}\p{N}\p{M}'’-]*$/u.test(replacement) &&
    /\p{L}/u.test(replacement) &&
    /[\p{L}\p{N}\p{M}] ?$/u.test(context)
  );
}

/**
 * `sel` after an edit at the caret `at` that changed the length of its paragraph by `by`
 * characters at offset `from`: the points after `from` in that paragraph move with the text.
 */
function shifted(
  sel: EditorSelection,
  at: EditorSelection,
  from: number,
  by: number
): EditorSelection {
  if (!by || sel.kind !== 'text' || at.kind !== 'text') return sel;
  if (sel.nodeId !== at.nodeId || sel.diagramNode !== undefined) return sel;
  if (JSON.stringify(sel.cell ?? null) !== JSON.stringify(at.cell ?? null)) return sel;
  const paragraph = at.focus.paragraph;
  const move = <P extends { paragraph: number; offset: number }>(p: P): P =>
    p.paragraph === paragraph && p.offset >= from ? { ...p, offset: p.offset + by } : p;
  return { ...sel, anchor: move(sel.anchor), focus: move(sel.focus) };
}

/** Whether the selection `a` has a point in the paragraph of the caret `b`. */
function sameParagraph(a: EditorSelection, b: EditorSelection & { kind: 'text' }): boolean {
  if (a.kind !== 'text' || a.nodeId !== b.nodeId || a.diagramNode !== b.diagramNode) return false;
  if (a.cell?.row !== b.cell?.row || a.cell?.column !== b.cell?.column) return false;
  const paragraph = b.focus.paragraph;
  return a.anchor.paragraph === paragraph || a.focus.paragraph === paragraph;
}

/** Whether `a` is the collapsed caret `b`: the same text body, paragraph, and offset. */
function sameCaret(a: EditorSelection, b: EditorSelection & { kind: 'text' }): boolean {
  if (!sameParagraph(a, b) || a.kind !== 'text') return false;
  const same = (p: { paragraph: number; offset: number }) =>
    p.paragraph === b.focus.paragraph && p.offset === b.focus.offset;
  return same(a.anchor) && same(a.focus);
}

/** What `createSystemText` asks of the editor. */
export interface SystemTextHost {
  session(): DocumentSession | null;
  /** The current slide's part. */
  part(): string | undefined;
  /** Whether the editor is in view mode: the input holds no text then. */
  viewing(): boolean;
  selection(): EditorSelection;
  /** Set the selection and publish it. */
  select(next: EditorSelection): void;
  /** Run `update`; a caret format that waits at the caret moves with the caret. */
  follow(update: () => void): void;
  /** Publish the editor state. */
  emit(): void;
  /** Publish why an input was not applied. */
  refuse(reason: string, detail?: string): void;
  /**
   * The editor's text input. With `before`, a system replacement: no AutoCorrect, and an undo
   * gives `before` back as the selection. With none, the user's own typing.
   */
  textInput(
    kind: 'text' | 'deleteBackward',
    data: string | undefined,
    before?: EditorSelection
  ): void;
}

/** The editor's side of system text changes: the input's context and the replacement. */
export function createSystemText(host: SystemTextHost) {
  const caretAt = (sel: EditorSelection) => {
    const session = host.session();
    const part = host.part();
    return session && part && !host.viewing() ? caretParagraph(session, part, sel) : null;
  };
  /** The text the surface's input holds for the system. */
  const typingContext = (): TypingContext | null => {
    const selection = host.selection();
    const at = caretAt(selection);
    if (!at) return null;
    const revision = host.session()?.store.revision;
    const text = contextBefore(at.text, at.offset);
    const before = at.text.slice(0, at.offset);
    return { key: at.key, text, selection, revision, paragraphId: at.id, before };
  };
  /** Run one text input; whether it applied (an applied input sets a new selection). */
  const input = (replacement: string, before: EditorSelection): boolean => {
    const was = host.selection();
    if (replacement) host.textInput('text', replacement, before);
    else host.textInput('deleteBackward', undefined, before);
    return host.selection() !== was;
  };
  /** Refuse a system change with its reason: it did not apply. */
  const refuse = (detail: string): false => {
    host.refuse('errors.inputUnsupported', detail);
    return false;
  };
  /** The caret of a composition's place, when the place still holds (`placeCaret`). */
  const caretOf = (place: TextPlace) => {
    const session = host.session();
    const part = host.part();
    return session && part && !host.viewing()
      ? placeCaret(session.probeBatch(), part, place)
      : null;
  };
  /** The place of the selection now, for a composition over it; null with no slide. */
  const placeOfSelection = (): SelectionPlace | null => {
    const session = host.session();
    const part = host.part();
    if (!session || !part || host.viewing()) return null;
    const { revision } = session.store;
    return selectionPlace(session.probeBatch(), part, host.selection(), revision);
  };
  return {
    /** The text the surface's input holds for the system. */
    typingContext,
    /** Whether the store delivers a peer's transaction now (origin `remote`). */
    remoteChange(): boolean {
      const store = host.session()?.store;
      return !!store && deliveringChange(store)?.origin === 'remote';
    },
    /** The place of the selection now, for a composition over a selected range. */
    selectionPlace: placeOfSelection,
    /**
     * Write a composition over a selected range: `replacement` replaces the selection, when it
     * is still `place` (`holdsSelection`). Else the change is refused with its reason, and goes
     * in nowhere else. Whether it applied.
     */
    composeSelection(place: SelectionPlace, replacement: string): boolean {
      const now = placeOfSelection();
      if (!now || !holdsSelection(place, now)) {
        host.refuse('errors.compositionChanged');
        return false;
      }
      const was = host.selection();
      host.textInput('text', replacement);
      return host.selection() !== was;
    },
    /**
     * The context at a composition's place after a peer's transaction, with the caret put back
     * there: the caret that the collaboration session moved is not the place. The surface
     * asks only for a remote change; a local or API change ends the composition, and its
     * caret stays where the host set it. Null when the place no longer holds. The text alone
     * can miss an edit before the place that leaves the same text there (a peer types "Hello "
     * before "Hello|"). The session moves the caret with its characters, so a caret that the
     * edit moved inside the place's paragraph also ends the place. That caret only refuses:
     * it never places text.
     */
    keepPlace(place: TextPlace): TypingContext | null {
      const at = caretOf(place);
      if (!at) return null;
      const now = host.selection();
      if (!sameCaret(now, at)) {
        if (now.kind === 'text' && sameParagraph(now, at)) return null;
        host.select(at);
      }
      return typingContext();
    },
    /**
     * Write a composition's change at its place: the only way a composition writes. The place
     * must still hold its paragraph, by element id, and the exact text before its offset
     * (`placeCaret`); else the change is refused with its reason, and goes in nowhere else.
     * Added text at the caret is the editor's own typing (AutoCorrect, list starts). Any other
     * change is one text input over its target; the caret stays where it is, and moves with
     * the text when it comes after the change. Whether it applied.
     */
    compose(place: TextPlace, target: ReplacementTarget, replacement: string): boolean {
      const at = caretOf(place);
      const there = at && caretAt(at);
      const range =
        there &&
        replacementRange(there.text, there.offset, there.key, { ...target, paragraph: null });
      if (!at || !there || !range) {
        host.refuse('errors.compositionChanged');
        return false;
      }
      // A change that changes nothing writes nothing: no undo step, no collaboration op.
      if (there.text.slice(range.from, range.to) === replacement) return true;
      const user = host.selection();
      if (range.from === range.to && range.to === there.offset && sameCaret(user, at)) {
        host.textInput('text', replacement);
        return host.selection() !== user;
      }
      const point = (offset: number) => ({ paragraph: at.focus.paragraph, offset });
      let done = false;
      host.follow(() => {
        host.select({ ...at, anchor: point(range.from), focus: point(range.to) });
        done = input(replacement, at);
        const by = done ? replacement.length - range.to + range.from : 0;
        host.select(shifted(user, at, range.to, by));
      });
      host.emit();
      return done;
    },
    /**
     * Apply a system text change. With no target, the replacement goes in at the caret, but
     * one word right after a word is refused (`correctsWord`). With one, it replaces the
     * target when the model text before the caret holds it, in the caret's paragraph; else the
     * input is refused, so no word goes in twice. An empty replacement deletes the target. A
     * caret format that waits at the caret stays for the next typed text. The edit is one
     * text input over the target: one store transaction, one undo step, and an undo puts the
     * caret back where it was. Whether it applied: a change that changes nothing applies; a
     * refusal publishes its reason.
     */
    replaceText(target: ReplacementTarget | null, replacement: string): boolean {
      const sel = host.selection();
      const at = caretAt(sel);
      if (!target) {
        // A single word right after a word can be a correction of that word whose target
        // the editor no longer has (its selection went a task before it): it is refused, so
        // no word goes in twice. Other text with no target goes in at the caret.
        if (at && correctsWord(contextBefore(at.text, at.offset), replacement))
          return refuse('insertReplacementText');
        return !replacement || input(replacement, sel);
      }
      const range = at && replacementRange(at.text, at.offset, at.key, target);
      if (sel.kind !== 'text' || !at || !range) return refuse('insertReplacementText');
      // A change that changes nothing writes nothing: no undo step, no collaboration op.
      if (at.text.slice(range.from, range.to) === replacement) return true;
      const point = (offset: number) => ({ paragraph: sel.focus.paragraph, offset });
      let done = false;
      host.follow(() => {
        host.select({ ...sel, anchor: point(range.from), focus: point(range.to) });
        done = input(replacement, sel);
        const now = host.selection();
        // A refused input leaves the target selected: the caret goes back.
        if (!done) return host.select(sel);
        // The text between the target and the caret stays: the caret goes back after it.
        if (target.after > 0 && now.kind === 'text') {
          const caret = point(now.focus.offset + target.after);
          host.select({ ...now, anchor: caret, focus: caret });
        }
      });
      host.emit();
      return done;
    },
  };
}

/** One change to the context: `target` becomes `text`. */
export interface ContextEdit {
  readonly target: ReplacementTarget;
  readonly text: string;
}

const isHigh = (c: number): boolean => c >= 0xd800 && c <= 0xdbff;
const isLow = (c: number): boolean => c >= 0xdc00 && c <= 0xdfff;

/**
 * How many UTF-16 units `a` and `b` share at the start and at the end (not overlapping),
 * by whole code points: a shared end never splits a surrogate pair, so 😀 → 😃 is one change
 * of one emoji.
 */
function sharedEnds(a: string, b: string): [number, number] {
  const most = Math.min(a.length, b.length);
  let lead = 0;
  while (lead < most && a[lead] === b[lead]) lead += 1;
  // A shared high surrogate whose low half differs belongs to the change.
  if (lead > 0 && isHigh(a.charCodeAt(lead - 1))) lead -= 1;
  let trail = 0;
  while (trail < most - lead && a[a.length - 1 - trail] === b[b.length - 1 - trail]) trail += 1;
  // A shared low surrogate whose high half differs belongs to the change.
  if (trail > 0 && isLow(a.charCodeAt(a.length - trail))) trail -= 1;
  return [lead, trail];
}

/**
 * The change from the context `before` to the field value `after`, by the characters both
 * share at the start and the end. `before` ends at the caret, so the shared end is the
 * target's distance to the caret.
 */
export function fieldEdit(before: string, after: string, key: string): ContextEdit {
  const [lead, trail] = sharedEnds(before, after);
  return {
    target: {
      text: before.slice(lead, before.length - trail),
      after: trail,
      between: before.slice(before.length - trail),
      paragraph: key,
    },
    text: after.slice(lead, after.length - trail),
  };
}

/**
 * What an IME composition changed in the mirrored context, from the field's value before the
 * composition (`before`), after it (`after`), the caret after it, and the committed text
 * (`data`). A keyboard can recompose a word that is already in the field, so the committed
 * text can cover characters of `before`, and an empty commit deletes them. The change keeps
 * the characters that both values share; it can be empty when nothing changed. Null when the
 * field does not hold `data` at the caret: then `data` goes in at the caret.
 */
export function compositionEdit(
  before: string,
  after: string,
  caret: number,
  data: string,
  key: string | null
): ContextEdit | null {
  const at = caret - data.length;
  if (at < 0 || after.slice(at, caret) !== data) return null;
  const tail = after.length - caret;
  // The field outside the composition is unchanged.
  if (tail > before.length - at || before.slice(0, at) !== after.slice(0, at)) return null;
  if (before.slice(before.length - tail) !== after.slice(caret)) return null;
  const covered = before.slice(at, before.length - tail);
  const [lead, trail] = sharedEnds(covered, data);
  const replaced = covered.slice(lead, covered.length - trail);
  const text = data.slice(lead, data.length - trail);
  const reach = tail + trail;
  const between = before.slice(before.length - reach);
  return { target: { text: replaced, after: reach, between, paragraph: key }, text };
}
