// Adapted from docx-editor/packages/core/src/editor/surface-input.ts at f189d6503 (EigenPal).
// The model side of `applyReplacementText`: a system text change (autocorrect, a capital,
// a curly quote, dictation, the double-space period of macOS) arrives as
// `insertReplacementText`, and the text it targets is matched against the model text before
// the caret. On agreement the replacement replaces those characters. A target that the model
// text does not hold, or that names another paragraph, is never matched: the input is
// refused, so no text the system did not target changes and no word goes in twice. A
// replacement with no target goes in at the caret.
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

/** The text the hidden input holds for the system: the caret paragraph before the caret. */
export interface TypingContext {
  /** The paragraph the text comes from (`paragraphKey`). */
  readonly key: string;
  readonly text: string;
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
): { readonly key: string; readonly text: string; readonly offset: number } | null {
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
   * The editor's text input, as a system replacement: no AutoCorrect, and an undo gives
   * `before` back as the selection.
   */
  textInput(
    kind: 'text' | 'deleteBackward',
    data: string | undefined,
    before: EditorSelection
  ): void;
}

/** The editor's side of system text changes: the input's context and the replacement. */
export function createSystemText(host: SystemTextHost) {
  const caretAt = (sel: EditorSelection) => {
    const session = host.session();
    const part = host.part();
    return session && part && !host.viewing() ? caretParagraph(session, part, sel) : null;
  };
  return {
    /** The text the surface's input holds for the system. */
    typingContext(): TypingContext | null {
      const at = caretAt(host.selection());
      return at ? { key: at.key, text: contextBefore(at.text, at.offset) } : null;
    },
    /**
     * Apply a system text change. With no target, the replacement goes in at the caret. With
     * one, it replaces the target when the model text before the caret holds it, in the
     * caret's paragraph; else the input is refused, so no word goes in twice. An empty
     * replacement deletes the target. A caret format that waits at the caret stays for the
     * next typed text. The edit is one text input over the target: one store transaction,
     * one undo step, and an undo puts the caret back where it was.
     */
    replaceText(target: ReplacementTarget | null, replacement: string, committed?: string): void {
      const sel = host.selection();
      if (!target) {
        if (replacement) host.textInput('text', replacement, sel);
        return;
      }
      const at = caretAt(sel);
      const range = at && replacementRange(at.text, at.offset, at.key, target);
      if (sel.kind !== 'text' || !at || !range) {
        // An IME's text is the user's typing: when a peer's edit moved its target, the
        // characters the model does not hold before the caret go in there, as typing would.
        const fresh = committed && at ? newCharacters(at.text.slice(0, at.offset), committed) : '';
        if (fresh) return host.textInput('text', fresh, sel);
        return host.refuse('errors.inputUnsupported', 'insertReplacementText');
      }
      // A change that changes nothing writes nothing: no undo step, no collaboration op.
      if (at.text.slice(range.from, range.to) === replacement) return;
      const point = (offset: number) => ({ paragraph: sel.focus.paragraph, offset });
      host.follow(() => {
        host.select({ ...sel, anchor: point(range.from), focus: point(range.to) });
        const over = host.selection();
        if (replacement) host.textInput('text', replacement, sel);
        else host.textInput('deleteBackward', undefined, sel);
        const now = host.selection();
        // A refused input leaves the target selected: the caret goes back.
        if (now === over) return host.select(sel);
        // The text between the target and the caret stays: the caret goes back after it.
        if (target.after > 0 && now.kind === 'text') {
          const caret = point(now.focus.offset + target.after);
          host.select({ ...now, anchor: caret, focus: caret });
        }
      });
      host.emit();
    },
  };
}

/**
 * The end of `committed` that `before` (the model text before the caret) does not already
 * end with: the characters to type so that the text before the caret ends with `committed`.
 * A word the model holds is never typed twice.
 */
export function newCharacters(before: string, committed: string): string {
  for (let k = Math.min(committed.length, before.length); k > 0; k -= 1)
    if (before.endsWith(committed.slice(0, k))) return committed.slice(k);
  return committed;
}

/** One change to the context: `target` becomes `text`. */
export interface ContextEdit {
  readonly target: ReplacementTarget;
  readonly text: string;
}

/**
 * The change from the context `before` to the field value `after`, by the characters both
 * share at the start and the end. `before` ends at the caret, so the shared end is the
 * target's distance to the caret.
 */
export function fieldEdit(before: string, after: string, key: string): ContextEdit {
  let lead = 0;
  const most = Math.min(before.length, after.length);
  while (lead < most && before[lead] === after[lead]) lead += 1;
  let trail = 0;
  while (
    trail < most - lead &&
    before[before.length - 1 - trail] === after[after.length - 1 - trail]
  )
    trail += 1;
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
  let replaced = before.slice(at, before.length - tail);
  let text = data;
  let lead = 0;
  while (lead < replaced.length && lead < text.length && replaced[lead] === text[lead]) lead += 1;
  replaced = replaced.slice(lead);
  text = text.slice(lead);
  let trail = 0;
  while (
    trail < replaced.length &&
    trail < text.length &&
    replaced[replaced.length - 1 - trail] === text[text.length - 1 - trail]
  )
    trail += 1;
  replaced = replaced.slice(0, replaced.length - trail);
  text = text.slice(0, text.length - trail);
  const reach = tail + trail;
  const between = before.slice(before.length - reach);
  return { target: { text: replaced, after: reach, between, paragraph: key }, text };
}
