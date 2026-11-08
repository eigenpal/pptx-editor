// Adapted from docx-editor/packages/core/src/editor/surface-input.ts at f189d6503 (EigenPal).
// The model side of `applyReplacementText`: a system text change (autocorrect, a capital,
// a curly quote, dictation, the double-space period of macOS) arrives as
// `insertReplacementText`, and the text it targets is matched against the model text before
// the caret. On agreement the replacement replaces those characters; on anything else it
// goes in at the caret, which never edits text the system did not target. A target that
// names another paragraph is never matched.
//
// PPTX Editor types into a hidden textarea, not into painted pages, so the target comes from
// the textarea's selection over the text that the surface mirrors into it (`TypingContext`):
// the caret's paragraph up to the caret. A target can end before the caret (`after`), as when
// the system corrects a word after the space that ends it. The paragraph check compares the
// mirror's paragraph key with the caret's. DOM-free: `surface-input.ts` reads the event.

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
  /** The paragraph key of the target, or null when the event does not say. */
  readonly paragraph: string | null;
}

/** The most characters of context the hidden input holds. */
export const CONTEXT_LIMIT = 200;

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
 * at most `CONTEXT_LIMIT` characters, cut at a space where it is long.
 */
export function contextBefore(paragraphText: string, offset: number): string {
  const before = paragraphText.slice(0, offset);
  let from = Math.max(0, before.length - CONTEXT_LIMIT);
  // A line break reads as `\n` in the model text (`paragraphText`).
  const lineStart = before.lastIndexOf('\n') + 1;
  if (lineStart > from) from = lineStart;
  else if (from > 0) {
    const space = before.indexOf(' ', from);
    if (space >= 0) from = space + 1;
  }
  return before.slice(from);
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
  /** Whether a caret format waits for the next typed text. */
  armed(): boolean;
  /** The editor's text input; `exact`: no AutoCorrect. */
  textInput(kind: 'text', data: string, exact: boolean): void;
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
     * Apply a system text change. The replacement replaces its target when the model text
     * before the caret holds it, in the caret's paragraph; else it goes in at the caret. An
     * armed caret format keeps the caret path, so the format rides the text as it rides a
     * keystroke. The edit is a text input over the target: one store transaction, and the
     * undo step of text typed over a range.
     */
    replaceText(target: ReplacementTarget | null, replacement: string): void {
      if (!replacement) return;
      const sel = host.selection();
      const at = target && !host.armed() ? caretAt(sel) : null;
      const range = at && target && replacementRange(at.text, at.offset, at.key, target);
      if (sel.kind !== 'text' || !target || !range)
        return host.textInput('text', replacement, true);
      const point = (offset: number) => ({ paragraph: sel.focus.paragraph, offset });
      host.select({ ...sel, anchor: point(range.from), focus: point(range.to) });
      const before = host.selection();
      host.textInput('text', replacement, true);
      const now = host.selection();
      // A refused input leaves the target selected: the caret goes back.
      if (now === before) return host.select(sel);
      // The text between the target and the caret stays: the caret goes back after it.
      if (target.after > 0 && now.kind === 'text') {
        const caret = point(now.focus.offset + target.after);
        host.select({ ...now, anchor: caret, focus: caret });
      }
    },
  };
}

/** One change to the context: `target` becomes `text`. */
export interface ContextEdit {
  readonly target: ReplacementTarget;
  readonly text: string;
}

/**
 * What an IME composition changed in the mirrored context, from the field's value before the
 * composition (`before`), after it (`after`), the caret after it, and the committed text
 * (`data`). A keyboard can recompose a word that is already in the field, so the committed
 * text can cover characters of `before`. The change keeps the characters that both values
 * share, and at least one committed character. Null when the field does not hold `data` at
 * the caret: then `data` goes in at the caret.
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
  // At least one committed character stays, so the change is never an empty replacement.
  let lead = 0;
  while (lead < replaced.length && lead < text.length - 1 && replaced[lead] === text[lead])
    lead += 1;
  replaced = replaced.slice(lead);
  text = text.slice(lead);
  let trail = 0;
  while (
    trail < replaced.length &&
    trail < text.length - 1 &&
    replaced[replaced.length - 1 - trail] === text[text.length - 1 - trail]
  )
    trail += 1;
  replaced = replaced.slice(0, replaced.length - trail);
  text = text.slice(0, text.length - trail);
  return { target: { text: replaced, after: tail + trail, paragraph: key }, text };
}
