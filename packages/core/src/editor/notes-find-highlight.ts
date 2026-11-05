// The find match in the speaker notes field. Find keeps the focus in its own field, so the match
// is not a DOM selection: it is painted with a CSS custom highlight (`::highlight(pptx-find)`),
// which needs no change to the field's DOM and leaves the caret and the focus where they are.
// Framework-free, so the React and Vue notes panes share it.

import type { FindMatch, NotesParagraph } from '../contracts/index.ts';
import { findNotesMatch } from './find-session.ts';
import { notesOffsetOf, notesPointAt } from './notes-dom.ts';
import { notesFieldFocused } from './pane-support.ts';
import type { PptxEditorInstance } from './pptx-editor-types.ts';

/** The highlight name the editor stylesheet paints (`::highlight(pptx-find)`). */
export const NOTES_FIND_HIGHLIGHT = 'pptx-find';

/** The match range of each painted notes field. */
const painted = new Map<HTMLElement, Range>();

interface HighlightRegistry {
  set(name: string, value: unknown): void;
  delete(name: string): void;
}
type HighlightConstructor = new () => { add(range: Range): void };

/** Paint every field's range, where the browser has CSS custom highlights. */
function publish(): void {
  const css = (globalThis as { CSS?: { highlights?: HighlightRegistry } }).CSS;
  const Highlight = (globalThis as { Highlight?: HighlightConstructor }).Highlight;
  if (!css?.highlights || !Highlight) return;
  if (!painted.size) {
    css.highlights.delete(NOTES_FIND_HIGHLIGHT);
    return;
  }
  const highlight = new Highlight();
  for (const range of painted.values()) highlight.add(range);
  css.highlights.set(NOTES_FIND_HIGHLIGHT, highlight);
}

/**
 * The plain-text range of `match` in the notes field that shows `paragraphs`: paragraphs are
 * `\n` between them, as `notesFieldText` reads them. Null when the match is past the text.
 * @internal
 */
export function notesMatchRange(
  paragraphs: readonly NotesParagraph[],
  match: Pick<FindMatch, 'paragraph' | 'start' | 'end'>
): { start: number; end: number } | null {
  if (match.paragraph >= paragraphs.length) return null;
  let at = 0;
  for (let i = 0; i < match.paragraph; i += 1)
    at += paragraphs[i]!.runs.reduce((n, r) => n + r.text.length, 0) + 1;
  const length = paragraphs[match.paragraph]!.runs.reduce((n, r) => n + r.text.length, 0);
  if (match.end > length) return null;
  return { start: at + match.start, end: at + match.end };
}

/**
 * Paint the speaker notes match that `editor`'s find shows in the notes field `root`, which
 * shows `paragraphs`, and scroll the field to it. Clears the paint when find shows no notes
 * match, while the field holds a draft, and while the field has the focus (the user's own
 * selection is then the one to see). Returns whether a match is painted. @internal
 */
export function paintNotesFindMatch(
  editor: PptxEditorInstance | null,
  root: HTMLElement | null,
  paragraphs: readonly NotesParagraph[],
  drafting = false
): boolean {
  const match = editor ? findNotesMatch(editor.find) : null;
  const range =
    root && match && !drafting && !notesFieldFocused(root)
      ? notesMatchRange(paragraphs, match)
      : null;
  if (!root || !range) {
    if (root && painted.delete(root)) publish();
    return false;
  }
  const a = notesPointAt(root, range.start);
  const b = notesPointAt(root, range.end);
  const r = root.ownerDocument.createRange();
  r.setStart(a.node, a.offset);
  r.setEnd(b.node, b.offset);
  painted.set(root, r);
  publish();
  // Scroll the field, not the page, so the match shows.
  const box = typeof r.getBoundingClientRect === 'function' ? r.getBoundingClientRect() : null;
  const field = root.getBoundingClientRect();
  if (box && box.height > 0) {
    if (box.top < field.top) root.scrollTop -= field.top - box.top;
    else if (box.bottom > field.bottom) root.scrollTop += box.bottom - field.bottom;
  }
  return true;
}

/** Stop painting a notes field that unmounts. @internal */
export function clearNotesFindMatch(root: HTMLElement | null): void {
  if (root && painted.delete(root)) publish();
}

/**
 * A painter for one notes pane: it paints the field it is given and stops painting a field that
 * the pane no longer shows (the pane closed, or the field was drawn again). @internal
 */
export function createNotesFindPainter(): {
  paint(
    editor: PptxEditorInstance | null,
    root: HTMLElement | null,
    paragraphs: readonly NotesParagraph[],
    drafting?: boolean
  ): boolean;
  stop(): void;
} {
  let last: HTMLElement | null = null;
  return {
    paint(editor, root, paragraphs, drafting = false) {
      if (last && last !== root) clearNotesFindMatch(last);
      last = root;
      return paintNotesFindMatch(editor, root, paragraphs, drafting);
    },
    stop() {
      clearNotesFindMatch(last);
      last = null;
    },
  };
}

/** The plain-text range of the find match painted in the notes field `root`, or null. @internal */
export function paintedNotesFindRange(root: HTMLElement): { start: number; end: number } | null {
  const r = painted.get(root);
  return r
    ? {
        start: notesOffsetOf(root, r.startContainer, r.startOffset),
        end: notesOffsetOf(root, r.endContainer, r.endOffset),
      }
    : null;
}
