// Text editing at a text selection: typing and pasting, paragraph and line breaks, deletion by
// character, word, or line, and caret movement. Positions come from the laid-out paragraph and
// edits compile to ops on one batch; the facade commits them as one transaction.

import type { LayoutItem } from '../layout/index.ts';
import type { EditorSelection, TextPoint } from '../contracts/index.ts';
import { CommandRefusal, type OpBatch } from '../automation/commands/op-batch.ts';
import {
  comparePoints,
  deleteRange,
  insertBreak,
  insertLines,
  insertText,
  splitParagraph,
} from '../automation/commands/text-model.ts';
import {
  cleanText,
  paragraphUnits,
  snapToStop,
  stopAfter,
  stopBefore,
  nextWordStart,
  wordStartBefore,
  type ParagraphUnits,
} from '../automation/commands/text-units.ts';

/** Every kind of text input the editor applies. @public */
export type TextInputKind =
  | 'text'
  | 'deleteBackward'
  | 'deleteForward'
  | 'deleteWordBackward'
  | 'deleteWordForward'
  | 'deleteLineBackward'
  | 'deleteLineForward'
  | 'paragraph'
  | 'lineBreak';

type ShapeItem = Extract<LayoutItem, { kind: 'shape' }>;
type TextSelection = Extract<EditorSelection, { kind: 'text' }>;

const unitsOf = (item: ShapeItem, paragraph: number): ParagraphUnits =>
  paragraphUnits(item.element.text?.paragraphs[paragraph]);

/** The laid-out line a caret sits on. */
function lineAt(item: ShapeItem, at: TextPoint) {
  const lines = (item.text?.lines ?? []).filter((l) => l.paragraph === at.paragraph);
  return (
    lines.find((l, i) => at.offset >= l.start && (at.offset < l.end || i === lines.length - 1)) ??
    lines[0]
  );
}

/** The other end of a collapsed deletion, or null when there is nothing to delete. */
function deletionEnd(item: ShapeItem, kind: TextInputKind, at: TextPoint): TextPoint | null {
  const count = item.element.text?.paragraphs.length ?? 1;
  const u = unitsOf(item, at.paragraph);
  const backward = kind === 'deleteBackward' || kind === 'deleteWordBackward';
  const line =
    kind === 'deleteLineBackward' || kind === 'deleteLineForward' ? lineAt(item, at) : null;
  if (backward || kind === 'deleteLineBackward') {
    if (at.offset > 0) {
      const to =
        kind === 'deleteWordBackward'
          ? wordStartBefore(u, at.offset)
          : line && line.start < at.offset
            ? line.start
            : stopBefore(u, at.offset);
      return { paragraph: at.paragraph, offset: to };
    }
    if (at.paragraph === 0) return null;
    return { paragraph: at.paragraph - 1, offset: unitsOf(item, at.paragraph - 1).length };
  }
  if (at.offset < u.length) {
    const to =
      kind === 'deleteWordForward'
        ? nextWordStart(u, at.offset)
        : line && line.end > at.offset
          ? line.end
          : stopAfter(u, at.offset);
    return { paragraph: at.paragraph, offset: to };
  }
  if (at.paragraph >= count - 1) return null;
  return { paragraph: at.paragraph + 1, offset: 0 };
}

/**
 * Apply one text input at a text selection. A selection is replaced (or, for a deletion,
 * removed); a caret deletes one character, word, or line. Returns the caret after the edit.
 * Refuses text that holds nothing a slide can show.
 */
export function editText(
  batch: OpBatch,
  part: string,
  txBodyId: string,
  item: ShapeItem,
  sel: TextSelection,
  kind: TextInputKind,
  data?: string
): TextPoint {
  let at: TextPoint = sel.focus;
  const collapsed = comparePoints(sel.anchor, sel.focus) === 0;
  if (kind === 'text') {
    // The Tab key types a tab, which the layout moves to the next tab stop; a tab in pasted
    // text stays a space.
    const clean = data === '\t' ? data : cleanText(data ?? '').replace(/\t/g, ' ');
    if (!clean) throw new CommandRefusal('errors.invalidText');
    if (!collapsed) at = deleteRange(batch, part, txBodyId, sel.anchor, sel.focus);
    // Newlines start paragraphs; U+000B (the clipboard's line break) is a line break.
    const typeLine = (line: string): void =>
      line.split('\u000B').forEach((piece, j) => {
        if (j > 0) at = insertBreak(batch, part, txBodyId, at);
        if (piece) at = insertText(batch, part, txBodyId, at, piece);
      });
    const lines = clean.split('\n');
    typeLine(lines[0]!);
    if (lines.length === 1) return at;
    // The text after the caret moves to a paragraph of its own; the lines between the first
    // and the last go in before it as whole paragraphs, and the last line joins it.
    at = splitParagraph(batch, part, txBodyId, at);
    if (lines.length > 2) at = insertLines(batch, part, txBodyId, at, lines.slice(1, -1));
    typeLine(lines[lines.length - 1]!);
    return at;
  }
  if (!collapsed) at = deleteRange(batch, part, txBodyId, sel.anchor, sel.focus);
  if (kind === 'paragraph') return splitParagraph(batch, part, txBodyId, at);
  if (kind === 'lineBreak') return insertBreak(batch, part, txBodyId, at);
  if (!collapsed) return at;
  const end = deletionEnd(item, kind, at);
  return end ? deleteRange(batch, part, txBodyId, at, end) : at;
}

/** Collapse a horizontal selection without stepping past its edge, in either selection direction. */
export function collapseTextRange(sel: TextSelection, key: string): TextPoint | null {
  const order = comparePoints(sel.anchor, sel.focus);
  if (order === 0 || (key !== 'ArrowLeft' && key !== 'ArrowRight')) return null;
  const start = order < 0 ? sel.anchor : sel.focus;
  const end = order < 0 ? sel.focus : sel.anchor;
  return key === 'ArrowLeft' ? start : end;
}

/**
 * Move a caret or extend a range while keeping the selection's editing target. A plain Left or
 * Right collapses a range to its edge; with a modifier (a word or an edge step) the caret moves
 * on from that edge.
 */
export function moveTextSelection(
  item: ShapeItem,
  selection: TextSelection,
  key: string,
  extend: boolean,
  mods: CaretModifiers = {}
): TextSelection {
  const at = !extend && collapseTextRange(selection, key);
  const modified = !!(mods.altKey || mods.ctrlKey || mods.metaKey);
  const focus = at && !modified ? at : moveCaret(item, at || selection.focus, key, mods);
  return { ...selection, anchor: extend ? selection.anchor : focus, focus };
}

const CARET_KEYS = new Set(['ArrowLeft', 'ArrowRight', 'ArrowUp', 'ArrowDown', 'Home', 'End']);

/** True for a key `moveCaret` handles. */
export function isCaretKey(key: string): boolean {
  return CARET_KEYS.has(key);
}

/** The modifier keys of a caret key. */
export interface CaretModifiers {
  readonly altKey?: boolean;
  readonly ctrlKey?: boolean;
  readonly metaKey?: boolean;
}

/**
 * The caret after a navigation key, with the key map of DOCX Editor's `surface-input.ts`.
 * Left and Right step one character (a whole emoji, a whole field). With Alt (macOS) or Ctrl
 * (other platforms) they step to the start of the word before or the end of the word after
 * (UAX #29 words, the same units as a word deletion; forward, past the spaces after the word,
 * to the start of the next one), and at a paragraph edge into the next paragraph; both keys
 * are accepted on every platform. With Cmd alone they go to the line edges, as Home and End
 * do, for keyboards without those keys. Up and Down keep the x position on the next line; with
 * Alt (macOS) or Ctrl alone they go to the start of the paragraph (Up, or of the one before
 * when the caret is there already) or of the next paragraph (Down, or the end of the last).
 * Home and End go to the edges of the laid-out line (`lineEdge`). Home and End with Ctrl or
 * Cmd, and Up and Down with Cmd, go to the start or the end of the text body.
 */
export function moveCaret(
  item: ShapeItem,
  focus: TextPoint,
  pressed: string,
  mods: CaretModifiers = {}
): TextPoint {
  const lens = item.text?.paragraphLengths ?? [0];
  const accel = !!(mods.ctrlKey || mods.metaKey);
  const vertical = pressed === 'ArrowUp' || pressed === 'ArrowDown';
  const toEdge = vertical ? !!mods.metaKey : accel;
  const last = Math.max(0, lens.length - 1);
  if (toEdge && (pressed === 'Home' || pressed === 'ArrowUp')) return { paragraph: 0, offset: 0 };
  if (toEdge && (pressed === 'End' || pressed === 'ArrowDown'))
    return { paragraph: last, offset: lens[last] ?? 0 };
  if (vertical && (mods.altKey || mods.ctrlKey)) {
    if (pressed === 'ArrowUp')
      return {
        paragraph: Math.max(0, focus.offset > 0 ? focus.paragraph : focus.paragraph - 1),
        offset: 0,
      };
    return focus.paragraph < last
      ? { paragraph: focus.paragraph + 1, offset: 0 }
      : { paragraph: last, offset: lens[last] ?? 0 };
  }
  const sideways = pressed === 'ArrowLeft' || pressed === 'ArrowRight';
  const word = sideways && !!(mods.altKey || mods.ctrlKey);
  const edge = sideways && !word && !!mods.metaKey;
  const key = edge ? (pressed === 'ArrowLeft' ? 'Home' : 'End') : pressed;
  let { paragraph, offset } = focus;
  const u = unitsOf(item, paragraph);
  if (key === 'ArrowLeft') {
    if (offset > 0) offset = word ? wordStartBefore(u, offset) : stopBefore(u, offset);
    else if (paragraph > 0) offset = lens[--paragraph] ?? 0;
  } else if (key === 'ArrowRight') {
    if (offset < u.length) offset = word ? nextWordStart(u, offset) : stopAfter(u, offset);
    else if (paragraph < lens.length - 1) ((paragraph += 1), (offset = 0));
  } else if (key === 'Home' || key === 'End')
    offset =
      lineEdge(item, paragraph, offset, key === 'Home') ??
      (key === 'Home' ? 0 : (lens[paragraph] ?? 0));
  else if (key === 'ArrowUp' || key === 'ArrowDown') {
    const lines = item.text?.lines ?? [];
    const li = lineIndexAt(lines, paragraph, offset);
    const target = lines[li + (key === 'ArrowUp' ? -1 : 1)];
    if (target && li >= 0) {
      const x = lines[li].caretX[offset - lines[li].start] ?? 0;
      let best = 0;
      target.caretX.forEach((cx, i) => {
        if (Math.abs(cx - x) < Math.abs((target.caretX[best] ?? 0) - x)) best = i;
      });
      paragraph = target.paragraph;
      offset = snapToStop(unitsOf(item, paragraph), target.start + best);
    }
  }
  return { paragraph, offset };
}

type Lines = NonNullable<ShapeItem['text']>['lines'];

/** The laid-out line that holds the caret: a line's end offset belongs to the next line. */
function lineIndexAt(lines: Lines, paragraph: number, offset: number): number {
  return lines.findIndex(
    (l, i) =>
      l.paragraph === paragraph &&
      offset >= l.start &&
      (offset < l.end || lines[i + 1]?.paragraph !== paragraph)
  );
}

/**
 * The start or the end of the laid-out line that holds the caret, as Home and End go there.
 * The end of a line that wraps stops before the spaces it wraps at, so the caret stays at the
 * end of that line: the caret has no side, and the offset after the spaces starts the next
 * line. Undefined when the layout has no line there.
 */
function lineEdge(
  item: ShapeItem,
  paragraph: number,
  offset: number,
  start: boolean
): number | undefined {
  const lines = item.text?.lines ?? [];
  const li = lineIndexAt(lines, paragraph, offset);
  const line = lines[li];
  if (!line) return undefined;
  if (start) return line.start;
  if (lines[li + 1]?.paragraph !== paragraph) return line.end;
  const p = item.element.text?.paragraphs[paragraph];
  const text = (p?.runs ?? []).map((r) => (r.kind === 'break' ? '\n' : r.text)).join('');
  // A line break (`a:br`) ends its line: the caret goes before it.
  if (text[line.end - 1] === '\n') return line.end - 1;
  let end = line.end;
  while (end > line.start && /[^\S\n]/.test(text[end - 1] ?? '')) end -= 1;
  if (end < line.end && end > line.start) return end;
  // A wrap with no space (a long word, CJK text): the line's end, so text goes in after its
  // last character. The caret has no side, so it draws at the start of the next line there.
  return line.end;
}

/** A text selection with both ends moved out of a cluster or a field, and inside the text. */
export function snapTextSelection(item: ShapeItem, sel: TextSelection): TextSelection {
  const count = Math.max(1, item.element.text?.paragraphs.length ?? 1);
  const snap = (p: TextPoint): TextPoint => {
    const paragraph = Math.max(0, Math.min(p.paragraph, count - 1));
    return { paragraph, offset: snapToStop(unitsOf(item, paragraph), p.offset) };
  };
  const anchor = snap(sel.anchor);
  const focus = snap(sel.focus);
  return comparePoints(anchor, sel.anchor) === 0 && comparePoints(focus, sel.focus) === 0
    ? sel
    : { ...sel, anchor, focus };
}
