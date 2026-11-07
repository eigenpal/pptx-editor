// Where an edit happened and what was selected, kept with each history entry as its
// `selectionBefore` and `selectionAfter`. Node ids are per part, so a selection means something
// only on its own slide: undo and redo go back to the slide the edit changed and restore the
// selection there, or select nothing when its shapes are gone.

import type { SlideLayout } from '../layout/index.ts';
import type { TextEditAt } from '../store/index.ts';
import type { EditorCommand, EditorSelection } from '../contracts/index.ts';
import { findItem } from '../automation/commands/placement.ts';
import { cellLayoutAt, snapCellAddress, tableItem, textItemOf } from './table-cells.ts';
import { snapTextSelection } from './text-input.ts';

/** The slide part of an edit and the selection on it. */
export interface HistoryMark {
  readonly part: string | undefined;
  readonly selection: EditorSelection;
}

/** The slide a command changes: the slide it names, or the current one. */
export function changedSlide(command: EditorCommand, current: number): number {
  if (
    command.type === 'setNotes' ||
    command.type === 'deleteSlide' ||
    command.type === 'duplicateSlide'
  )
    return command.index;
  return command.type === 'moveSlide' ? command.from : current;
}

/** Where a text edit in a shape starts: the start of the text selection of `mark`. */
export function textEditAt(mark: HistoryMark): TextEditAt | undefined {
  const sel = mark.selection;
  if (sel.kind !== 'text' || sel.cell || mark.part === undefined) return undefined;
  const { anchor: a, focus: f } = sel;
  const start =
    a.paragraph < f.paragraph || (a.paragraph === f.paragraph && a.offset <= f.offset) ? a : f;
  return { part: mark.part, node: sel.nodeId, paragraph: start.paragraph, offset: start.offset };
}

/**
 * The transaction options of an edit with history marks `before` and `after`: the marks, and
 * where a text edit starts.
 */
export function markOptions(
  before: HistoryMark,
  after?: HistoryMark
): { selectionBefore: HistoryMark; selectionAfter?: HistoryMark; textAt?: TextEditAt } {
  const at = textEditAt(before);
  return {
    selectionBefore: before,
    ...(after ? { selectionAfter: after } : {}),
    ...(at ? { textAt: at } : {}),
  };
}

/** The history mark for an edit of `part`: `selection` when `part` is the shown slide `shown`. */
export function markOf(
  part: string | undefined,
  shown: string | undefined,
  selection: EditorSelection
): HistoryMark {
  return { part, selection: part === shown ? selection : { kind: 'none' } };
}

/** A history entry's selection state, as the facade stores it. */
export function isHistoryMark(v: unknown): v is HistoryMark {
  return typeof v === 'object' && v !== null && 'selection' in v && 'part' in v;
}

/** The selection if every element it names is on `layout`, else none. */
export function validSelection(
  layout: SlideLayout | undefined,
  sel: EditorSelection
): EditorSelection {
  const none: EditorSelection = { kind: 'none' };
  if (!layout || sel.kind === 'none') return none;
  if (sel.kind === 'shapes') {
    const ids = sel.group ? [...sel.nodeIds, sel.group] : sel.nodeIds;
    return ids.length && ids.every((id) => findItem(layout.items, id)) ? sel : none;
  }
  if (sel.kind === 'cells') {
    const t = tableItem(layout, sel.nodeId)?.table;
    return t && cellLayoutAt(t, sel.anchor) && cellLayoutAt(t, sel.focus) ? sel : none;
  }
  const at = snapCellAddress(layout, sel);
  const item = textItemOf(layout, at);
  if (at.kind !== 'text' || !item?.element.text) return none;
  return snapTextSelection(item, at);
}
