// The selection pane, framework-free: the current slide's elements as a tree, top first, with
// their names and visibility (`p:cNvPr/@name`, `@hidden`). Both adapters' `SelectionPane` draw
// these rows and send these keys and clicks here, so React and Vue behave the same. The edits
// are engine commands: `renameShape`, `setShapesVisible`, and `order`. `@internal`: adapter glue.

import type { EditorSelection, ShapeSummary } from '../contracts/index.ts';
import type { PptxEditorInstance } from './pptx-editor.ts';
import { SHAPE_NAME_MAX } from '../automation/commands/arrange-commands.ts';

/** One row of the selection pane: an element, at its depth in the group tree. @internal */
export interface SelectionPaneRow {
  readonly nodeId: string;
  readonly name: string;
  readonly kind: ShapeSummary['kind'];
  readonly hidden: boolean;
  /** 1 for the slide's top level, 2 for a group's children, and so on (`aria-level`). */
  readonly level: number;
  /** The group the element is in, or null at the top level. */
  readonly parent: string | null;
  /** A group with children: it opens and closes. */
  readonly branch: boolean;
  /** For a branch: whether its children show. */
  readonly expanded: boolean;
  /** The 1-based position among its siblings and their count (`aria-posinset`, `aria-setsize`). */
  readonly position: number;
  readonly siblings: number;
}

/**
 * The rows that show: every top-level element, top first, and the children of each group in
 * `expanded`. A group starts closed, so a slide with large groups stays short. @internal
 */
export function selectionPaneRows(
  shapes: readonly ShapeSummary[] | undefined,
  expanded: ReadonlySet<string>
): readonly SelectionPaneRow[] {
  const out: SelectionPaneRow[] = [];
  const walk = (list: readonly ShapeSummary[], level: number, parent: string | null): void => {
    list.forEach((s, i) => {
      const branch = (s.children?.length ?? 0) > 0;
      const open = branch && expanded.has(s.nodeId);
      out.push({
        nodeId: s.nodeId,
        name: s.name,
        kind: s.kind,
        hidden: s.hidden,
        level,
        parent,
        branch,
        expanded: open,
        position: i + 1,
        siblings: list.length,
      });
      if (open) walk(s.children!, level + 1, s.nodeId);
    });
  };
  walk(shapes ?? [], 1, null);
  return out;
}

/**
 * The groups that hold the elements in `ids`, outermost first: the pane opens them, so a
 * selected child of a group shows. @internal
 */
export function selectionPaneAncestors(
  shapes: readonly ShapeSummary[] | undefined,
  ids: ReadonlySet<string>
): string[] {
  const out: string[] = [];
  const walk = (list: readonly ShapeSummary[], path: readonly string[]): void => {
    for (const s of list) {
      if (ids.has(s.nodeId)) for (const id of path) if (!out.includes(id)) out.push(id);
      if (s.children?.length) walk(s.children, [...path, s.nodeId]);
    }
  };
  walk(shapes ?? [], []);
  return out;
}

/** The elements the selection holds, as the pane marks them. @internal */
export function selectionPaneSelected(selection: EditorSelection): ReadonlySet<string> {
  if (selection.kind === 'shapes') return new Set(selection.nodeIds);
  if (selection.kind === 'text' || selection.kind === 'cells') return new Set([selection.nodeId]);
  return new Set();
}

/**
 * What a key does on the focused row of the pane (the tree pattern of WAI-ARIA): Up and Down
 * move, Home and End go to the ends, Right opens a group or moves into it, Left closes a group
 * or moves to its parent. F2 renames. Null for any other key. @internal
 */
export function selectionPaneKey(
  rows: readonly SelectionPaneRow[],
  focused: string | null,
  key: string
):
  | { readonly focus: string }
  | { readonly expand: string; readonly open: boolean }
  | { readonly rename: string }
  | null {
  if (rows.length === 0) return null;
  const at = Math.max(
    0,
    rows.findIndex((r) => r.nodeId === focused)
  );
  const row = rows[at]!;
  switch (key) {
    case 'ArrowDown':
      return { focus: rows[Math.min(rows.length - 1, at + 1)]!.nodeId };
    case 'ArrowUp':
      return { focus: rows[Math.max(0, at - 1)]!.nodeId };
    case 'Home':
      return { focus: rows[0]!.nodeId };
    case 'End':
      return { focus: rows[rows.length - 1]!.nodeId };
    case 'ArrowRight':
      if (!row.branch) return null;
      return row.expanded ? { focus: rows[at + 1]!.nodeId } : { expand: row.nodeId, open: true };
    case 'ArrowLeft':
      if (row.branch && row.expanded) return { expand: row.nodeId, open: false };
      return row.parent ? { focus: row.parent } : null;
    case 'F2':
      return { rename: row.nodeId };
    default:
      return null;
  }
}

/**
 * Select the element of a row. With `extend` (Ctrl, Cmd, or Shift), the click adds the element
 * to the selection or takes it out, among the elements of the same group; an element of another
 * group starts a new selection. A hidden element can be selected: the pane is the only place
 * that reaches it. @internal
 */
export function selectFromPane(
  editor: PptxEditorInstance,
  row: SelectionPaneRow,
  extend: boolean
): void {
  const group = row.parent ?? undefined;
  const sel = editor.snapshot().selection;
  const same = sel.kind === 'shapes' && (sel.group ?? null) === row.parent;
  let ids: readonly string[] = [row.nodeId];
  if (extend && same) {
    ids = sel.nodeIds.includes(row.nodeId)
      ? sel.nodeIds.filter((id) => id !== row.nodeId)
      : [...sel.nodeIds, row.nodeId];
  }
  editor.setSelection(
    ids.length === 0
      ? { kind: 'none' }
      : { kind: 'shapes', nodeIds: ids, ...(group ? { group } : {}) }
  );
}

/** The top-level elements of the slide: Show all and Hide all change these. @internal */
export function selectionPaneTopIds(shapes: readonly ShapeSummary[] | undefined): string[] {
  return (shapes ?? []).map((s) => s.nodeId);
}

/** The longest name the pane writes, as `renameShape` takes it. @internal */
export const SELECTION_PANE_NAME_MAX = SHAPE_NAME_MAX;

/** The tabs of the side panel, in order. @internal */
export const SIDE_PANEL_TABS = ['format', 'selection', 'comments'] as const;

/** A tab of the side panel: Format, Selection (the selection pane), or Comments. @public */
export type SidePanelTab = (typeof SIDE_PANEL_TABS)[number];
