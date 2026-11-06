// Right-click on the slide. The surface stops the browser's own menu and sets the selection
// before its context menu opens: a shape that is not selected becomes the
// selection; a shape in the selection keeps it, and so does the shape whose text is being
// edited; empty slide area clears the selection. A right-click on a table cell places the caret
// in that cell (a cell range in the table stays, and so does a table selected as an object), so
// the table menu acts on the cell under the pointer. The chrome then reads the selection and
// draws its menu at the event's point, so no DOM type crosses the editor contract. A right-click
// on the table's select handle selects the table as an object.

import type { EditorSelection } from '../contracts/index.ts';
import { parentGroupOf, selectionScope } from '../automation/commands/placement.ts';
import type { SurfaceHost } from './surface.ts';

type Point = { x: number; y: number };

/** The selection a right-click on `id` (or on empty slide area) gives, or null to keep it. */
export function contextMenuSelection(
  selection: EditorSelection,
  id: string | undefined,
  group: string | undefined,
  cell?: Extract<EditorSelection, { kind: 'text' }>
): EditorSelection | null {
  if (!id) return selection.kind === 'none' ? null : { kind: 'none' };
  // A table selected as an object stays selected: its menu is the object menu.
  if (selection.kind === 'shapes' && selection.nodeIds.includes(id)) return null;
  if (cell && !(selection.kind === 'cells' && selection.nodeId === id)) {
    const same =
      selection.kind === 'text' &&
      selection.nodeId === id &&
      selection.cell?.row === cell.cell?.row &&
      selection.cell?.column === cell.cell?.column;
    return same ? null : cell;
  }
  // A caret or a cell range in the table under the pointer stays, as a text selection does.
  if ((selection.kind === 'text' || selection.kind === 'cells') && selection.nodeId === id)
    return null;
  return { kind: 'shapes', nodeIds: [id], ...(group ? { group } : {}) };
}

const REOPENED = new WeakSet<Event>();

/**
 * Send a right-click on the handle again from the stage. The new selection repaints the overlay
 * and takes the handle out of the document before the event bubbles on, so the adapter that
 * opens the menu would not see it; the stage stays.
 */
function reopenFromStage(e: MouseEvent): void {
  const stage = e.currentTarget as globalThis.Element | null;
  const View = stage?.ownerDocument.defaultView?.MouseEvent;
  if (!stage || !View) return;
  e.stopPropagation();
  const again = new View('contextmenu', {
    bubbles: true,
    cancelable: true,
    clientX: e.clientX,
    clientY: e.clientY,
    screenX: e.screenX,
    screenY: e.screenY,
    button: e.button,
    buttons: e.buttons,
  });
  REOPENED.add(again);
  stage.dispatchEvent(again);
}

/** Handle a `contextmenu` event on the slide stage. */
export function selectForContextMenu(
  e: MouseEvent,
  editor: Pick<SurfaceHost, 'layout' | 'selection' | 'select'>,
  hit: (p: Point, scope?: string) => string | undefined,
  pointOf: (e: MouseEvent) => Point,
  cellAt?: (id: string, p: Point) => Extract<EditorSelection, { kind: 'text' }> | undefined
): void {
  e.preventDefault();
  // The event `reopenFromStage` sends again: the selection is already set.
  if (REOPENED.has(e)) return;
  const layout = editor.layout();
  if (!layout) return;
  const p = pointOf(e);
  const selection = editor.selection();
  // The table's select handle stands for the table: a right-click selects it as an object.
  const target = e.target as { closest?(selector: string): Element | null } | null;
  const handle = target?.closest?.('[data-table-handle]')?.getAttribute('data-table-handle');
  if (handle) {
    if (!(selection.kind === 'shapes' && selection.nodeIds.includes(handle))) {
      const group = parentGroupOf(layout.items, handle);
      editor.select({ kind: 'shapes', nodeIds: [handle], ...(group ? { group } : {}) });
      reopenFromStage(e);
    }
    return;
  }
  // Inside an entered group, the group's children are hit first, as a click does.
  const scope = selectionScope(layout.items, selection);
  const inScope = scope ? hit(p, scope) : undefined;
  const id = inScope ?? hit(p);
  const cell = id ? cellAt?.(id, p) : undefined;
  const next = contextMenuSelection(selection, id, inScope ? scope : undefined, cell);
  if (next) editor.select(next);
}
