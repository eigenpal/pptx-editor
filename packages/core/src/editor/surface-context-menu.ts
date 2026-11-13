// Right-click on the slide. The surface stops the browser's own menu and sets the selection
// before its context menu opens: a shape that is not selected becomes the
// selection; a shape in the selection keeps it, and so does the shape whose text is being
// edited, where a right-click outside the text selection puts the caret at the click (after
// DOCX Editor 627e74ab7); empty slide area clears the selection. A right-click on a table
// cell places the caret in that cell (a cell range in the table stays, and so does a table
// selected as an object), so the table menu acts on the cell under the pointer. The chrome then reads the selection and
// draws its menu at the event's point, so no DOM type crosses the editor contract. A right-click
// on the table's select handle selects the table as an object.

import type { EditorSelection, TableCellAddress, TextPoint } from '../contracts/index.ts';
import type { LayoutItem, SlideLayout } from '../layout/index.ts';
import { itemById } from '../automation/commands/placement.ts';
import { pointAt } from './surface-geometry.ts';
import { comparePoints } from '../automation/commands/text-model.ts';
import { parentGroupOf, selectionScope } from '../automation/commands/placement.ts';
import type { SurfaceHost } from './surface.ts';

type Point = { x: number; y: number };

/** The selection a right-click on `id` (or on empty slide area) gives, or null to keep it. */
export function contextMenuSelection(
  selection: EditorSelection,
  id: string | undefined,
  group: string | undefined,
  caret?: Extract<EditorSelection, { kind: 'text' }>
): EditorSelection | null {
  if (!id) return selection.kind === 'none' ? null : { kind: 'none' };
  // A table selected as an object stays selected: its menu is the object menu.
  if (selection.kind === 'shapes' && selection.nodeIds.includes(id)) return null;
  // The text selection under the pointer stays; elsewhere in its text, the caret goes there.
  const keeps = (sel: Extract<EditorSelection, { kind: 'text' }>, at: TextPoint): boolean => {
    const [from, to] =
      comparePoints(sel.anchor, sel.focus) <= 0 ? [sel.anchor, sel.focus] : [sel.focus, sel.anchor];
    return comparePoints(from, at) <= 0 && comparePoints(at, to) <= 0;
  };
  if (caret?.cell && !(selection.kind === 'cells' && selection.nodeId === id)) {
    const same =
      selection.kind === 'text' &&
      selection.nodeId === id &&
      selection.cell?.row === caret.cell.row &&
      selection.cell?.column === caret.cell.column;
    return same && keeps(selection, caret.focus) ? null : caret;
  }
  if (
    caret &&
    !caret.cell &&
    selection.kind === 'text' &&
    selection.nodeId === id &&
    !selection.cell &&
    selection.diagramNode === undefined &&
    !keeps(selection, caret.focus)
  )
    return caret;
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
  const caret = id ? cellAt?.(id, p) : undefined;
  const next = contextMenuSelection(selection, id, inScope ? scope : undefined, caret);
  if (next) editor.select(next);
}

/**
 * A caret at slide point `p` in the table cell (`cellItemAt`) or the text of `id`, for a
 * right-click; none in view mode.
 */
export function menuCaret(
  editor: Pick<SurfaceHost, 'layout' | 'readOnly'>,
  id: string,
  p: Point,
  cellItemAt: (layout: SlideLayout) => { item: LayoutItem; cell: TableCellAddress } | undefined
): Extract<EditorSelection, { kind: 'text' }> | undefined {
  const layout = editor.layout();
  if (!layout || editor.readOnly()) return undefined;
  const at = cellItemAt(layout);
  const item = at?.item ?? itemById(layout, id);
  const point = item ? pointAt(item, p.x, p.y) : undefined;
  if (!point) return undefined;
  return { kind: 'text', nodeId: id, ...(at && { cell: at.cell }), anchor: point, focus: point };
}

/**
 * Bind the presses and right-clicks of the slide `stage`, and of the canvas `root` around it.
 * The canvas acts as the slide: a drag there selects the shapes it encloses, a click clears the
 * selection, and a right-click opens the slide menu. The host scrolls, so the root has no scroll
 * bar to press. @internal
 */
export function bindPresses(
  root: HTMLElement,
  stage: HTMLElement,
  press: (e: PointerEvent) => void,
  menu: (e: MouseEvent) => void
): void {
  stage.addEventListener('pointerdown', press);
  stage.addEventListener('contextmenu', menu);
  root.addEventListener('pointerdown', (e) => {
    if (e.target === root) press(e);
  });
  root.addEventListener('contextmenu', (e) => {
    if (e.target === root) menu(e);
  });
}
