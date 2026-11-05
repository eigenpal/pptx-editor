// Keyboard navigation for chrome menus, as DOM helpers rather than an item registry. Shared by
// the React and Vue adapters (moved from packages/react/src/menu-keyboard.ts).
//
// Adapted from DOCX Editor (eigenpal/docx-editor, packages/react/src/editor/menu/menu-keyboard.ts):
// the same rule that a `menu` role is a promise. A screen reader hands the arrow keys to the
// widget, so the keys must work. Items are found in the DOM, so a host's own rows join the
// navigation without a registry. PPTX Editor adds one thing: swatch grids, where Up and Down
// move by a row of the grid.

/**
 * Every row of a menu or a list inside a popover, in visual order. A field marked
 * `data-menu-field` (the custom color's hex field) is a stop of the arrow keys too.
 */
const ITEM_SELECTOR =
  '[role="menuitem"]:not(:disabled), [role="menuitemcheckbox"]:not(:disabled), [role="menuitemradio"]:not(:disabled), [role="option"]:not(:disabled), [data-menu-field]:not(:disabled)';

/** The rows a panel owns. @internal */
export function panelItems(panel: HTMLElement): HTMLElement[] {
  return [...panel.querySelectorAll<HTMLElement>(ITEM_SELECTOR)];
}

/** Move focus within a list by a step, wrapping at both ends. */
function focusBy(items: readonly HTMLElement[], from: Element | null, step: number): boolean {
  if (items.length === 0) return false;
  const index = from ? items.indexOf(from as HTMLElement) : -1;
  const next =
    index === -1 ? (step > 0 ? 0 : items.length - 1) : (index + step + items.length) % items.length;
  items[next]?.focus();
  return true;
}

/** Focus the first or last row of a list. @internal */
export function focusEdge(items: readonly HTMLElement[], edge: 'first' | 'last'): boolean {
  const target = edge === 'first' ? items[0] : items[items.length - 1];
  if (!target) return false;
  target.focus();
  return true;
}

/**
 * Up or Down from a swatch in a grid (`data-columns`): the swatch one row above or below.
 * From the top or bottom row, focus leaves the grid to the row before or after it.
 */
function gridStep(items: readonly HTMLElement[], from: HTMLElement, down: boolean): boolean {
  const grid = from.closest<HTMLElement>('[data-columns]');
  const columns = Number(grid?.dataset.columns);
  if (!grid || !(columns > 0)) return false;
  const cells = items.filter((i) => grid.contains(i));
  const index = cells.indexOf(from);
  const target = cells[index + (down ? columns : -columns)];
  if (target) {
    target.focus();
    return true;
  }
  // Leave the grid: the item after its last cell, or before its first cell.
  const edge = items.indexOf(down ? cells[cells.length - 1]! : cells[0]!);
  return focusBy(items, items[edge] ?? null, down ? 1 : -1);
}

/**
 * Handle one key in a popover. Returns true when the key was used. Up and Down move through
 * the rows (by grid rows in a swatch grid), Left and Right move through a swatch grid, and
 * Home and End go to the first and last row. A key typed in a text field is left alone,
 * except Down, which moves from a search field to the first row, and Up and Down from a field
 * among the rows (`data-menu-field`), which move to the next row. @internal
 */
export function handleMenuKey(panel: HTMLElement, key: string): boolean {
  const items = panelItems(panel);
  const focused = panel.ownerDocument.activeElement as HTMLElement | null;
  const inField = !!focused && /^(INPUT|TEXTAREA|SELECT)$/.test(focused.tagName);
  if (inField) {
    // A field among the rows moves to the row before or after it; a search field above the
    // list moves down to the first row.
    const among = items.includes(focused!);
    if (key === 'ArrowDown') return among ? focusBy(items, focused, 1) : focusEdge(items, 'first');
    return key === 'ArrowUp' && among ? focusBy(items, focused, -1) : false;
  }
  const inGrid = !!focused?.closest('[data-columns]') && items.includes(focused!);
  switch (key) {
    case 'ArrowDown':
      return (inGrid && gridStep(items, focused!, true)) || focusBy(items, focused, 1);
    case 'ArrowUp':
      return (inGrid && gridStep(items, focused!, false)) || focusBy(items, focused, -1);
    case 'ArrowRight':
      return inGrid ? focusBy(items, focused, 1) : false;
    case 'ArrowLeft':
      return inGrid ? focusBy(items, focused, -1) : false;
    case 'Home':
      return focusEdge(items, 'first');
    case 'End':
      return focusEdge(items, 'last');
    default:
      return false;
  }
}
