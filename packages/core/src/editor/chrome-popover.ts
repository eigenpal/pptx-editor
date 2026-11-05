// The behavior of the chrome popovers and cascading submenus, shared by the React and Vue
// adapters (moved from packages/react/src/chrome.tsx and ContextMenu.tsx): where a panel opens,
// which element takes focus, and when a panel closes. The adapters keep only their glue (a
// portal, refs, and effects), so the two cannot drift.

import { inFixedView, placeInFixedFrame } from './fixed-frame.ts';
import { dialogFocusables } from './dialog-draft.ts';
import { focusEdge, handleMenuKey, panelItems } from './menu-keyboard.ts';

/** Where a popover panel goes, in viewport pixels. @internal */
export interface PopoverPlacement {
  readonly left: number;
  readonly top: number;
  readonly maxHeight: number;
  /** Whether the panel is taller than its room and scrolls. */
  readonly scroll: boolean;
}

/** The space between the trigger and the panel, and the margin to the window edge. */
const GAP = 6;
const EDGE = 8;

/**
 * Place a panel under (or over) its trigger: the side with more room, so a short window does not
 * flip the panel over its trigger; the panel scrolls within that room. @internal
 */
export function placePopover(
  anchor: {
    readonly left: number;
    readonly right: number;
    readonly top: number;
    readonly bottom: number;
  },
  panel: { readonly width: number; readonly height: number },
  align: 'start' | 'end',
  viewport: { readonly width: number; readonly height: number }
): PopoverPlacement {
  const { width, height } = panel;
  const roomBelow = viewport.height - anchor.bottom - GAP - EDGE;
  const roomAbove = anchor.top - GAP - EDGE;
  const below = height <= roomBelow || roomBelow >= roomAbove;
  const room = Math.max(0, below ? roomBelow : roomAbove);
  return {
    left: Math.max(
      EDGE,
      Math.min(align === 'end' ? anchor.right - width : anchor.left, viewport.width - width - EDGE)
    ),
    top: below ? anchor.bottom + GAP : Math.max(EDGE, anchor.top - Math.min(height, room) - GAP),
    maxHeight: room,
    // Scroll only a panel taller than its room, so a short one never clips a nested layer.
    scroll: height > room,
  };
}

/**
 * Measure `panel` against `anchor` and write its place to the node at once: a browser does not
 * focus into a `visibility: hidden` panel, and the focus step runs before the adapter re-renders.
 * Returns the placement. @internal
 */
export function positionPopover(
  panel: HTMLElement | null,
  anchor: HTMLElement,
  align: 'start' | 'end'
): PopoverPlacement {
  const anchorRect = anchor.getBoundingClientRect();
  const next = placeInFixedFrame(
    panel?.parentElement,
    { width: panel?.offsetWidth ?? 200, height: panel?.offsetHeight ?? 0 },
    (size, area) => placePopover(inFixedView(anchorRect, area), size, align, area)
  );
  if (panel) {
    panel.style.left = `${next.left}px`;
    panel.style.top = `${next.top}px`;
    panel.style.maxHeight = `${next.maxHeight}px`;
    panel.style.overflowY = next.scroll ? 'auto' : '';
    panel.style.visibility = '';
  }
  return next;
}

/** How a popover takes focus and handles keys. @internal */
export interface PopoverFocusOptions {
  readonly role: 'menu' | 'dialog';
  /** A dialog panel opened from the keyboard takes focus at its first control. */
  readonly focusFirst: boolean;
  /** A dialog with a field to type in: its field takes focus for every user. */
  readonly field?: boolean;
  /**
   * A panel that takes no focus when it opens (the link popup's reading mode): the keys stay
   * in the slide text, so typing goes on. A click reaches it, and F6 moves focus into it
   * (`listenLinkPanelKey`); Tab does not, as Tab in slide text types a tab.
   */
  readonly passive?: boolean;
}

/** The attribute a panel carries while a pointer opened it: rows show no focus ring. @internal */
export const POINTER_OPEN_ATTRIBUTE = 'data-pointer-open';

/**
 * Move focus into a popover that just opened, so the keys reach it and not the slide text or a
 * side-panel field. A `field` always takes focus. When the trigger has focus (a keyboard user),
 * a menu focuses its first row and a `focusFirst` dialog its first control. Otherwise (a pointer
 * user) the panel itself takes focus, without a scroll, as DOCX Editor's context menu does: the
 * arrow keys then move to the first or last row, Escape returns to the trigger, and the panel
 * carries `data-pointer-open` until a key is pressed, so no focus ring shows for the pointer.
 * Returns the element that had focus before a pointer open, for `restorePopoverFocus`. @internal
 */
export function focusPopoverOnOpen(
  panel: HTMLElement | null,
  anchor: HTMLElement | null,
  options: PopoverFocusOptions
): HTMLElement | null {
  if (options.passive) return null;
  const input = options.field ? panel?.querySelector<HTMLElement>('input, textarea') : null;
  if (input) {
    input.focus();
    return null;
  }
  const doc = panel?.ownerDocument;
  if (!panel || !doc) return null;
  const active = doc.activeElement as HTMLElement | null;
  // A menu with a chosen row (the zoom level) opens on that row, as DOCX Editor's mode menu
  // does: the WAI-ARIA menu pattern puts focus on the checked radio row.
  const checked = options.role === 'menu' ? checkedMenuRow(panel) : null;
  if (anchor?.contains(active)) {
    if (checked) checked.focus();
    else if (options.role === 'menu') focusEdge(panelItems(panel), 'first');
    else if (options.focusFirst) dialogFocusables(panel)[0]?.focus();
    return null;
  }
  // Focus already inside the panel (a host part focused itself) stays.
  if (active && panel.contains(active)) return null;
  focusPanelFromPointer(panel);
  // After a pointer open the checked row takes focus too; `data-pointer-open` hides its ring.
  checked?.focus();
  return active && active !== doc.body ? active : null;
}

/** The enabled checked radio row of a menu panel (not of a nested panel), if any. */
function checkedMenuRow(panel: HTMLElement): HTMLElement | null {
  const row = panel.querySelector<HTMLElement>(
    '[role="menuitemradio"][aria-checked="true"]:not(:disabled):not([aria-disabled="true"])'
  );
  return row && row.closest('.pptx-menu, [role="menu"]') === panel ? row : null;
}

/**
 * Focus a panel that a pointer opened: the panel itself, without a scroll, marked with
 * `data-pointer-open` so its rows show no focus ring until a key is pressed (`panelKeyPressed`).
 * @internal
 */
export function focusPanelFromPointer(panel: HTMLElement): void {
  if (!panel.hasAttribute('tabindex')) panel.tabIndex = -1;
  panel.setAttribute(POINTER_OPEN_ATTRIBUTE, '');
  panel.focus({ preventScroll: true });
}

/** A key was pressed in a panel: its focus ring shows from now on. @internal */
export function panelKeyPressed(panel: HTMLElement | null): void {
  panel?.removeAttribute(POINTER_OPEN_ATTRIBUTE);
}

/** The last pointer position each menu panel saw, so a row takes focus only on a real move. */
const lastPointer = new WeakMap<HTMLElement, { x: number; y: number }>();

/**
 * The pointer moves the active row of a menu: a row under the pointer takes focus, so the
 * focused row and the hovered row are always one row. Without it, a menu opened with a click
 * keeps its first row focused while the pointer hovers another, and two rows look active.
 * Only a real move counts: a menu that opens under a resting pointer (the browser sends a
 * pointer event when the layout changes) keeps its focus, so the arrow keys start at the first
 * row. Rows of a nested panel are left to that panel. Call it on the panel's `mousemove`.
 * @internal
 */
export function focusMenuRowAtPointer(
  panel: HTMLElement | null,
  event: { readonly target: EventTarget | null; readonly clientX: number; readonly clientY: number }
): void {
  if (!panel) return;
  const before = lastPointer.get(panel);
  const at = { x: event.clientX, y: event.clientY };
  lastPointer.set(panel, at);
  if (!before || (before.x === at.x && before.y === at.y)) return;
  const el = event.target as Element | null;
  if (typeof el?.closest !== 'function') return;
  const row = el.closest<HTMLElement>('[role^="menuitem"]');
  if (!row || row.closest('[role="menu"]') !== panel) return;
  if (row.ownerDocument.activeElement !== row) row.focus({ preventScroll: true });
}

/**
 * When a popover closes with focus inside it, give focus back instead of dropping it on
 * `<body>`: to `back`, the element that had focus before a pointer open (the slide text, so
 * typing goes on), else to the trigger. Focus that the popover moved on purpose (Escape to the
 * trigger, a dialog, the slide text) stays. @internal
 */
export function restorePopoverFocus(
  anchor: HTMLElement | null,
  focusInside: boolean,
  back: HTMLElement | null = null
): void {
  if (!focusInside || !anchor?.isConnected) return;
  const active = anchor.ownerDocument.activeElement;
  if (active && active !== anchor.ownerDocument.body && active.isConnected) return;
  if (back?.isConnected) back.focus({ preventScroll: true });
  else anchor.focus();
}

/**
 * The parts of a key event the chrome helpers read: a DOM `KeyboardEvent` or a framework's
 * synthetic event, so `stopPropagation` stops the adapter's own propagation. @internal
 */
export interface ChromeKeyEvent {
  readonly key: string;
  readonly target: EventTarget | null;
  readonly shiftKey: boolean;
  preventDefault(): void;
  stopPropagation(): void;
}

/** The document of `el`, or the page's when there is no element yet. */
function ownDocument(el: Node | null | undefined): Document {
  return el?.ownerDocument ?? document;
}

/** The open popovers, innermost last: Escape closes only the last. */
const OPEN_POPOVERS: object[] = [];

/**
 * Close a popover on a press outside it, on Escape, and when focus leaves it. Returns the
 * function that stops listening. @internal
 */
export function listenPopoverDismiss(options: {
  readonly panel: () => HTMLElement | null;
  readonly anchor: HTMLElement | null;
  /** Whether focus has been inside the popover. */
  readonly focusInside: () => boolean;
  readonly onClose: () => void;
}): () => void {
  const { anchor, panel } = options;
  const down = (e: Event): void => {
    if (panel()?.contains(e.target as Node) || anchor?.contains(e.target as Node)) return;
    // A click in a menu opened from this one (the toolbar's More menu) keeps this one open.
    if ((e.target as Element | null)?.closest?.('.pptx-menu')) return;
    options.onClose();
  };
  // Escape closes the topmost popover, in the capture phase, after DOCX Editor's
  // `listenForPopupEscape`: the slide surface stops Escape before a bubbling listener would
  // see it. An IME composition keeps its Escape, a modal dialog (the editor's or the host's)
  // handles its own, and only an Escape from the editor is stopped; one Escape does one thing.
  const token = {};
  OPEN_POPOVERS.push(token);
  const key = (e: KeyboardEvent): void => {
    if (e.key !== 'Escape' || OPEN_POPOVERS[OPEN_POPOVERS.length - 1] !== token) return;
    if (e.isComposing || e.keyCode === 229) return;
    const target = e.target as Element | null;
    const modal = target?.closest?.('dialog');
    if (modal && !modal.contains(panel())) return;
    // An open submenu of this popover closes itself first, one level per Escape.
    if (target?.closest?.('.pptx-menu--submenu')) return;
    // The editor's Find bar keeps its own Escape (it closes the bar).
    if (target?.closest?.('.pptx-find')) return;
    // The editor, else the slide canvas a popup opened on (a root with no viewport).
    const scope = anchor?.closest('.pptx-editor') ?? anchor?.closest('.pptx-canvas');
    if (panel()?.contains(target) || scope?.contains(target) || anchor?.contains(target)) {
      e.stopPropagation();
      e.preventDefault();
    }
    // Escape from inside the popover returns to its trigger, also after a pointer open.
    if (panel()?.contains(target) && anchor?.isConnected) anchor.focus();
    options.onClose();
  };
  // Focus that leaves the popover for something else closes it, after DOCX Editor's
  // `listenForPopupFocusLeave` (Find from the More menu moves focus to the find bar).
  const focus = (e: FocusEvent): void => {
    const target = e.target as Node | null;
    if (!target || panel()?.contains(target) || anchor?.contains(target)) return;
    if ((target as Element).closest?.('.pptx-menu')) return;
    if (options.focusInside()) options.onClose();
  };
  // The window losing focus, a resize, or a scroll outside the panel close it too (`away`).
  // The document that holds the popover (an editor can live in another frame).
  const doc = ownDocument(anchor ?? panel());
  const stopAway = listenAway(panel, options.onClose, doc);
  doc.addEventListener('mousedown', down);
  // The slide surface prevents the default of a press on a handle, and then no `mousedown`
  // follows: the pointer press closes the popover as well.
  doc.addEventListener('pointerdown', down);
  doc.addEventListener('keydown', key, true);
  doc.addEventListener('focusin', focus);
  return () => {
    OPEN_POPOVERS.splice(OPEN_POPOVERS.indexOf(token), 1);
    stopAway();
    doc.removeEventListener('mousedown', down);
    doc.removeEventListener('pointerdown', down);
    doc.removeEventListener('keydown', key, true);
    doc.removeEventListener('focusin', focus);
  };
}

/**
 * Close a menu or popover when the page moves away from it: the window loses focus or resizes,
 * the page scrolls, or the user scrolls with the wheel outside the panel. A scroll that code
 * makes inside the page (the rail follows the current slide) does not close it. Returns the
 * function that stops listening. @internal
 */
export function listenAway(
  panel: () => HTMLElement | null,
  onClose: () => void,
  doc: Document = ownDocument(panel())
): () => void {
  const wheel = (e: Event): void => {
    if (!panel()?.contains(e.target as Node)) onClose();
  };
  const scroll = (e: Event): void => {
    if (e.target === doc || e.target === doc.documentElement) onClose();
  };
  // The browser's color selector of a custom color field opens outside the page: the window
  // loses focus, and the panel stays open for the color it returns.
  const blur = (): void => {
    const active = doc.activeElement;
    if (active?.matches('input[type="color"]') && panel()?.contains(active)) return;
    onClose();
  };
  const view = doc.defaultView;
  doc.addEventListener('wheel', wheel, { capture: true, passive: true });
  doc.addEventListener('scroll', scroll, true);
  view?.addEventListener('resize', onClose);
  view?.addEventListener('blur', blur);
  return () => {
    doc.removeEventListener('wheel', wheel, { capture: true });
    doc.removeEventListener('scroll', scroll, true);
    view?.removeEventListener('resize', onClose);
    view?.removeEventListener('blur', blur);
  };
}

/**
 * Close a context menu: on a press outside it (pointer or mouse, also a right press that opens
 * the next menu, so one menu shows at a time), on Escape anywhere, and when the page moves away
 * (`listenAway`). A press in a menu opened from it (`.pptx-menu`) keeps it. Returns the function
 * that stops listening. @internal
 */
export function listenContextMenuDismiss(
  panel: () => HTMLElement | null,
  onClose: () => void
): () => void {
  const down = (e: Event): void => {
    const target = e.target as Element | null;
    if (panel()?.contains(target) || target?.closest?.('.pptx-menu')) return;
    onClose();
  };
  // Escape with focus outside the menu (the slide kept it) closes it too.
  const key = (e: KeyboardEvent): void => {
    if (e.key === 'Escape') onClose();
  };
  const doc = ownDocument(panel());
  const stopAway = listenAway(panel, onClose, doc);
  doc.addEventListener('pointerdown', down, true);
  doc.addEventListener('mousedown', down, true);
  doc.addEventListener('keydown', key);
  return () => {
    stopAway();
    doc.removeEventListener('pointerdown', down, true);
    doc.removeEventListener('mousedown', down, true);
    doc.removeEventListener('keydown', key);
  };
}

/**
 * A key pressed inside a popover panel: Escape closes it and returns to the trigger, Tab leaves
 * it (a menu at once, a `focusFirst` dialog past its first or last control), and the menu keys
 * move between rows. Keys from a popover opened inside this one are left alone. @internal
 */
export function popoverKeyDown(
  panel: HTMLElement | null,
  anchor: HTMLElement | null,
  event: ChromeKeyEvent,
  options: PopoverFocusOptions & { readonly onClose: () => void }
): void {
  // A key press shows the focus ring from now on.
  panelKeyPressed(panel);
  if (!panel?.contains(event.target as Node)) return;
  if (event.key === 'Escape') {
    event.preventDefault();
    event.stopPropagation();
    anchor?.focus();
    options.onClose();
    return;
  }
  if (event.key === 'Tab') {
    // Tab leaves the popover: it closes, and focus goes on from the trigger.
    const items = options.focusFirst ? dialogFocusables(panel) : [];
    const edge = event.shiftKey ? items[0] : items[items.length - 1];
    if (options.role === 'menu' || (edge !== undefined && event.target === edge)) {
      event.preventDefault();
      anchor?.focus();
      options.onClose();
    }
    return;
  }
  if (handleMenuKey(panel, event.key)) {
    event.preventDefault();
    event.stopPropagation();
  }
}

/** The editor scope a panel renders in, so it inherits the theme tokens. @internal */
export function popoverScope(el: Element | null | undefined): HTMLElement {
  return (el?.closest('.pptx-editor') as HTMLElement | null) ?? document.body;
}

/**
 * Give the caret back to the slide text after a toolbar field or picker is done, as DOCX
 * Editor's pickers call `editor.focus()`: a format armed at the caret applies to what is typed,
 * and the next key (Cmd+Z) reaches the slide. @internal
 */
export function focusTextAfterChrome(
  editor: { focus(): void } | null,
  from?: Element | null
): void {
  // The editor focuses the text caret or the slide surface, by the selection.
  editor?.focus();
  const root = from?.closest('.pptx-editor');
  const active = from?.ownerDocument.activeElement;
  if (active && root?.contains(active) && !active.closest('.pptx-toolbar, .pptx-menu')) return;
  // No surface took it (none attached): give it to the slide text's input.
  root?.querySelector<HTMLElement>('.pptx-surface__input')?.focus({ preventScroll: true });
}

/**
 * Place a cascading submenu to the right of its row, or to its left at the window's edge.
 * @internal
 */
export function placeSubmenu(
  row: { readonly left: number; readonly right: number; readonly top: number },
  panel: { readonly width: number; readonly height: number },
  viewport: { readonly width: number; readonly height: number }
): { readonly left: number; readonly top: number } {
  const right = row.right + panel.width + EDGE <= viewport.width;
  return {
    left: right ? row.right : Math.max(EDGE, row.left - panel.width),
    top: Math.max(EDGE, Math.min(row.top - 4, viewport.height - panel.height - EDGE)),
  };
}

/** Whether ArrowLeft moves within a grid row of a submenu, rather than closing it. @internal */
export function submenuLeftInGrid(target: Element): boolean {
  const grid = target.closest<HTMLElement>('[data-columns]');
  const columns = Number(grid?.dataset.columns);
  if (!grid || !(columns > 0)) return false;
  return panelItems(grid).indexOf(target as HTMLElement) % columns !== 0;
}

/**
 * A key on a submenu's panel: ArrowLeft (outside a grid row), Escape, and Tab close it and
 * return to its row; the menu keys move between rows. True when the submenu must close.
 * @internal
 */
export function submenuPanelKey(panel: HTMLElement | null, event: ChromeKeyEvent): boolean {
  panelKeyPressed(panel);
  const back = event.key === 'ArrowLeft' && !submenuLeftInGrid(event.target as Element);
  if (back || event.key === 'Escape' || event.key === 'Tab') {
    event.preventDefault();
    event.stopPropagation();
    return true;
  }
  if (panel && handleMenuKey(panel, event.key)) {
    event.preventDefault();
    event.stopPropagation();
  }
  return false;
}
