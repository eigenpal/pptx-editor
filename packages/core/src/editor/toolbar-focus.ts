// One Tab stop and arrow keys for a toolbar (WAI-ARIA toolbar pattern).
// Framework-free, so the React and Vue toolbars move focus the same way. A disabled control
// carries `aria-disabled`, not `disabled`, so it stays reachable and a keyboard user can hear
// its reason. DOCX Editor's toolbar has no roving focus and removes a disabled button from
// focus. PPTX Editor keeps it reachable, so a keyboard user hears why it is off.

const ITEMS = 'button, input, select, textarea, a[href], [tabindex]';

/** The controls of a toolbar in order: not those inside an open menu, panel, or dialog. @internal */
export function toolbarItems(bar: HTMLElement): HTMLElement[] {
  return [...bar.querySelectorAll<HTMLElement>(ITEMS)].filter(
    (el) =>
      !el.hidden &&
      !(el as HTMLButtonElement).disabled &&
      el.closest('[hidden], .pptx-menu, [role="dialog"], [role="menu"]') === null &&
      el.getAttribute('aria-hidden') !== 'true'
  );
}

/**
 * Give the toolbar one Tab stop: the control that last had focus, else the first. Run it after
 * each render, since controls come and go. @internal
 */
export function syncToolbarTabStop(bar: HTMLElement, current?: Element | null): void {
  const items = toolbarItems(bar);
  if (!items.length) return;
  const stop =
    (current && items.find((el) => el === current)) ??
    items.find((el) => el.tabIndex === 0 && el.dataset.rovingStop === '') ??
    items[0]!;
  for (const el of items) {
    el.tabIndex = el === stop ? 0 : -1;
    if (el === stop) el.dataset.rovingStop = '';
    else delete el.dataset.rovingStop;
  }
}

/**
 * Keep the focus in the toolbar when the control that had it leaves the bar: the bar shows the
 * groups of the selection, so Delete or Escape can remove the focused control with its group.
 * When `last` is no longer in the document and nothing else has the focus, the bar's Tab stop
 * takes it, as the WAI-ARIA toolbar pattern keeps focus inside the toolbar.
 */
function keepToolbarFocus(bar: HTMLElement, last: Element | null): void {
  if (!last || last.isConnected) return;
  const doc = bar.ownerDocument;
  const active = doc.activeElement;
  // Some engines keep a removed element as the active one; that focus is lost too.
  if (active && active !== doc.body && active.isConnected) return;
  const items = toolbarItems(bar);
  (items.find((el) => el.tabIndex === 0) ?? items[0])?.focus();
}

/**
 * What `toolbarFocusTracker` returns: one call per focus event, and one after each render.
 * @internal
 */
export interface ToolbarFocusTracker {
  /** A control of the bar took the focus. A target outside the bar (a portalled menu) is none. */
  readonly focus: (bar: HTMLElement, target: EventTarget | null) => void;
  /** The focus left a control. `related` is where it went, or null. */
  readonly blur: (
    bar: HTMLElement,
    target: EventTarget | null,
    related: EventTarget | null
  ) => void;
  /** After a render: give the focus back if the focused control left the bar. */
  readonly restore: (bar: HTMLElement) => void;
}

/**
 * Track the control that has the focus, so the bar can keep the focus when that control leaves
 * with its group. Only a control that leaves the document while it has the focus counts: focus
 * that the user moves elsewhere (to the slide, a menu, the page) clears the record, and the
 * record is used once. @internal
 */
export function toolbarFocusTracker(): ToolbarFocusTracker {
  let last: Element | null = null;
  return {
    focus: (bar, target) => {
      const el = target as Element | null;
      last = el && bar.contains(el) ? el : null;
    },
    blur: (bar, target, related) => {
      const to = related as Node | null;
      if (to) {
        if (!bar.contains(to)) last = null;
        return;
      }
      // No new focus: the control left the document (keep the record) or the user clicked
      // something that takes no focus (clear it once the event is over).
      const el = target as Element | null;
      queueMicrotask(() => {
        if (el?.isConnected && last === el) last = null;
      });
    },
    restore: (bar) => {
      keepToolbarFocus(bar, last);
      if (last && !last.isConnected) last = null;
    },
  };
}

interface FieldLike {
  readonly localName?: string;
  readonly isContentEditable?: boolean;
  readonly value?: string;
  readonly selectionStart?: number | null;
  readonly selectionEnd?: number | null;
}

/**
 * Whether a field keeps the key. A field in a toolbar keeps Home and End for its caret, and
 * keeps ArrowLeft and ArrowRight until the caret is at that end of the text, as the WAI-ARIA
 * toolbar pattern says: so the arrows still leave a field such as Font size.
 */
function fieldOwnsKey(target: EventTarget | null, key: string, forward: boolean): boolean {
  const el = target as FieldLike | null;
  if (!el) return false;
  if (el.localName === 'textarea' || el.isContentEditable) return true;
  if (el.localName !== 'input' && el.localName !== 'select') return false;
  if (key === 'Home' || key === 'End') return true;
  if (el.localName === 'select') return false;
  const start = el.selectionStart ?? null;
  const end = el.selectionEnd ?? null;
  // An input with no caret (a number input) gives the arrows to the toolbar.
  if (start === null || end === null) return false;
  if (start !== end) return true;
  return forward ? end < (el.value ?? '').length : start > 0;
}

/**
 * Move focus among a toolbar's controls: ArrowRight and ArrowLeft (swapped in a right-to-left
 * interface), Home, and End. True when it moved focus, so the caller prevents the default.
 * @internal
 */
export function toolbarArrowKey(
  bar: HTMLElement,
  event: Pick<KeyboardEvent, 'key' | 'target'>,
  rtl = false
): boolean {
  const next = rtl ? 'ArrowLeft' : 'ArrowRight';
  const previous = rtl ? 'ArrowRight' : 'ArrowLeft';
  if (fieldOwnsKey(event.target, event.key, event.key === next)) return false;
  const items = toolbarItems(bar);
  const at = items.indexOf(event.target as HTMLElement);
  if (at === -1) return false;
  const to =
    event.key === next
      ? (at + 1) % items.length
      : event.key === previous
        ? (at - 1 + items.length) % items.length
        : event.key === 'Home'
          ? 0
          : event.key === 'End'
            ? items.length - 1
            : -1;
  if (to === -1) return false;
  const target = items[to]!;
  syncToolbarTabStop(bar, target);
  target.focus();
  return true;
}
