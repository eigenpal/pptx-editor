// Chrome tooltips, after DOCX Editor's `Tooltip` (packages/react/src/components/ui/Tooltip.tsx):
// a small bubble under a control after a 400 ms hover delay, in place of the native `title`.
// A control names its tooltip with `data-pptx-tooltip` (and `data-pptx-tooltip-reason` when it is
// disabled for a reason); one listener per document finds the control under the pointer or the
// keyboard focus. The React and Vue `ChromeTooltip` parts render the bubble with the same markup;
// this module owns the timing and the position, so the two cannot drift.
//
// The bubble is visual only (`aria-hidden`): the control's `aria-label` is its accessible name,
// and a disabled control keeps its reason in an `aria-describedby` note, as before.

import { popoverScope } from './chrome-popover.ts';

/** The hover delay before a tooltip opens, as DOCX Editor's `Tooltip` (`delayMs = 400`). @internal */
export const TOOLTIP_DELAY_MS = 400;
/**
 * After a tooltip closes, the next one opens at once within this time (moving along a bar).
 * @internal
 */
export const TOOLTIP_SKIP_DELAY_MS = 300;
/** The label attribute a control carries. @internal */
export const TOOLTIP_LABEL_ATTRIBUTE = 'data-pptx-tooltip';
/** The reason attribute a disabled control carries. @internal */
export const TOOLTIP_REASON_ATTRIBUTE = 'data-pptx-tooltip-reason';

const GAP = 8;
const EDGE = 8;

/** The tooltip attributes of a control. @internal */
export interface ChromeTooltipAttributes {
  readonly 'data-pptx-tooltip': string;
  readonly 'data-pptx-tooltip-reason'?: string;
}

/**
 * The tooltip attributes of a control: its label, and its reason when it is disabled. The
 * reason shows on its own line in the bubble. @internal
 */
export function tooltipAttributes(
  label: string,
  disabled?: boolean,
  reason?: string | null
): ChromeTooltipAttributes {
  return disabled && reason
    ? { 'data-pptx-tooltip': label, 'data-pptx-tooltip-reason': reason }
    : { 'data-pptx-tooltip': label };
}

/** The tooltip on screen: the control, its text, and the editor scope it renders in. @internal */
export interface ChromeTooltipState {
  readonly target: HTMLElement;
  readonly label: string;
  readonly reason: string | null;
  /** The nearest `.pptx-editor`, so the bubble inherits the theme tokens and dark mode. */
  readonly scope: HTMLElement;
}

/** Where the bubble goes: centered under the control, or over it when there is no room. @internal */
export function placeTooltip(
  anchor: {
    readonly left: number;
    readonly right: number;
    readonly top: number;
    readonly bottom: number;
  },
  bubble: { readonly width: number; readonly height: number },
  viewport: { readonly width: number; readonly height: number }
): { readonly left: number; readonly top: number; readonly side: 'bottom' | 'top' } {
  const below = anchor.bottom + GAP + bubble.height + EDGE <= viewport.height;
  const above = anchor.top - GAP - bubble.height >= EDGE;
  const side = below || !above ? 'bottom' : 'top';
  const center = (anchor.left + anchor.right) / 2;
  const left = Math.max(
    EDGE,
    Math.min(center - bubble.width / 2, viewport.width - bubble.width - EDGE)
  );
  const top = side === 'bottom' ? anchor.bottom + GAP : anchor.top - GAP - bubble.height;
  return { left: Math.round(left), top: Math.round(top), side };
}

/** Place a rendered bubble for its control: `left`, `top`, and `data-side`. @internal */
export function positionChromeTooltip(bubble: HTMLElement, target: HTMLElement): void {
  const win = bubble.ownerDocument.defaultView;
  const rect = target.getBoundingClientRect();
  const at = placeTooltip(
    rect,
    { width: bubble.offsetWidth, height: bubble.offsetHeight },
    { width: win?.innerWidth ?? 0, height: win?.innerHeight ?? 0 }
  );
  bubble.style.left = `${at.left}px`;
  bubble.style.top = `${at.top}px`;
  bubble.setAttribute('data-side', at.side);
}

/** The control with a tooltip at or around a node. */
function tooltipTarget(node: EventTarget | null): HTMLElement | null {
  const el = node as Element | null;
  if (!el || typeof el.closest !== 'function') return null;
  return el.closest<HTMLElement>(`[${TOOLTIP_LABEL_ATTRIBUTE}]`);
}

/** Called with the tooltip to show, or null to hide it. @internal */
export type ChromeTooltipListener = (state: ChromeTooltipState | null) => void;

interface Controller {
  readonly listeners: ChromeTooltipListener[];
  state: ChromeTooltipState | null;
  dispose(): void;
}

const controllers = new WeakMap<Document, Controller>();

function createController(doc: Document): Controller {
  const win = doc.defaultView!;
  let timer: number | null = null;
  let pending: HTMLElement | null = null;
  let shown: HTMLElement | null = null;
  // The control a click or Escape closed: it opens again only after the pointer leaves it.
  let suppressed: HTMLElement | null = null;
  let source: 'pointer' | 'focus' = 'pointer';
  let closedAt = -Infinity;
  // The last input was a key, so a focus move shows the tooltip (`:focus-visible`).
  let keyboard = false;

  const controller: Controller = { listeners: [], state: null, dispose: () => {} };
  const notify = (): void => controller.listeners[0]?.(controller.state);

  const read = (el: HTMLElement): ChromeTooltipState | null => {
    const label = el.getAttribute(TOOLTIP_LABEL_ATTRIBUTE);
    // A control whose menu is open shows the menu, not its tooltip, as DOCX Editor's pickers.
    if (!label || !el.isConnected || el.getAttribute('aria-expanded') === 'true') return null;
    return {
      target: el,
      label,
      reason: el.getAttribute(TOOLTIP_REASON_ATTRIBUTE) || null,
      scope: popoverScope(el),
    };
  };
  const cancel = (): void => {
    if (timer !== null) win.clearTimeout(timer);
    timer = null;
    pending = null;
  };
  const hide = (): void => {
    cancel();
    if (!shown) return;
    shown = null;
    controller.state = null;
    closedAt = Date.now();
    notify();
  };
  const show = (el: HTMLElement): void => {
    cancel();
    const state = read(el);
    shown = state ? el : null;
    controller.state = state;
    notify();
  };
  const request = (el: HTMLElement, from: 'pointer' | 'focus'): void => {
    if (el === suppressed || el === shown || el === pending) return;
    const warm = shown !== null || Date.now() - closedAt < TOOLTIP_SKIP_DELAY_MS;
    hide();
    source = from;
    if (warm) return show(el);
    pending = el;
    timer = win.setTimeout(() => show(el), TOOLTIP_DELAY_MS);
  };
  const dismiss = (): void => {
    suppressed = shown ?? pending;
    hide();
    closedAt = -Infinity;
  };

  const onPointerOver = (e: PointerEvent | MouseEvent): void => {
    if ((e as PointerEvent).pointerType === 'touch') return;
    const el = tooltipTarget(e.target);
    if (el !== suppressed) suppressed = null;
    if (!el) return hide();
    request(el, 'pointer');
  };
  const onPointerOut = (e: PointerEvent | MouseEvent): void => {
    // The pointer left the window.
    if (!e.relatedTarget) (hide(), (suppressed = null));
  };
  const onFocusIn = (e: FocusEvent): void => {
    const el = tooltipTarget(e.target);
    if (el && keyboard) request(el, 'focus');
    else if (source === 'focus') hide();
  };
  const onFocusOut = (e: FocusEvent): void => {
    const el = tooltipTarget(e.target);
    if (source === 'focus' && el && (el === shown || el === pending)) hide();
  };
  const onKeyDown = (e: KeyboardEvent): void => {
    keyboard = true;
    if (e.key === 'Escape' && (shown || pending)) dismiss();
  };
  const onPointerDown = (): void => {
    keyboard = false;
    dismiss();
  };

  const options = { capture: true, passive: true } as const;
  const events: [string, EventListener][] = [
    ['pointerover', onPointerOver as EventListener],
    ['pointerout', onPointerOut as EventListener],
    ['focusin', onFocusIn as EventListener],
    ['focusout', onFocusOut as EventListener],
    ['keydown', onKeyDown as EventListener],
    ['pointerdown', onPointerDown],
    ['click', dismiss],
    ['wheel', hide],
    ['scroll', hide],
  ];
  for (const [type, fn] of events) doc.addEventListener(type, fn, options);
  win.addEventListener('blur', hide);
  controller.dispose = () => {
    cancel();
    for (const [type, fn] of events) doc.removeEventListener(type, fn, options);
    win.removeEventListener('blur', hide);
  };
  return controller;
}

/**
 * Watch a document for chrome tooltips. Every editor root on a page calls it; the first one
 * still mounted renders the bubble, so two editors never show two bubbles. Returns the
 * unsubscribe. @internal
 */
export function listenChromeTooltips(doc: Document, onChange: ChromeTooltipListener): () => void {
  if (!doc.defaultView) return () => {};
  let controller = controllers.get(doc);
  if (!controller) {
    controller = createController(doc);
    controllers.set(doc, controller);
  }
  const c = controller;
  c.listeners.push(onChange);
  if (c.listeners[0] === onChange && c.state) onChange(c.state);
  return () => {
    const owner = c.listeners[0] === onChange;
    const at = c.listeners.indexOf(onChange);
    if (at !== -1) c.listeners.splice(at, 1);
    if (c.listeners.length === 0) {
      c.dispose();
      controllers.delete(doc);
    } else if (owner) c.listeners[0]!(c.state);
  };
}
