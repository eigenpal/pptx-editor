// Fixed chrome (menus, popovers, tooltips) is placed in window coordinates. A host ancestor with
// `transform`, `filter`, `perspective`, `contain`, or a container type becomes the containing
// block of a `position: fixed` element: `left`/`top` then count from that ancestor, a scaled
// ancestor (or CSS `zoom`) scales every length, and the overflow of that ancestor and of the
// ancestors around it clips the panel. A probe in the same parent measures the frame's origin
// and scale, and the placement keeps the panel in the part of the window that stays visible.

/** A place in window coordinates, with an optional height limit. @internal */
export interface FixedPoint {
  readonly left: number;
  readonly top: number;
  readonly maxHeight?: number;
}

/** A panel's size. @internal */
export interface FixedSize {
  readonly width: number;
  readonly height: number;
}

/**
 * The area a fixed panel can show in, in window pixels: the window, or less inside a host that
 * clips its overflow. Placement rules see it as their viewport, at the origin. @internal
 */
export interface FixedView {
  readonly left: number;
  readonly top: number;
  readonly width: number;
  readonly height: number;
}

interface Frame {
  readonly x: number;
  readonly y: number;
  readonly sx: number;
  readonly sy: number;
}

const PROBE = 100;
// The values that make an element the containing block of its `position: fixed` descendants.
const NONE = /^(none|normal|auto|visible|flat)?$/;
const CONTAIN = /paint|layout|strict|content/;
const WILL_CHANGE = /transform|translate|scale|rotate|perspective|filter/;

/** The origin and scale of the frame that `position: fixed` children of `host` use. */
function measureFrame(host: Element): Frame | null {
  const doc = host.ownerDocument;
  const probe = doc.createElement('div');
  probe.setAttribute('aria-hidden', 'true');
  probe.style.cssText =
    `position:fixed;left:0;top:0;width:${PROBE}px;height:${PROBE}px;margin:0;padding:0;` +
    'border:0;visibility:hidden;pointer-events:none;';
  host.appendChild(probe);
  const rect = probe.getBoundingClientRect();
  probe.remove();
  // A layout-free DOM (tests, server) measures nothing.
  if (!(rect.width > 0) || !(rect.height > 0)) return null;
  return { x: rect.left, y: rect.top, sx: rect.width / PROBE, sy: rect.height / PROBE };
}

/** Whether an element's style makes it the containing block of fixed descendants. */
function holdsFixed(style: CSSStyleDeclaration): boolean {
  // A browser without a property reads it as undefined: that sets nothing.
  const set = (value: string | undefined): boolean => !NONE.test(value ?? '');
  const webkit = style as unknown as { webkitBackdropFilter?: string };
  return (
    set(style.transform) ||
    set(style.translate) ||
    set(style.scale) ||
    set(style.rotate) ||
    set(style.perspective) ||
    set(style.filter) ||
    set(style.backdropFilter) ||
    set(webkit.webkitBackdropFilter) ||
    set(style.transformStyle) ||
    style.contentVisibility === 'auto' ||
    CONTAIN.test(style.contain ?? '') ||
    set(style.containerType) ||
    WILL_CHANGE.test(style.willChange ?? '')
  );
}

/** The nearest element at or above `el` that holds fixed descendants, or null. */
function fixedHolder(el: Element | null, win: Window): Element | null {
  for (; el; el = el.parentElement) {
    if (holdsFixed(win.getComputedStyle(el))) return el;
    // A modal dialog or an open popover is in the top layer: no ancestor holds or clips it.
    if (inTopLayer(el)) return null;
  }
  return null;
}

/** Whether an element is in the top layer, where the page around it neither moves nor clips it. */
function inTopLayer(el: Element): boolean {
  try {
    return el.matches(':modal, :popover-open');
  } catch {
    // A browser without these selectors has no top layer of this kind.
    return false;
  }
}

/** Whether an overflow value clips (`hidden`, `clip`, `auto`, `scroll`); unset does not. */
const clips = (overflow: string | undefined): boolean => !!overflow && overflow !== 'visible';
const clipsAny = (style: CSSStyleDeclaration): boolean =>
  clips(style.overflowX) ||
  clips(style.overflowY) ||
  clips(style.overflow) ||
  // Paint containment clips to the box as overflow does.
  /paint|strict|content/.test(style.contain ?? '') ||
  style.contentVisibility === 'auto';

/** The whole window, at the origin. */
function windowView(win: Window | null): FixedView {
  return { left: 0, top: 0, width: win?.innerWidth ?? 0, height: win?.innerHeight ?? 0 };
}

/**
 * The window, cut to the boxes that clip a fixed panel held by `holder`: the holder and the
 * boxes up its containing-block chain. An ancestor outside that chain does not clip it: a fixed
 * holder escapes every ancestor up to the next holder, and an absolute one escapes static ones.
 */
function visibleView(holder: Element, win: Window): FixedView {
  let { left, top, width: right, height: bottom } = windowView(win);
  const doc = holder.ownerDocument;
  for (let el: Element | null = holder; el && el !== doc.documentElement;) {
    const style = win.getComputedStyle(el);
    // The body's overflow moves to the window when the root's is visible: it clips nothing.
    const toWindow = el === doc.body && !clipsAny(win.getComputedStyle(doc.documentElement));
    if (clipsAny(style) && !toWindow) {
      const box = el.getBoundingClientRect();
      left = Math.max(left, box.left);
      top = Math.max(top, box.top);
      right = Math.min(right, box.right);
      bottom = Math.min(bottom, box.bottom);
    }
    if (inTopLayer(el)) break;
    if (style.position === 'fixed') el = fixedHolder(el.parentElement, win);
    else if (style.position === 'absolute') {
      el = el.parentElement;
      while (
        el &&
        win.getComputedStyle(el).position === 'static' &&
        !holdsFixed(win.getComputedStyle(el))
      )
        el = el.parentElement;
    } else el = el.parentElement;
  }
  return { left, top, width: Math.max(0, right - left), height: Math.max(0, bottom - top) };
}

/**
 * Move a window rectangle or point into `view`'s coordinates, where placement rules work: each
 * horizontal value less `view.left`, each vertical value less `view.top`. @internal
 */
export function inFixedView<R extends object>(r: R, view: FixedView): R {
  const out: Record<string, unknown> = {};
  for (const key of ['left', 'right', 'x', 'width'] as const)
    if (key in r)
      out[key] = (r as Record<string, number>)[key]! - (key === 'width' ? 0 : view.left);
  for (const key of ['top', 'bottom', 'y', 'height'] as const)
    if (key in r)
      out[key] = (r as Record<string, number>)[key]! - (key === 'height' ? 0 : view.top);
  return out as R;
}

/**
 * Place a `position: fixed` child of `host`. `place` gets the panel's size and the area it can
 * show in, both in window pixels, and returns a place in that area's coordinates (move anchors
 * there with `inFixedView`). The result is the `left`, `top`, and `maxHeight` to write on the
 * panel. In a page with no host ancestor that holds fixed elements, the area is the window and
 * the place is returned unchanged. @internal
 */
export function placeInFixedFrame<T extends FixedPoint>(
  host: Element | null | undefined,
  size: FixedSize,
  place: (size: FixedSize, view: FixedView) => T
): T {
  const win = host?.ownerDocument?.defaultView ?? (typeof window === 'undefined' ? null : window);
  if (!host || !win) return place(size, windowView(win));
  // The measured frame finds a transform or scale; the styles find a holder at the window's
  // origin with no scale, which still clips (a drawer at the left edge).
  const frame = measureFrame(host);
  if (!frame) return place(size, windowView(win));
  const holder = fixedHolder(host, win);
  const moved = frame.x !== 0 || frame.y !== 0 || frame.sx !== 1 || frame.sy !== 1;
  if (!holder && !moved) return place(size, windowView(win));
  const view = holder ? visibleView(holder, win) : windowView(win);
  const at = place({ width: size.width * frame.sx, height: size.height * frame.sy }, view);
  return {
    ...at,
    left: (at.left + view.left - frame.x) / frame.sx,
    top: (at.top + view.top - frame.y) / frame.sy,
    ...(at.maxHeight === undefined ? {} : { maxHeight: at.maxHeight / frame.sy }),
  };
}
