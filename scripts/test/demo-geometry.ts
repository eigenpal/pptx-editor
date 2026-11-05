// Layout for a demo in happy-dom, which lays nothing out: the boxes that a browser would give
// the slide stage, the canvas, the rail items, and the painted SVG, and the client points of
// slide points and painted words. The stage gets the size that the surface sets on it
// (`style.width` and `style.height`); an SVG element gets the box of its attributes under its
// ancestors' transforms, so a test reads painted positions as `getBoundingClientRect` gives them
// in a browser (text boxes are approximate: from the glyph positions and the font size).

/** Slide size of the basic deck, in EMU (16:9). */
export const SLIDE_W = 12_192_000;
export const SLIDE_H = 6_858_000;
export const IN = 914_400;
/** Where the stage's top-left corner sits on the page. */
const STAGE_AT = { x: 300, y: 200 };

const realRect = Element.prototype.getBoundingClientRect;

function px(v: string): number {
  return Number.parseFloat(v) || 0;
}

/** A DOMRect from a box. */
export function rect(x: number, y: number, width: number, height: number): DOMRect {
  return {
    x,
    y,
    left: x,
    top: y,
    width,
    height,
    right: x + width,
    bottom: y + height,
    toJSON: () => ({}),
  } as DOMRect;
}

/** The box that the browser gives the stage: its styled size at `STAGE_AT`. */
export function stageRect(el: HTMLElement): DOMRect {
  return rect(STAGE_AT.x, STAGE_AT.y, px(el.style.width), px(el.style.height));
}

/** The canvas size at a 1440 × 900 window, with the rail and the side panel beside it. */
const CANVAS = { width: 880, height: 560 };

function findGetter(
  name: 'clientWidth' | 'clientHeight' | 'offsetWidth'
): (this: Element) => number {
  for (let p: object | null = HTMLElement.prototype; p; p = Object.getPrototypeOf(p)) {
    const d = Object.getOwnPropertyDescriptor(p, name);
    if (d?.get) return d.get as (this: Element) => number;
  }
  return () => 0;
}
const realClientWidth = findGetter('clientWidth');
const realClientHeight = findGetter('clientHeight');
const realOffsetWidth = findGetter('offsetWidth');
const OWN_CLIENT_SIZE = (['clientWidth', 'clientHeight', 'offsetWidth'] as const).map(
  (name) => [name, Object.getOwnPropertyDescriptor(HTMLElement.prototype, name)] as const
);

type Matrix = readonly [number, number, number, number, number, number];
const IDENTITY: Matrix = [1, 0, 0, 1, 0, 0];

function multiply(m: Matrix, n: Matrix): Matrix {
  return [
    m[0] * n[0] + m[2] * n[1],
    m[1] * n[0] + m[3] * n[1],
    m[0] * n[2] + m[2] * n[3],
    m[1] * n[2] + m[3] * n[3],
    m[0] * n[4] + m[2] * n[5] + m[4],
    m[1] * n[4] + m[3] * n[5] + m[5],
  ];
}

/** An SVG `transform` attribute as one matrix. */
function parseTransform(value: string | null): Matrix {
  let m = IDENTITY;
  for (const [, fn, args] of (value ?? '').matchAll(/(\w+)\(([^)]*)\)/g)) {
    const a = args
      .split(/[\s,]+/)
      .filter(Boolean)
      .map(Number);
    if (fn === 'matrix') m = multiply(m, a as unknown as Matrix);
    else if (fn === 'translate') m = multiply(m, [1, 0, 0, 1, a[0], a[1] ?? 0]);
    else if (fn === 'scale') m = multiply(m, [a[0], 0, 0, a[1] ?? a[0], 0, 0]);
    else if (fn === 'rotate') {
      const r = (a[0] * Math.PI) / 180;
      const [cx, cy] = [a[1] ?? 0, a[2] ?? 0];
      m = multiply(m, [1, 0, 0, 1, cx, cy]);
      m = multiply(m, [Math.cos(r), Math.sin(r), -Math.sin(r), Math.cos(r), 0, 0]);
      m = multiply(m, [1, 0, 0, 1, -cx, -cy]);
    }
  }
  return m;
}

export const nums = (el: Element, name: string): number[] =>
  (el.getAttribute(name) ?? '')
    .split(/[\s,]+/)
    .filter(Boolean)
    .map(Number);

/** An SVG element's own box in its user space: [x0, y0, x1, y1], or null. */
function ownBox(el: Element): [number, number, number, number] | null {
  const tag = el.localName;
  if (tag === 'text') {
    // Each glyph has its own x; the run ends about one glyph advance after the last.
    const xs = nums(el, 'x');
    const y = nums(el, 'y')[0] ?? 0;
    const size = Number(el.getAttribute('font-size') ?? 0);
    if (!xs.length) return null;
    const step = xs.length > 1 ? (xs[xs.length - 1] - xs[0]) / (xs.length - 1) : size / 2;
    return [xs[0], y - size * 0.8, xs[xs.length - 1] + step, y + size * 0.2];
  }
  if (tag === 'rect' || tag === 'image' || tag === 'foreignObject') {
    const [x, y] = [Number(el.getAttribute('x') ?? 0), Number(el.getAttribute('y') ?? 0)];
    return [
      x,
      y,
      x + Number(el.getAttribute('width') ?? 0),
      y + Number(el.getAttribute('height') ?? 0),
    ];
  }
  if (tag === 'line') {
    const [x1, y1, x2, y2] = ['x1', 'y1', 'x2', 'y2'].map((a) => Number(el.getAttribute(a) ?? 0));
    return [Math.min(x1, x2), Math.min(y1, y2), Math.max(x1, x2), Math.max(y1, y2)];
  }
  if (tag === 'circle' || tag === 'ellipse') {
    const [cx, cy] = [Number(el.getAttribute('cx') ?? 0), Number(el.getAttribute('cy') ?? 0)];
    const rx = Number(el.getAttribute('rx') ?? el.getAttribute('r') ?? 0);
    const ry = Number(el.getAttribute('ry') ?? el.getAttribute('r') ?? 0);
    return [cx - rx, cy - ry, cx + rx, cy + ry];
  }
  return null;
}

/** The union of `el`'s box and its children's, in the coordinates of `el`'s parent. */
function boxIn(el: Element, outer: Matrix): [number, number, number, number] | null {
  const m = multiply(outer, parseTransform(el.getAttribute('transform')));
  const boxes: [number, number, number, number][] = [];
  const own = ownBox(el);
  if (own) {
    const corners = [
      [own[0], own[1]],
      [own[2], own[1]],
      [own[0], own[3]],
      [own[2], own[3]],
    ].map(([x, y]) => [m[0] * x + m[2] * y + m[4], m[1] * x + m[3] * y + m[5]]);
    const xs = corners.map((c) => c[0]);
    const ys = corners.map((c) => c[1]);
    boxes.push([Math.min(...xs), Math.min(...ys), Math.max(...xs), Math.max(...ys)]);
  }
  if (el.localName !== 'text')
    for (const child of el.children) {
      if (child.localName === 'defs' || child.localName === 'clipPath') continue;
      const b = boxIn(child, m);
      if (b) boxes.push(b);
    }
  if (!boxes.length) return null;
  return [
    Math.min(...boxes.map((b) => b[0])),
    Math.min(...boxes.map((b) => b[1])),
    Math.max(...boxes.map((b) => b[2])),
    Math.max(...boxes.map((b) => b[3])),
  ];
}

/**
 * The page box of an SVG element on the stage, from its attributes and the transforms of its
 * ancestors: what a browser's `getBoundingClientRect` gives it (text boxes are approximate).
 */
function svgRect(el: SVGElement): DOMRect | null {
  const root = el.closest('.pptx-surface__stage > svg');
  if (!root) return null;
  let m = IDENTITY;
  for (let n = el.parentElement; n && n !== root; n = n.parentElement)
    m = multiply(parseTransform(n.getAttribute('transform')), m);
  const box = el === root ? boxIn(el, IDENTITY) : boxIn(el, m);
  const view = nums(root, 'viewBox');
  const st = stageRect(root.parentElement as HTMLElement);
  if (!box || view.length !== 4) return rect(st.x, st.y, 0, 0);
  const sx = st.width / view[2];
  const sy = st.height / view[3];
  return rect(
    st.x + box[0] * sx,
    st.y + box[1] * sy,
    (box[2] - box[0]) * sx,
    (box[3] - box[1]) * sy
  );
}

/**
 * Give the canvas and, with `barWidth`, the toolbar a size. A toolbar measured this way is the
 * toolbar tests' model (`toolbar-parts.test.tsx`): every control is 9 pixels wide, and the bar
 * width decides how many groups fit before the rest move to More.
 */
export function installGeometry(barWidth: number | undefined): void {
  const inBar = (el: HTMLElement): boolean =>
    barWidth !== undefined && !!el.closest('.pptx-toolbar');
  // The canvas has room for the slide, so the fit zoom is the one a 1440 px window gives.
  Object.defineProperty(HTMLElement.prototype, 'clientWidth', {
    configurable: true,
    get(this: HTMLElement) {
      if (this.classList.contains('pptx-canvas')) return CANVAS.width;
      if (
        barWidth !== undefined &&
        (this.classList.contains('pptx-toolbar') || this.classList.contains('demo-dock'))
      )
        return barWidth;
      return realClientWidth.call(this);
    },
  });
  Object.defineProperty(HTMLElement.prototype, 'clientHeight', {
    configurable: true,
    get(this: HTMLElement) {
      return this.classList.contains('pptx-canvas') ? CANVAS.height : realClientHeight.call(this);
    },
  });
  Object.defineProperty(HTMLElement.prototype, 'offsetWidth', {
    configurable: true,
    get(this: HTMLElement) {
      return inBar(this) && !this.classList.contains('pptx-toolbar')
        ? 9
        : realOffsetWidth.call(this);
    },
  });
  Element.prototype.getBoundingClientRect = function (this: Element): DOMRect {
    if (this instanceof HTMLElement && this.classList.contains('pptx-surface__stage'))
      return stageRect(this);
    if (this instanceof SVGElement) {
      const box = svgRect(this);
      if (box) return box;
    }
    if (this instanceof HTMLElement && this.classList.contains('pptx-canvas'))
      return rect(STAGE_AT.x - 40, STAGE_AT.y - 40, 1360, 800);
    // A rail item: 120 px high, one under the other, so a drop point is above or below its
    // middle.
    if (this instanceof HTMLElement && this.classList.contains('pptx-rail__item')) {
      const items = [...(this.parentElement?.querySelectorAll('.pptx-rail__item') ?? [])];
      const i = Math.max(0, items.indexOf(this));
      return rect(20, 100 + i * 130, 200, 120);
    }
    return realRect.call(this);
  };
}

/** Put back the browser's own measurements. */
export function restoreGeometry(): void {
  Element.prototype.getBoundingClientRect = realRect;
  for (const [name, own] of OWN_CLIENT_SIZE) {
    if (own) Object.defineProperty(HTMLElement.prototype, name, own);
    else Reflect.deleteProperty(HTMLElement.prototype, name);
  }
}

/** The client point of a slide point (EMU), on the stage. */
export function slidePoint(x: number, y: number): { x: number; y: number } {
  const stage = document.querySelector<HTMLElement>('.pptx-surface__stage')!;
  const box = stageRect(stage);
  return { x: box.x + (x / SLIDE_W) * box.width, y: box.y + (y / SLIDE_H) * box.height };
}

/** The slide stage of the canvas. */
export function stage(): HTMLElement {
  return document.querySelector<HTMLElement>('.pptx-surface__stage')!;
}

/**
 * The client point of the middle of a painted word, as a DOM range over its letters gives it
 * in a browser (from the second letter to the one before the last).
 */
export function wordPoint(word: string): { x: number; y: number } {
  for (const t of document.querySelectorAll<SVGTextElement>('.pptx-surface__slide text')) {
    const at = (t.textContent ?? '').indexOf(word);
    if (at < 0) continue;
    const xs = nums(t, 'x');
    const size = Number(t.getAttribute('font-size') ?? 0);
    const x = (xs[at + 1] + xs[at + word.length - 1]) / 2;
    const y = (nums(t, 'y')[0] ?? 0) - size * 0.3;
    return svgPoint(t, x, y);
  }
  throw new Error(`no painted word ${word}`);
}

/** The client point of a point in an SVG element's own coordinates. */
function svgPoint(el: Element, x: number, y: number): { x: number; y: number } {
  const root = el.closest('.pptx-surface__stage > svg')!;
  let m = parseTransform(el.getAttribute('transform'));
  for (let n = el.parentElement; n && n !== root; n = n.parentElement)
    m = multiply(parseTransform(n.getAttribute('transform')), m);
  const view = nums(root, 'viewBox');
  const st = stageRect(root.parentElement as HTMLElement);
  const sx = m[0] * x + m[2] * y + m[4];
  const sy = m[1] * x + m[3] * y + m[5];
  return { x: st.x + (sx * st.width) / view[2], y: st.y + (sy * st.height) / view[3] };
}
