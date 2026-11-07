// Whether an SVG picture part is safe for the browser to draw. The browser draws an SVG
// picture on the page's main thread, and some elements multiply the work of others: `<use>`
// copies its target, so ten levels of ten copies each draw 10^10 shapes, and a marker draws
// its content at each vertex of a path. This check reads the part with the bounded XML reader
// and counts the elements that the picture draws, with every `<use>` expanded and every marker
// drawn at every vertex of every path, line, polyline, and polygon. A part over the limit draws
// the picture's raster alternative instead. So does a part that the reader refuses (an internal
// DTD subset, too deep, too many elements), one with a `<use>` cycle, one with a marker inside
// marker content, and one with a `data:` URI that is not a raster picture: a nested SVG
// document there escapes this count. The check needs no DOM.

import { readXml, type XmlNode } from '@eigenpal/ooxml-core';

/**
 * The most elements an SVG picture may draw, with each `<use>` expanded and each marker drawn
 * at each vertex. Icons and diagrams draw a few thousand at most.
 */
export const MAX_SVG_DRAWN_ELEMENTS = 200_000;
/** The most `<use>` links that an expansion may follow in a chain. */
const MAX_USE_CHAIN = 16;

/** Elements whose content draws only where a `<use>` or a marker property links to it. */
const UNDRAWN: ReadonlySet<string> = new Set(['defs', 'symbol', 'marker']);
/** The elements that draw markers (SVG 2, 11.6.2 "markable elements"). */
const MARKABLE: ReadonlySet<string> = new Set(['path', 'line', 'polyline', 'polygon']);

type Element = Extract<XmlNode, { type: 'element' }>;

const localName = (name: string): string => name.slice(name.indexOf(':') + 1);

/** The `#id` that a `<use>` links to, from `href` or any prefixed `href` (`xlink:href`). */
function linkTarget(el: Element): string | undefined {
  for (const [name, value] of Object.entries(el.attributes))
    if (localName(name) === 'href' && value.trim().startsWith('#')) return value.trim().slice(1);
  return undefined;
}

/** A document type declaration without an internal subset: it declares nothing to expand. */
const PLAIN_DOCTYPE = /<!DOCTYPE[^[>]*>/;

/** CSS escapes (`\64` or `\:`): a style sheet can spell `data:` with them. */
const CSS_ESCAPE = /\\(?:([0-9a-f]{1,6})\s?|(.))/gi;
/** A `data:` URI of anything but a raster picture, in a value with no space, quote, or escape. */
const NESTED_DATA = /data:(?!image\/(?:png|jpe?g|gif|webp|bmp|avif)[;,])/;

/** Whether an attribute value or a style sheet holds a `data:` URI of a document. */
function nestedData(value: string): boolean {
  if (!/data/i.test(value) && !value.includes('\\')) return false;
  const plain = value
    .replace(CSS_ESCAPE, (_, hex: string | undefined, ch: string | undefined) =>
      hex ? String.fromCodePoint(Math.min(parseInt(hex, 16), 0x10ffff)) : (ch ?? '')
    )
    .toLowerCase()
    .replace(/[\s"']/g, '');
  return NESTED_DATA.test(plain);
}

/** A marker property: a presentation attribute, or a declaration in a style. */
const MARKER_ATTRIBUTES = ['marker-start', 'marker-mid', 'marker-end'];
const MARKER_DECLARATION = /marker(?:-start|-mid|-end)?\s*:([^;]*)/gi;
const URL_ID = /url\(\s*["']?#([^"')\s]+)/g;

/** The ids of the markers that `el` names in its own marker properties. */
function markerIds(el: Element): string[] {
  const values: string[] = [];
  for (const name of MARKER_ATTRIBUTES) {
    const v = el.attributes[name];
    if (v !== undefined) values.push(v);
  }
  const style = el.attributes.style;
  if (style) for (const m of style.matchAll(MARKER_DECLARATION)) values.push(m[1]!);
  return values.flatMap((v) => [...v.matchAll(URL_ID)].map((m) => m[1]!));
}

const NUMBER = /[+-]?(?:\d+\.?\d*|\.\d+)(?:e[+-]?\d+)?/gi;

/** How many numbers `text` holds, with no array of them. */
function countNumbers(text: string): number {
  const re = new RegExp(NUMBER.source, NUMBER.flags);
  let n = 0;
  while (re.exec(text)) n += 1;
  return n;
}

/** At most how many vertices a markable element has: each one can draw a marker. */
function vertices(el: Element): number {
  const kind = localName(el.name);
  if (kind === 'line') return 2;
  if (kind === 'path') {
    // Each segment ends with at least one number, and `z` closes with one more vertex.
    const d = el.attributes.d ?? '';
    let close = 0;
    for (let i = 0; i < d.length; i += 1) if ((d.charCodeAt(i) | 0x20) === 0x7a) close += 1;
    return countNumbers(d) + close + 1;
  }
  return Math.ceil(countNumbers(el.attributes.points ?? '') / 2) + 1;
}

const results = new WeakMap<Uint8Array, boolean>();

/** Whether the SVG document `bytes` draws within the limits (see the module comment). */
export function svgDrawable(bytes: Uint8Array): boolean {
  const hit = results.get(bytes);
  if (hit !== undefined) return hit;
  const ok = check(bytes);
  results.set(bytes, ok);
  return ok;
}

function check(bytes: Uint8Array): boolean {
  let text: string;
  try {
    text = new TextDecoder('utf-8', { fatal: false }).decode(bytes);
  } catch {
    return false;
  }
  const read = readXml(text.replace(PLAIN_DOCTYPE, ''), {
    maxBytes: bytes.length + 1,
    maxElements: MAX_SVG_DRAWN_ELEMENTS,
  });
  if (!read.ok) return false;
  const ids = new Map<string, Element>();
  const roots: Element[] = [];
  /** The ids that a marker property names. */
  const named = new Set<string>();
  /** Markable elements in marker content: they draw no marker of their own. */
  const inMarker = new Set<Element>();
  let sheets = '';
  // The tree is at most `MAX_XML_DEPTH` deep: a walk on an explicit stack. Each entry says
  // whether it is in marker content, and whether a marker property reaches it from above.
  type Entry = { el: Element; marker: boolean; inherited: boolean };
  const stack: Entry[] = [];
  for (const n of read.nodes)
    if (n.type === 'element')
      (roots.push(n), stack.push({ el: n, marker: false, inherited: false }));
  const pending: Element[] = [];
  while (stack.length) {
    const { el, marker, inherited } = stack.pop()!;
    const id = el.attributes.id;
    if (id !== undefined && !ids.has(id)) ids.set(id, el);
    for (const value of Object.values(el.attributes)) if (nestedData(value)) return false;
    const own = markerIds(el);
    for (const m of own) named.add(m);
    const kind = localName(el.name);
    const reached = inherited || own.length > 0;
    const isMarker = marker || kind === 'marker';
    if (isMarker && MARKABLE.has(kind)) {
      // A marker inside marker content multiplies again: draw the raster alternative.
      if (reached) return false;
      inMarker.add(el);
      pending.push(el);
    }
    if (isMarker && kind === 'use' && reached) return false;
    if (kind === 'style')
      for (const c of el.children) if (c.type === 'text') sheets += `${c.value}\n`;
    for (const c of el.children)
      if (c.type === 'element') stack.push({ el: c, marker: isMarker, inherited: reached });
  }
  if (nestedData(sheets)) return false;
  // A style sheet can set a marker on any element, also in marker content.
  const sheetMarkers = [...sheets.matchAll(MARKER_DECLARATION)];
  if (sheetMarkers.length > 0 && pending.length > 0) return false;
  for (const m of sheetMarkers) for (const u of m[1]!.matchAll(URL_ID)) named.add(u[1]!);
  const markers = [...named]
    .map((m) => ids.get(m))
    .filter((el): el is Element => !!el && localName(el.name) === 'marker');

  // The drawn size of each element: itself, its children, the target of its `<use>`, and, for
  // a markable element, the largest marker at each vertex. Any marker that the picture names
  // counts at each vertex of each markable element outside marker content: an upper bound.
  const drawn = new Map<Element, number>();
  const active = new Set<Element>();
  let chain = 0;
  let largest: number | undefined;
  let measuring = false;
  const markerSize = (): number => {
    if (largest !== undefined) return largest;
    if (measuring) return Infinity; // marker content that draws a marker
    measuring = true;
    let max = 0;
    for (const m of markers) max = Math.max(max, size(m));
    measuring = false;
    return (largest = max);
  };
  const size = (el: Element): number => {
    const known = drawn.get(el);
    if (known !== undefined) return known;
    if (active.has(el)) return Infinity; // a cycle
    active.add(el);
    let total = 1;
    for (const c of el.children) {
      if (c.type !== 'element') continue;
      // Definitions draw only through a link, which counts them there.
      total += UNDRAWN.has(localName(c.name)) ? 1 : size(c);
      if (total > MAX_SVG_DRAWN_ELEMENTS) break;
    }
    const kind = localName(el.name);
    if (markers.length > 0 && MARKABLE.has(kind) && !inMarker.has(el))
      if (total <= MAX_SVG_DRAWN_ELEMENTS) total += vertices(el) * markerSize();
    const target = kind === 'use' ? linkTarget(el) : undefined;
    const linked = target === undefined ? undefined : ids.get(target);
    if (linked && total <= MAX_SVG_DRAWN_ELEMENTS) {
      chain += 1;
      total += chain > MAX_USE_CHAIN ? Infinity : size(linked);
      chain -= 1;
    }
    active.delete(el);
    drawn.set(el, total);
    return total;
  };
  let total = 0;
  for (const root of roots) total += size(root);
  return total <= MAX_SVG_DRAWN_ELEMENTS;
}
