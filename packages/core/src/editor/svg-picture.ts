// Whether an SVG picture part draws as vector art. The browser draws an SVG picture on the
// page's main thread, and most SVG features can cost seconds per picture: text, gradients,
// filters, masks, patterns, markers, blending, and animation. Every SVG picture of a
// presentation carries a raster alternative, so the editor draws an SVG picture as vector art
// only when it is a plain icon, and draws the raster alternative otherwise.
//
// An icon here: at most `MAX_SVG_BYTES`, a root `<svg>` with a finite `viewBox`, only the
// elements in `ALLOWED`, attributes from a fixed list with the value rules of
// svg-picture-css.ts, solid paint, links only from `<use>` and `clip-path` to elements of the
// icon, and simple style sheets. With each `<use>` expanded, also where it sits in `<defs>` or
// `<symbol>`, its shapes, path vertices, and groups with opacity, each as a share of
// `MAX_SVG_SHAPES`, `MAX_SVG_VERTICES`, and `MAX_SVG_LAYERS`, add up to at most 1. With the
// even-odd rule anywhere, the squares of each path's vertices also count, as a share of
// `MAX_SVG_EVEN_ODD`: a browser fills a path that crosses itself with the even-odd rule in
// time that grows with the square of its vertices. Each limit alone draws in about 250 ms in
// the browsers the editor supports. The check takes time linear in the part and needs no DOM.

import { readXml, type XmlNode } from '@eigenpal/ooxml-core';
import {
  declarations,
  isLength,
  isTransform,
  numbers,
  paintValue,
  styleSheet,
} from './svg-picture-css.ts';

/** The longest SVG picture part that can draw as vector art. */
export const MAX_SVG_BYTES = 256 * 1024;
/** The most shapes, path vertices, and groups with opacity an icon draws, each alone. */
export const MAX_SVG_SHAPES = 1_000;
export const MAX_SVG_VERTICES = 4_000;
export const MAX_SVG_LAYERS = 100;
/** The most of the sum of the squares of path vertices, with the even-odd rule. */
export const MAX_SVG_EVEN_ODD = 1_000_000;
/** The most elements, attributes on one element, and `<use>` links in a chain. */
const MAX_ELEMENTS = 10_000;
const MAX_ATTRIBUTES = 64;
const MAX_USE_CHAIN = 8;

const SVG_NS = 'http://www.w3.org/2000/svg';
const XLINK_NS = 'http://www.w3.org/1999/xlink';
const SHAPES: ReadonlySet<string> = new Set([
  'path',
  'rect',
  'circle',
  'ellipse',
  'line',
  'polyline',
  'polygon',
]);
const ALLOWED: ReadonlySet<string> = new Set([
  ...SHAPES,
  'svg',
  'g',
  'defs',
  'use',
  'symbol',
  'clipPath',
  'style',
  'title',
  'desc',
  'metadata',
]);
/** Elements whose content is not drawn and not read. */
const IGNORED: ReadonlySet<string> = new Set(['title', 'desc', 'metadata']);
/** Geometry attributes: finite lengths. */
const LENGTHS: ReadonlySet<string> = new Set([
  'x',
  'y',
  'width',
  'height',
  'cx',
  'cy',
  'r',
  'rx',
  'ry',
  'x1',
  'y1',
  'x2',
  'y2',
]);
/** Attributes that draw nothing, with any value. */
const INERT: ReadonlySet<string> = new Set(['id', 'class', 'version', 'role', 'focusable', 'type']);

type Element = Extract<XmlNode, { type: 'element' }>;

/**
 * What an element draws, with each `<use>` and clip path expanded: its shapes, vertices, and
 * layers as one share of the limits, and the squares of its paths' vertices.
 */
interface Count {
  readonly load: number;
  readonly squares: number;
}

interface Node {
  readonly name: string;
  readonly own: Count;
  readonly children: Node[];
  readonly use: string | undefined;
  readonly clips: readonly string[];
}

const PATH_CHARS = /[^\d\s.,+\-eEMmLlHhVvCcSsQqTtAaZz]/;
const TOKEN = /[MmLlHhVvCcSsQqTtAaZz]|[+-]?(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?/g;

/**
 * The vertices of path data: one for each coordinate pair, each number of `H` and `V`, each
 * arc segment (counted as four numbers, as its flags can pack), and each `Z`; or -1 for path
 * data with characters it cannot hold.
 */
function pathVertices(d: string): number {
  if (PATH_CHARS.test(d)) return -1;
  let total = 0;
  let per = 2;
  let run = 0;
  for (const [token] of d.matchAll(TOKEN)) {
    const c = token.charCodeAt(0);
    const letter = (c >= 0x41 && c <= 0x5a) || (c >= 0x61 && c <= 0x7a);
    if (!letter) {
      run += 1;
      continue;
    }
    total += Math.ceil(run / per);
    run = 0;
    const lower = token.toLowerCase();
    if (lower === 'z') total += 1;
    else per = lower === 'h' || lower === 'v' ? 1 : lower === 'a' ? 4 : 2;
  }
  return total + Math.ceil(run / per);
}

const results = new WeakMap<Uint8Array, boolean>();

/** Whether the SVG document `bytes` draws as vector art (see the module comment). */
export function svgDrawable(bytes: Uint8Array): boolean {
  const hit = results.get(bytes);
  if (hit !== undefined) return hit;
  const ok = check(bytes);
  results.set(bytes, ok);
  return ok;
}

/** The text without its document type declaration, or null for one with an internal subset. */
function withoutDoctype(text: string): string | null {
  const at = text.indexOf('<!DOCTYPE');
  if (at < 0) return text;
  const end = text.indexOf('>', at);
  if (end < 0 || text.slice(at, end).includes('[')) return null;
  if (text.indexOf('<!DOCTYPE', end) >= 0) return null;
  return text.slice(0, at) + text.slice(end + 1);
}

function check(bytes: Uint8Array): boolean {
  if (bytes.length > MAX_SVG_BYTES) return false;
  const decoded = new TextDecoder('utf-8', { fatal: false }).decode(bytes);
  if (decoded.includes('<?xml-stylesheet')) return false;
  const text = withoutDoctype(decoded);
  if (text === null) return false;
  const read = readXml(text, { maxBytes: bytes.length + 1, maxElements: MAX_ELEMENTS });
  if (!read.ok) return false;

  const ids = new Map<string, Node[]>();
  const sheet = { clips: [] as string[], opacity: false, evenOdd: false };
  let elements = 0;

  /** The namespaces an element declares, and its parent's: a chain, not a copy each. */
  type Scope = { readonly names: ReadonlyMap<string, string>; readonly parent: Scope | null };
  const resolve = (scope: Scope | null, prefix: string): string | undefined => {
    for (let s = scope; s; s = s.parent) {
      const ns = s.names.get(prefix);
      if (ns !== undefined) return ns;
    }
    return undefined;
  };

  /** Read `el` into a node, or null when it is not part of an icon. */
  const build = (el: Element, parent: Scope | null, root: boolean): Node | null => {
    elements += 1;
    const entries = Object.entries(el.attributes);
    if (entries.length > MAX_ATTRIBUTES) return null;
    let declared: Map<string, string> | undefined;
    for (const [name, value] of entries)
      if (name === 'xmlns' || name.startsWith('xmlns:'))
        (declared ??= new Map()).set(name === 'xmlns' ? '' : name.slice(6), value);
    const scope = declared ? { names: declared, parent } : parent;
    const colon = el.name.indexOf(':');
    const local = el.name.slice(colon + 1);
    if (resolve(scope, colon < 0 ? '' : el.name.slice(0, colon)) !== SVG_NS) return null;
    if (!ALLOWED.has(local) || (root && local !== 'svg')) return null;
    const empty = { load: 0, squares: 0 };
    if (IGNORED.has(local))
      return { name: local, own: empty, children: [], use: undefined, clips: [] };

    const found = { clips: [] as string[], opacity: false, evenOdd: false };
    let use: string | undefined;
    let vertices = 0;
    let box = false;
    for (const [name, value] of entries) {
      if (name === 'xmlns' || name.startsWith('xmlns:')) continue;
      const at = name.indexOf(':');
      if (at >= 0) {
        // `xlink:href` and `xml:space` are the only prefixed attributes.
        const ns = name.slice(0, at) === 'xml' ? 'xml' : resolve(scope, name.slice(0, at));
        const attr = name.slice(at + 1);
        if (ns === 'xml' && attr === 'space') continue;
        if (ns !== XLINK_NS || attr !== 'href') return null;
      }
      const attr = name.slice(at + 1);
      if (attr === 'href') {
        if (local !== 'use' || !value.trim().startsWith('#')) return null;
        use = value.trim().slice(1);
        continue;
      }
      if (INERT.has(attr) || attr.startsWith('data-') || attr.startsWith('aria-')) continue;
      if (root && (attr === 'width' || attr === 'height')) continue;
      if (LENGTHS.has(attr)) {
        if (!isLength(value)) return null;
        continue;
      }
      switch (attr) {
        case 'd': {
          if (local !== 'path') return null;
          const n = pathVertices(value);
          if (n < 0) return null;
          vertices += n;
          continue;
        }
        case 'points': {
          const list = numbers(value);
          if (!list) return null;
          vertices += Math.ceil(list.length / 2);
          continue;
        }
        case 'viewBox': {
          const v = numbers(value);
          if (!v || v.length !== 4 || !(v[2]! > 0) || !(v[3]! > 0)) return null;
          box = true;
          continue;
        }
        case 'preserveAspectRatio':
          if (/slice/i.test(value)) return null;
          continue;
        case 'transform':
          if (!isTransform(value)) return null;
          continue;
        case 'clipPathUnits':
          if (!/^(?:userSpaceOnUse|objectBoundingBox)$/.test(value.trim())) return null;
          continue;
        case 'style':
          if (!declarations(value, found)) return null;
          continue;
        default:
          if (paintValue(attr, value, found) !== true) return null;
      }
    }
    if (root && !box) return null;

    const children: Node[] = [];
    for (const c of el.children) {
      if (c.type === 'text') {
        if (local === 'style' && !styleSheet(c.value, sheet)) return null;
        continue;
      }
      const node = build(c, scope, false);
      if (!node) return null;
      children.push(node);
    }
    if (found.evenOdd) sheet.evenOdd = true;
    const own = {
      load:
        (SHAPES.has(local) ? 1 / MAX_SVG_SHAPES : 0) +
        vertices / MAX_SVG_VERTICES +
        (found.opacity ? 1 / MAX_SVG_LAYERS : 0),
      squares: vertices * vertices,
    };
    const node: Node = { name: local, own, children, use, clips: found.clips };
    const id = el.attributes.id;
    if (id !== undefined) {
      const same = ids.get(id);
      if (same) same.push(node);
      else ids.set(id, [node]);
    }
    return node;
  };

  const roots: Node[] = [];
  for (const n of read.nodes) {
    if (n.type !== 'element') continue;
    const node = build(n, null, true);
    if (!node || roots.length > 0) return false;
    roots.push(node);
  }
  if (roots.length !== 1) return false;

  // Each node and each id counted once. A `<use>` and a `clip-path` must link to an element of
  // the icon (a clip path for `clip-path`); a missing one, a cycle, or a long chain refuses.
  const counts = new Map<Node, Count>();
  const idCounts = new Map<string, Count>();
  const active = new Set<Node>();
  let chain = 0;
  let refused = false;
  const OVER: Count = { load: Infinity, squares: Infinity };
  const add = (a: Count, b: Count, times = 1): Count => ({
    load: a.load + times * b.load,
    squares: a.squares + times * b.squares,
  });
  const share = (c: Count): number => c.load + (sheet.evenOdd ? c.squares / MAX_SVG_EVEN_ODD : 0);
  const over = (c: Count): boolean => !(share(c) <= 1);
  const linked = (id: string, clip: boolean): Count => {
    const key = `${clip ? 'c' : 'u'}#${id}`;
    const known = idCounts.get(key);
    if (known) return known;
    const nodes = (ids.get(id) ?? []).filter((n) => !clip || n.name === 'clipPath');
    if (nodes.length === 0 || ++chain > MAX_USE_CHAIN) return ((refused = true), OVER);
    let largest: Count = { load: 0, squares: 0 };
    for (const n of nodes) {
      const c = count(n);
      if (share(c) > share(largest)) largest = c;
    }
    chain -= 1;
    idCounts.set(key, largest);
    return largest;
  };
  const count = (node: Node): Count => {
    const known = counts.get(node);
    if (known) return known;
    if (active.has(node) || refused) return ((refused = true), OVER);
    active.add(node);
    let total = node.own;
    for (const child of node.children) {
      total = add(total, count(child));
      if (over(total)) break;
    }
    if (node.use !== undefined && !over(total)) total = add(total, linked(node.use, false));
    for (const id of node.clips) if (!over(total)) total = add(total, linked(id, true));
    active.delete(node);
    counts.set(node, total);
    return total;
  };
  let total = count(roots[0]!);
  // A sheet's clip paths and opacity can reach every element.
  for (const id of sheet.clips) total = add(total, linked(id, true), elements);
  if (sheet.opacity) total = add(total, { load: elements / MAX_SVG_LAYERS, squares: 0 });
  return !refused && !over(total);
}
