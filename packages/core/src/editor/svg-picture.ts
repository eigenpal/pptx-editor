// Whether an SVG picture part draws as vector art. The browser draws an SVG picture on the
// page's main thread, and many SVG features cost per pixel, per vertex, or per layer, not per
// element: filters, masks, patterns, markers, animation, and links that copy content. Every
// SVG picture of a presentation carries a raster alternative, so the editor draws an SVG
// picture only when it stays inside a small subset whose cost this check can bound, and draws
// the raster alternative otherwise.
//
// The subset: the elements in `ALLOWED`, in the SVG namespace; `<image>` with a raster `data:`
// URI of at most 40 megapixels; links only to ids in the same document (`<use>`, and
// `url(#id)` to a gradient or a clip path); style sheets with no escape, custom property, or
// at-rule other than `@font-face` and `@media`; and no filter, mask, marker, dash array, or a
// path that a style sets. Its cost (`MAX_SVG_COST`) counts each element, path vertex, text
// character, layer, picture pixel, and the share of the picture that each shape covers, with
// each `<use>` expanded where it sits, also in `<defs>` and `<symbol>`, where browsers build
// it too. The check reads at most `MAX_SVG_BYTES`, in time and memory linear in the part. It
// needs no DOM.

import { imageSize, readXml, type XmlNode } from '@eigenpal/ooxml-core';
import { fromBase64 } from '../automation/commands/base64.ts';
import {
  foreignLink,
  fragmentId,
  largestLength,
  rasterData,
  scalingTransform,
  scanCss,
  squeeze,
  type CssScan,
} from './svg-picture-css.ts';

/** The longest SVG picture part this check reads. A longer one draws its raster alternative. */
export const MAX_SVG_BYTES = 2 * 1024 * 1024;
/**
 * The most an SVG picture may cost. One unit is about 0.4 microseconds of drawing on a
 * laptop, so the budget keeps a picture under about 200 ms.
 */
export const MAX_SVG_COST = 500_000;
/** The cost of a shape that covers the whole picture. */
const AREA_COST = 280;
/** The cost of an element drawn with opacity: a layer of its own. */
const LAYER_COST = 400;
const PIXELS_PER_UNIT = 200;
const MAX_IMAGE_PIXELS = 40_000_000;
/** The most `<use>` links that an expansion may follow in a chain. */
const MAX_USE_CHAIN = 16;

const SVG_NS = 'http://www.w3.org/2000/svg';
/** The SVG elements of the subset. `title`, `desc`, and `metadata` draw nothing. */
const ALLOWED: ReadonlySet<string> = new Set([
  'svg',
  'g',
  'defs',
  'title',
  'desc',
  'metadata',
  'path',
  'rect',
  'circle',
  'ellipse',
  'line',
  'polyline',
  'polygon',
  'text',
  'tspan',
  'linearGradient',
  'radialGradient',
  'stop',
  'clipPath',
  'use',
  'symbol',
  'style',
  'image',
]);
/** Attributes that pull in what the subset leaves out (by local name). */
const REFUSED_ATTRIBUTES = /^(?:filter|mask|marker(?:-start|-mid|-end)?|stroke-dasharray)$/;

type Element = Extract<XmlNode, { type: 'element' }>;

/** One element of the subset, with what its cost needs. */
interface Node {
  readonly name: string;
  readonly children: Node[];
  /** Its own cost, less its area: itself, its vertices or characters, its layer, its picture. */
  readonly own: number;
  /** Its extent in user units, before the stroke, when it is a shape. */
  readonly extent: { readonly w: number; readonly h: number } | undefined;
  /** For text, how many characters it draws. */
  readonly glyphs: number;
  /** The id that its `<use>` links to. */
  readonly use: string | undefined;
  /** The ids that its `url(#id)` values name: a clip path among them costs its content. */
  readonly links: readonly string[];
}

// --- Path data -----------------------------------------------------------------------------

const TOKEN = /[a-df-z]|[+-]?(?:\d+\.?\d*|\.\d+)(?:e[+-]?\d+)?/gi;
/** The numbers of one segment of each path command. An arc counts 4: its flags can pack. */
const SEGMENT: Readonly<Record<string, number>> = {
  m: 2,
  l: 2,
  t: 2,
  h: 1,
  v: 1,
  c: 6,
  s: 4,
  q: 4,
  a: 4,
};

/**
 * At most how many vertices path data (or a point list, `points`) has, and a bound on its
 * extent: no point lies farther from the origin than the sum of all its numbers, and an arc
 * bulges at most twice that.
 */
function pathData(text: string, points: boolean): { vertices: number; reach: number } {
  let total = 0;
  let per = 2;
  let run = 0;
  let reach = 0;
  for (const m of text.matchAll(TOKEN)) {
    const token = m[0];
    const letter = token.length === 1 ? token.toLowerCase() : '';
    if (points || !(letter >= 'a' && letter <= 'z')) {
      run += 1;
      reach += Math.abs(Number(token)) || 0;
      continue;
    }
    total += Math.ceil(run / per);
    run = 0;
    if (letter === 'z') total += 1;
    else per = SEGMENT[letter] ?? 1;
  }
  return { vertices: total + Math.ceil(run / per) + 1, reach: 2 * reach };
}

/** A length attribute in user units; a percent or a relative unit reads as unbounded. */
const length = (value: string | undefined): number =>
  value === undefined ? 0 : largestLength(value);

/** The extent of a shape from its attributes, before the stroke. */
function shapeExtent(
  local: string,
  a: Readonly<Record<string, string>>
): { w: number; h: number } | undefined {
  switch (local) {
    case 'rect':
    case 'image':
      return { w: length(a.width), h: length(a.height) };
    case 'circle':
      return { w: 2 * length(a.r), h: 2 * length(a.r) };
    case 'ellipse':
      return { w: 2 * length(a.rx), h: 2 * length(a.ry) };
    case 'line':
      return { w: length(a.x1) + length(a.x2), h: length(a.y1) + length(a.y2) };
    case 'path':
    case 'polyline':
    case 'polygon': {
      const data = local === 'path' ? (a.d ?? '') : (a.points ?? '');
      const { reach } = pathData(data, local !== 'path');
      return { w: reach, h: reach };
    }
    default:
      return undefined;
  }
}

// --- The check -----------------------------------------------------------------------------

const results = new WeakMap<Uint8Array, boolean>();

/** Whether the SVG document `bytes` draws as vector art (see the module comment). */
export function svgDrawable(bytes: Uint8Array): boolean {
  const hit = results.get(bytes);
  if (hit !== undefined) return hit;
  const ok = check(bytes);
  results.set(bytes, ok);
  return ok;
}

/** A document type declaration without an internal subset: it declares nothing to expand. */
const PLAIN_DOCTYPE = /<!DOCTYPE[^[>]*>/;

function check(bytes: Uint8Array): boolean {
  if (bytes.length > MAX_SVG_BYTES) return false;
  const text = new TextDecoder('utf-8', { fatal: false }).decode(bytes);
  // A style sheet from outside the tree: the subset has none.
  if (text.includes('<?xml-stylesheet')) return false;
  const read = readXml(text.replace(PLAIN_DOCTYPE, ''), {
    maxBytes: bytes.length + 1,
    maxElements: MAX_SVG_COST,
  });
  if (!read.ok) return false;
  const ids = new Map<string, Node[]>();
  let sheet = '';
  let elements = 0;
  // What the whole picture can do to the size of a shape: scale it, stroke it, set its font.
  let scales = false;
  let stroke = 1;
  let font = 16;
  let view: { w: number; h: number } | undefined;
  const styled = (scan: CssScan): void => {
    scales ||= scan.scales;
    stroke = Math.max(stroke, scan.stroke);
    font = Math.max(font, scan.font);
  };

  /** Read the attributes of an SVG element into `node`, or answer false to leave the subset. */
  const readAttributes = (
    local: string,
    a: Readonly<Record<string, string>>,
    node: { own: number; use: string | undefined; links: string[] }
  ): boolean => {
    for (const [name, value] of Object.entries(a)) {
      if (name === 'xmlns' || name.startsWith('xmlns:')) continue;
      const attr = name.slice(name.indexOf(':') + 1);
      if (REFUSED_ATTRIBUTES.test(attr)) return false;
      if (attr === 'href') {
        const plain = squeeze(value);
        if (local === 'use' && plain.startsWith('#')) node.use = fragmentId(plain.slice(1));
        else if (local === 'image') {
          const data = rasterData(value);
          const bytes = data === undefined ? null : fromBase64(data);
          const size = bytes ? imageSize(bytes) : null;
          if (!size || size.width * size.height > MAX_IMAGE_PIXELS) return false;
          node.own += Math.ceil((size.width * size.height) / PIXELS_PER_UNIT);
        } else if (!plain.startsWith('#')) return false;
        continue;
      }
      let scan: CssScan | null | undefined;
      if (attr === 'style') scan = scanCss(value, false);
      else if (/url\(/i.test(value)) scan = scanCss(`${attr}:${value}`, false);
      if (scan === null || foreignLink(value)) return false;
      if (scan) {
        node.links.push(...scan.ids);
        styled(scan);
        if (scan.opacity) node.own += LAYER_COST;
      }
      if (attr === 'opacity' && !(Number(value) >= 1)) node.own += LAYER_COST;
      if (attr === 'transform' && scalingTransform(value)) scales = true;
      if (attr === 'stroke-width') stroke = Math.max(stroke, length(value));
      if (attr === 'font-size') font = Math.max(font, length(value));
    }
    return true;
  };

  /** Read `el` in the namespaces `scope` into a node, or null when it leaves the subset. */
  const build = (el: Element, scope: ReadonlyMap<string, string>, root: boolean): Node | null => {
    elements += 1;
    let inner: Map<string, string> | undefined;
    for (const [name, value] of Object.entries(el.attributes))
      if (name === 'xmlns' || name.startsWith('xmlns:'))
        (inner ??= new Map(scope)).set(name === 'xmlns' ? '' : name.slice(6), value);
    const names = inner ?? scope;
    const colon = el.name.indexOf(':');
    const svg = names.get(colon < 0 ? '' : el.name.slice(0, colon)) === SVG_NS;
    const local = el.name.slice(colon + 1);
    const a = el.attributes;
    const read = { own: 1, use: undefined as string | undefined, links: [] as string[] };
    // An element in another namespace draws nothing, but its SVG content can be linked to.
    if (svg) {
      if (!ALLOWED.has(local) || !readAttributes(local, a, read)) return null;
      if (local === 'image' && read.own === 1) return null; // no picture
      // A viewBox scales its content; the root's sets the picture's own size.
      if (a.viewBox !== undefined && !root) scales = true;
      if (root) {
        const box = a.viewBox
          ?.trim()
          .split(/[\s,]+/)
          .map(Number);
        view =
          box?.length === 4
            ? { w: box[2]!, h: box[3]! }
            : { w: length(a.width) || 300, h: length(a.height) || 150 };
      }
      if (local === 'path' || local === 'polyline' || local === 'polygon') {
        const data = local === 'path' ? (a.d ?? '') : (a.points ?? '');
        read.own += pathData(data, local !== 'path').vertices;
      }
    }
    const children: Node[] = [];
    let glyphs = 0;
    for (const c of el.children) {
      if (c.type === 'text') {
        if (!svg) continue;
        if (local === 'style') sheet += c.value;
        else if (local === 'text' || local === 'tspan') glyphs += c.value.length;
        continue;
      }
      const node = build(c, names, false);
      if (!node) return null;
      children.push(node);
    }
    const node: Node = {
      name: svg ? local : '',
      children,
      own: read.own + glyphs,
      extent: svg ? shapeExtent(local, a) : undefined,
      glyphs,
      use: read.use,
      links: read.links,
    };
    const id = a.id;
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
    const node = build(n, new Map(), true);
    if (!node) return false;
    roots.push(node);
  }
  // The sheets join their text and CDATA pieces with nothing between, as the browser does.
  const sheetScan = sheet ? scanCss(sheet, true) : undefined;
  if (sheetScan === null) return false;
  if (sheetScan) styled(sheetScan);

  // The share of the picture that a shape covers, as a cost: all of it when anything can
  // scale shapes, or when the picture has no size. Each glyph covers a font size square.
  const viewArea = view && view.w > 0 && view.h > 0 ? view.w * view.h : 0;
  const share = (w: number, h: number): number =>
    Math.min(1, ((w + stroke) * (h + stroke)) / viewArea);
  const areaCost = (node: Node): number => {
    if (!node.extent && node.glyphs === 0) return 0;
    if (scales || !viewArea) return AREA_COST * Math.max(1, node.glyphs);
    const shape = node.extent ? share(node.extent.w, node.extent.h) : 0;
    return Math.ceil(AREA_COST * (shape + node.glyphs * share(font, font)));
  };

  // The cost of each node with its content, and of each id (the largest of its elements),
  // each computed once. A `<use>` costs its target where it sits, also in `<defs>` and
  // `<symbol>`: browsers build those copies even where they draw nothing.
  const costs = new Map<Node, number>();
  const idCosts = new Map<string, number>();
  const active = new Set<Node>();
  let chain = 0;
  let refused = false;
  const idCost = (id: string, clipOnly: boolean): number => {
    const key = `${clipOnly ? 'c' : 'u'}#${id}`;
    const known = idCosts.get(key);
    if (known !== undefined) return known;
    let largest = 0;
    for (const node of ids.get(id) ?? [])
      if (!clipOnly || node.name === 'clipPath') largest = Math.max(largest, cost(node));
    idCosts.set(key, largest);
    return largest;
  };
  const follow = (id: string, clipOnly: boolean): number => {
    chain += 1;
    const c = chain > MAX_USE_CHAIN ? ((refused = true), Infinity) : idCost(id, clipOnly);
    chain -= 1;
    return c;
  };
  const cost = (node: Node): number => {
    const known = costs.get(node);
    if (known !== undefined) return known;
    if (active.has(node) || refused) return ((refused = true), Infinity);
    active.add(node);
    let total = node.own + areaCost(node);
    for (const child of node.children) {
      total += cost(child);
      if (total > MAX_SVG_COST) break;
    }
    if (node.use !== undefined && total <= MAX_SVG_COST) total += follow(node.use, false);
    for (const id of node.links) if (total <= MAX_SVG_COST) total += follow(id, true);
    active.delete(node);
    costs.set(node, total);
    return total;
  };
  let total = 0;
  for (const root of roots) total += cost(root);
  // A sheet's clip paths and opacity can reach every element.
  let sheetEach = sheetScan?.opacity ? LAYER_COST : 0;
  for (const id of sheetScan?.ids ?? []) sheetEach += idCost(id, true);
  total += sheetEach * elements;
  return !refused && total <= MAX_SVG_COST;
}
