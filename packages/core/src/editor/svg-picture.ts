// Whether an SVG picture part is safe for the browser to draw. The browser draws an SVG
// picture on the page's main thread, and links multiply its work: `<use>` copies its target,
// so ten levels of ten copies each draw 10^10 shapes; a mask, a pattern, a clip path, or a
// filter draws its content for each element that links to it; and a marker draws its content
// at each vertex of each path that it reaches. This check reads the part with the bounded XML
// reader and counts, as an upper bound, the elements that the picture draws with every link
// expanded. A picture over the limit draws its raster alternative instead, as every SVG picture
// of a presentation carries one, so a false refusal costs only the vector picture.
//
// The check also refuses what it cannot count: an internal DTD subset, a part that the reader
// refuses, an `xml-stylesheet` processing instruction, a `<foreignObject>`, and a `data:` URI
// that can hold a document (svg-picture-refs.ts). The check needs no DOM.

import { readXml, type XmlNode } from '@eigenpal/ooxml-core';
import {
  hrefId,
  loadsDocument,
  setsMarker,
  sheetRules,
  urlIds,
  type SheetRule,
} from './svg-picture-refs.ts';

/**
 * The most elements an SVG picture may draw, with each link expanded and each marker drawn at
 * each vertex. Icons and diagrams draw a few thousand at most.
 */
export const MAX_SVG_DRAWN_ELEMENTS = 200_000;
/** The most links that an expansion may follow in a chain. */
const MAX_LINK_CHAIN = 16;

/** Elements whose content draws only where a link reaches it (lower-case local names). */
const UNDRAWN: ReadonlySet<string> = new Set([
  'defs',
  'symbol',
  'marker',
  'mask',
  'pattern',
  'clippath',
  'filter',
  'lineargradient',
  'radialgradient',
]);
/** The elements that draw markers (SVG 2, 11.6.2 "markable elements"). */
const MARKABLE: ReadonlySet<string> = new Set(['path', 'line', 'polyline', 'polygon']);

type Element = Extract<XmlNode, { type: 'element' }>;

const kindOf = (el: Element): string => el.name.slice(el.name.indexOf(':') + 1).toLowerCase();

/** A document type declaration without an internal subset: it declares nothing to expand. */
const PLAIN_DOCTYPE = /<!DOCTYPE[^[>]*>/;

// --- Vertices ------------------------------------------------------------------------------

const TOKEN = /([a-z])|[+-]?(?:\d+\.?\d*|\.\d+)(?:e[+-]?\d+)?/gi;
/**
 * The numbers of one segment of each path command. An arc counts 4, not 7: its two flags can
 * be written with no space between them and the next number.
 */
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

/** At most how many vertices path data or a point list `text` has. */
function countVertices(text: string, points: boolean): number {
  let total = 0;
  let per = 2;
  let run = 0;
  for (const m of text.matchAll(TOKEN)) {
    const letter = m[1]?.toLowerCase();
    if (letter === undefined || points) {
      run += 1;
      continue;
    }
    total += Math.ceil(run / per);
    run = 0;
    if (letter === 'z') total += 1;
    else per = SEGMENT[letter] ?? 1;
  }
  return total + Math.ceil(run / per) + 1;
}

/** At most how many vertices a markable element has. A style can set path data (`d`) too. */
function vertices(el: Element, sheetNumbers: number): number {
  const kind = kindOf(el);
  if (kind === 'line') return 2;
  const style = el.attributes.style ?? '';
  const own = kind === 'path' ? (el.attributes.d ?? '') : (el.attributes.points ?? '');
  return countVertices(own, kind !== 'path') + countVertices(style, true) + sheetNumbers;
}

// --- The check -----------------------------------------------------------------------------

const results = new WeakMap<Uint8Array, boolean>();

/** Whether the SVG document `bytes` draws within the limits (see the module comment). */
export function svgDrawable(bytes: Uint8Array): boolean {
  const hit = results.get(bytes);
  if (hit !== undefined) return hit;
  const ok = check(bytes);
  results.set(bytes, ok);
  return ok;
}

/** A link from an element: its target id, and whether it is a marker property. */
interface Link {
  readonly id: string;
  readonly marker: boolean;
}

function check(bytes: Uint8Array): boolean {
  let text: string;
  try {
    text = new TextDecoder('utf-8', { fatal: false }).decode(bytes);
  } catch {
    return false;
  }
  // A style sheet from outside the tree: this check cannot read it.
  if (/<\?xml-stylesheet/i.test(text)) return false;
  const read = readXml(text.replace(PLAIN_DOCTYPE, ''), {
    maxBytes: bytes.length + 1,
    maxElements: MAX_SVG_DRAWN_ELEMENTS,
  });
  if (!read.ok) return false;
  const roots: Element[] = [];
  for (const n of read.nodes) if (n.type === 'element') roots.push(n);

  // Every element with an id (an id can repeat), the style sheets, and the refusals.
  const ids = new Map<string, Element[]>();
  const markers: string[] = [];
  let sheets = '';
  let customProperties = false;
  for (const stack = [...roots]; stack.length;) {
    const el = stack.pop()!;
    const kind = kindOf(el);
    if (kind === 'foreignobject') return false;
    const id = el.attributes.id;
    if (id !== undefined) {
      const same = ids.get(id);
      if (same) same.push(el);
      else ids.set(id, [el]);
      if (kind === 'marker') markers.push(id);
    }
    for (const value of Object.values(el.attributes)) {
      if (loadsDocument(value)) return false;
      if (/var\(/i.test(value)) customProperties = true;
    }
    if (kind === 'style')
      for (const c of el.children) if (c.type === 'text') sheets += `${c.value}\n`;
    for (const c of el.children) if (c.type === 'element') stack.push(c);
  }
  if (loadsDocument(sheets, true)) return false;

  // The style sheet rules that link to ids. A sheet or a custom property that this check
  // cannot read sets every marker on every markable element, and links every element to each
  // id that the sheets name.
  const rules = customProperties ? null : sheetRules(sheets);
  const sheetIds = rules ? [] : urlIds(sheets);
  const sheetNumbers = /\bd\s*:/i.test(sheets) ? countVertices(sheets, true) : 0;
  const matches = (rule: SheetRule, el: Element, kind: string): boolean => {
    if (rule.type && rule.type !== kind) return false;
    if (rule.id !== undefined && el.attributes.id !== rule.id) return false;
    if (rule.classes.length === 0) return true;
    const own = new Set((el.attributes.class ?? '').split(/\s+/));
    return rule.classes.every((c) => own.has(c));
  };
  /** The links of `el`: its `href`, each `url(#id)` it names, and the sheet rules it matches. */
  const linksOf = (el: Element, kind: string): Link[] => {
    const links: Link[] = [];
    for (const [name, value] of Object.entries(el.attributes)) {
      const local = name.slice(name.indexOf(':') + 1).toLowerCase();
      if (local === 'href') {
        const id = hrefId(value);
        if (id !== undefined) links.push({ id, marker: false });
        continue;
      }
      const marker = local.startsWith('marker') || (local === 'style' && setsMarker(value));
      for (const id of urlIds(value)) links.push({ id, marker });
    }
    if (rules) {
      for (const rule of rules)
        if (matches(rule, el, kind))
          for (const id of rule.ids) links.push({ id, marker: rule.marker });
    } else {
      for (const id of sheetIds) links.push({ id, marker: false });
      if (MARKABLE.has(kind)) for (const id of markers) links.push({ id, marker: true });
    }
    if (customProperties && MARKABLE.has(kind))
      for (const id of markers) links.push({ id, marker: true });
    return links;
  };

  // The drawn size of each element, and the vertices it draws: itself, its children, each
  // target it links to, and the largest marker that it links to at each vertex it draws.
  type Size = { readonly drawn: number; readonly vertices: number };
  const INFINITE: Size = { drawn: Infinity, vertices: Infinity };
  const NOTHING: Size = { drawn: 0, vertices: 0 };
  const sizes = new Map<Element, Size>();
  const active = new Set<Element>();
  let chain = 0;
  /** The size of a link to `id`: the largest of the elements with that id. */
  const target = (id: string, marker: boolean): Size => {
    let largest = NOTHING;
    for (const el of ids.get(id) ?? []) {
      // A marker in its own content: the browser breaks the cycle and draws it once.
      const s = active.has(el) ? (marker ? NOTHING : INFINITE) : size(el);
      if (s.drawn > largest.drawn) largest = s;
    }
    return largest;
  };
  const size = (el: Element): Size => {
    const known = sizes.get(el);
    if (known) return known;
    active.add(el);
    const kind = kindOf(el);
    let drawn = 1;
    let verts = MARKABLE.has(kind) ? vertices(el, sheetNumbers) : 0;
    for (const c of el.children) {
      if (c.type !== 'element') continue;
      if (drawn > MAX_SVG_DRAWN_ELEMENTS) break;
      // Content that draws only through a link counts there.
      if (UNDRAWN.has(kindOf(c))) drawn += 1;
      else {
        const s = size(c);
        drawn += s.drawn;
        verts += s.vertices;
      }
    }
    const links = linksOf(el, kind);
    const follow = (link: Link): Size => {
      chain += 1;
      const s = chain > MAX_LINK_CHAIN ? INFINITE : target(link.id, link.marker);
      chain -= 1;
      return s;
    };
    for (const link of links) {
      if (drawn > MAX_SVG_DRAWN_ELEMENTS) break;
      if (link.marker) continue;
      const s = follow(link);
      drawn += s.drawn;
      verts += s.vertices;
    }
    // A marker draws at each vertex that this element and its content draw.
    for (const link of links) {
      if (drawn > MAX_SVG_DRAWN_ELEMENTS) break;
      if (link.marker) drawn += follow(link).drawn * (verts + 1);
    }
    active.delete(el);
    const s = { drawn, vertices: verts };
    sizes.set(el, s);
    return s;
  };
  let total = 0;
  for (const root of roots) total += size(root).drawn;
  return total <= MAX_SVG_DRAWN_ELEMENTS;
}
