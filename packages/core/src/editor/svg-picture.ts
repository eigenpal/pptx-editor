// Whether an SVG picture part is safe for the browser to draw. The browser draws an SVG
// picture on the page's main thread, and `<use>` copies its target: ten levels of ten copies
// each draw 10^10 shapes, and the page stops answering. This check reads the part with the
// bounded XML reader and counts the elements that the picture draws with every `<use>`
// expanded. A part over the limit, one that the reader refuses (an internal DTD subset, too
// deep, too many elements), and one with a `<use>` cycle draw the picture's raster
// alternative instead. The check needs no DOM.

import { readXml, type XmlNode } from '@eigenpal/ooxml-core';

/**
 * The most elements an SVG picture may draw, with each `<use>` expanded. Icons and diagrams
 * draw a few thousand at most.
 */
export const MAX_SVG_DRAWN_ELEMENTS = 200_000;
/** The most `<use>` links that an expansion may follow in a chain. */
const MAX_USE_CHAIN = 16;

/** Elements whose content draws only where a `<use>` links to it. */
const UNDRAWN: ReadonlySet<string> = new Set(['defs', 'symbol']);

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
  // The tree is at most `MAX_XML_DEPTH` deep: a walk on an explicit stack.
  const stack: Element[] = [];
  for (const n of read.nodes) if (n.type === 'element') (roots.push(n), stack.push(n));
  while (stack.length) {
    const el = stack.pop()!;
    const id = el.attributes.id;
    if (id !== undefined && !ids.has(id)) ids.set(id, el);
    for (const c of el.children) if (c.type === 'element') stack.push(c);
  }
  // The drawn size of each element: itself, its children, and the target of its `<use>`.
  const drawn = new Map<Element, number>();
  const active = new Set<Element>();
  let chain = 0;
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
    const target = localName(el.name) === 'use' ? linkTarget(el) : undefined;
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
