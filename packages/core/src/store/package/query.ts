// Read helpers over the canonical tree. Pure and allocation-light: layout calls these on every
// pass, so they never build intermediate maps.

import { q, type QName } from './namespaces.ts';
import type { Element } from './tree.ts';

const nameOf = (name: string | QName): QName => (typeof name === 'string' ? q(name) : name);

/** Whether an element has the given expanded name. */
export function is(el: Element | undefined, name: string | QName): el is Element {
  if (!el) return false;
  const n = nameOf(name);
  return el.ns === n.ns && el.local === n.local;
}

/** The first child element with the given name. */
export function child(el: Element | undefined, name: string | QName): Element | undefined {
  if (!el) return undefined;
  const n = nameOf(name);
  for (const c of el.children)
    if (c.type === 'element' && c.ns === n.ns && c.local === n.local) return c;
  return undefined;
}

/** All child elements with the given name. */
export function children(el: Element | undefined, name: string | QName): Element[] {
  if (!el) return [];
  const n = nameOf(name);
  const out: Element[] = [];
  for (const c of el.children)
    if (c.type === 'element' && c.ns === n.ns && c.local === n.local) out.push(c);
  return out;
}

/** All child elements. */
export function elements(el: Element | undefined): Element[] {
  if (!el) return [];
  const out: Element[] = [];
  for (const c of el.children) if (c.type === 'element') out.push(c);
  return out;
}

/** Follow a chain of child names. */
export function path(el: Element | undefined, ...names: string[]): Element | undefined {
  let cur = el;
  for (const name of names) {
    cur = child(cur, name);
    if (!cur) return undefined;
  }
  return cur;
}

/** An attribute value, by conventional name (`val`, `r:embed`). */
export function attr(el: Element | undefined, name: string | QName): string | undefined {
  if (!el) return undefined;
  const n = nameOf(name);
  for (const a of el.attrs) if (a.local === n.local && a.ns === n.ns) return a.value;
  return undefined;
}

/**
 * An integer attribute. Non-numeric or out-of-range values read as `undefined`, so a corrupt
 * file falls back to the inherited value instead of propagating NaN into layout.
 */
export function intAttr(el: Element | undefined, name: string | QName): number | undefined {
  const raw = attr(el, name);
  if (raw === undefined || !/^-?\d{1,15}$/.test(raw.trim())) return undefined;
  const v = Number(raw);
  return Number.isSafeInteger(v) ? v : undefined;
}

/** An `xsd:boolean` attribute: `1`/`true` and `0`/`false`. */
export function boolAttr(el: Element | undefined, name: string | QName): boolean | undefined {
  const raw = attr(el, name);
  if (raw === '1' || raw === 'true') return true;
  if (raw === '0' || raw === 'false') return false;
  return undefined;
}

/** Concatenated descendant text. */
export function textOf(el: Element | undefined): string {
  if (!el) return '';
  let out = '';
  for (const c of el.children) out += c.type === 'text' ? c.value : textOf(c);
  return out;
}

/** Depth-first search for the first descendant element with the given name. */
export function find(
  el: Element | undefined,
  name: string | QName,
  maxDepth = 64
): Element | undefined {
  if (!el || maxDepth < 0) return undefined;
  const n = nameOf(name);
  for (const c of el.children) {
    if (c.type !== 'element') continue;
    if (c.ns === n.ns && c.local === n.local) return c;
    const hit = find(c, n, maxDepth - 1);
    if (hit) return hit;
  }
  return undefined;
}
