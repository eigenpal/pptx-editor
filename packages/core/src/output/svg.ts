// SVG construction helpers. Every attribute value goes through `setAttribute` and every text
// through `textContent`: no markup is ever parsed from a string (AGENTS.md "Security").

export const SVG_NS = 'http://www.w3.org/2000/svg';

/** Attribute values: numbers are rounded to keep the DOM compact and deterministic. */
export type AttrValue = string | number | undefined | null | false;

const round = (v: number): string => {
  if (!Number.isFinite(v)) return '0';
  const r = Math.round(v * 100) / 100;
  return Object.is(r, -0) ? '0' : String(r);
};

/** Format a number for an SVG attribute. */
export function num(v: number): string {
  return round(v);
}

/** Create an SVG element with attributes. */
export function svg(
  doc: Document,
  tag: string,
  attrs: Record<string, AttrValue> = {},
  children: readonly (Element | null | undefined)[] = []
): SVGElement {
  const el = doc.createElementNS(SVG_NS, tag) as SVGElement;
  for (const [k, v] of Object.entries(attrs)) {
    if (v === undefined || v === null || v === false) continue;
    el.setAttribute(k, typeof v === 'number' ? round(v) : v);
  }
  for (const c of children) if (c) el.appendChild(c);
  return el;
}

/** A `<defs>` registry that hands out unique ids within one painted slide. */
export class Defs {
  readonly element: SVGElement;
  #next = 0;
  #captured: SVGElement[] | null = null;
  constructor(
    readonly doc: Document,
    readonly prefix: string
  ) {
    this.element = svg(doc, 'defs');
  }

  /** A fresh id. */
  id(kind: string): string {
    this.#next += 1;
    return `${this.prefix}-${kind}${this.#next}`;
  }

  /** The definition a `url(#id)` paint refers to, if this registry holds it. */
  find(url: string): SVGElement | undefined {
    const id = /^url\(#(.+)\)$/.exec(url)?.[1];
    for (const c of this.element.children) if (c.getAttribute('id') === id) return c as SVGElement;
    return undefined;
  }

  /** Add a definition and return `url(#id)`. */
  add(el: SVGElement): string {
    this.element.appendChild(el);
    this.#captured?.push(el);
    return `url(#${el.getAttribute('id')})`;
  }

  /** Run `paint`, and return what it painted with the definitions it added. */
  capture<T>(paint: () => T): { readonly value: T; readonly defs: readonly SVGElement[] } {
    const outer = this.#captured;
    const defs: SVGElement[] = [];
    this.#captured = defs;
    try {
      return { value: paint(), defs };
    } finally {
      this.#captured = outer;
      outer?.push(...defs);
    }
  }
}

/** Format a matrix for `transform`. */
export function matrixAttr(m: readonly number[]): string {
  return `matrix(${m.map(round).join(' ')})`;
}
