// A DOM-free SVG document: the part of the DOM `paintSlide` builds with, and an XML serializer.
//
// It lets a host with no DOM (Node.js, Bun, a worker) run the editor's own paint and get its
// SVG as text, so a picture on a server comes from the same `paintSlide` as the canvas: no
// second renderer. It implements only what the paint uses: `createElementNS`, `setAttribute`,
// `getAttribute`, `appendChild`, `insertBefore`, `removeChild`, `children`, `firstChild`,
// `parentNode`, `tagName`, and `textContent`.
//
// Nothing is parsed from a string. The serializer escapes every attribute value and text, and
// replaces a character that XML 1.0 does not allow, so file text cannot break out of the markup
// (AGENTS.md "Security"). Element and attribute names come from the paint code; a name that is
// not a plain XML name throws.

import { escapeXml, escapeXmlAttribute } from '../store/index.ts';

/** A plain XML name, as the paint writes element and attribute names. */
const XML_NAME = /^[A-Za-z_][A-Za-z0-9_.:-]*$/;
/** A character outside XML 1.0. */
const INVALID_XML = /[^\u0009\u000A\u000D\u0020-\uD7FF\uE000-\uFFFD\u{10000}-\u{10FFFF}]/gu;
/** The deepest tree the serializer writes (groups nest up to 32 deep; each adds a few levels). */
const MAX_DEPTH = 512;

const clean = (value: string): string => value.replace(INVALID_XML, '�');

function checkName(name: string): string {
  if (!XML_NAME.test(name)) throw new Error('svg-document: not an XML name');
  return name;
}

/** One element of a {@link SvgDocument}. */
export class SvgNode {
  readonly tagName: string;
  readonly namespaceURI: string;
  parentNode: SvgNode | null = null;
  readonly #attributes = new Map<string, string>();
  #children: SvgNode[] = [];
  #text = '';

  constructor(namespace: string, tag: string) {
    this.namespaceURI = namespace;
    this.tagName = checkName(tag);
  }

  get children(): readonly SvgNode[] {
    return this.#children;
  }
  get firstChild(): SvgNode | null {
    return this.#children[0] ?? null;
  }
  get textContent(): string {
    return this.#text + this.#children.map((c) => c.textContent).join('');
  }
  set textContent(value: string | null) {
    for (const c of this.#children) c.parentNode = null;
    this.#children = [];
    this.#text = value ?? '';
  }

  setAttribute(name: string, value: string): void {
    this.#attributes.set(checkName(name), String(value));
  }
  getAttribute(name: string): string | null {
    return this.#attributes.get(name) ?? null;
  }
  hasAttribute(name: string): boolean {
    return this.#attributes.has(name);
  }
  removeAttribute(name: string): void {
    this.#attributes.delete(name);
  }

  appendChild<T extends SvgNode>(node: T): T {
    return this.insertBefore(node, null);
  }
  insertBefore<T extends SvgNode>(node: T, before: SvgNode | null): T {
    if (!(node instanceof SvgNode)) throw new Error('svg-document: not a node of this document');
    for (let up: SvgNode | null = this; up; up = up.parentNode)
      if (up === node) throw new Error('svg-document: a node cannot hold its ancestor');
    node.parentNode?.removeChild(node);
    const at = before ? this.#children.indexOf(before) : -1;
    if (at < 0) this.#children.push(node);
    else this.#children.splice(at, 0, node);
    node.parentNode = this;
    return node;
  }
  removeChild<T extends SvgNode>(node: T): T {
    const at = this.#children.indexOf(node);
    if (at < 0) throw new Error('svg-document: not a child');
    this.#children.splice(at, 1);
    node.parentNode = null;
    return node;
  }
  remove(): void {
    this.parentNode?.removeChild(this);
  }

  /** The element and its descendants as XML text. */
  serialize(): string {
    const out: string[] = [];
    write(this, out, 0, null);
    return out.join('');
  }

  /** @internal */
  get attributeEntries(): readonly (readonly [string, string])[] {
    return [...this.#attributes];
  }
  /** @internal The text set on this element itself. */
  get ownText(): string {
    return this.#text;
  }
}

function write(node: SvgNode, out: string[], depth: number, parentNs: string | null): void {
  if (depth > MAX_DEPTH) throw new Error('svg-document: the tree is too deep');
  out.push('<', node.tagName);
  if (node.namespaceURI !== parentNs && !node.hasAttribute('xmlns'))
    out.push(' xmlns="', escapeXmlAttribute(node.namespaceURI), '"');
  for (const [name, value] of node.attributeEntries)
    out.push(' ', name, '="', escapeXmlAttribute(clean(value)), '"');
  const text = node.ownText;
  if (!text && node.children.length === 0) {
    out.push('/>');
    return;
  }
  out.push('>');
  if (text) out.push(escapeXml(clean(text)));
  for (const c of node.children) write(c, out, depth + 1, node.namespaceURI);
  out.push('</', node.tagName, '>');
}

/** The document that makes {@link SvgNode}s. */
export class SvgDocument {
  createElementNS(namespace: string, tag: string): SvgNode {
    return new SvgNode(namespace, tag);
  }
}

/**
 * @internal A DOM-free document for `paintSlide`, typed as the DOM `Document` the paint takes.
 * Only the members listed in this module's header exist on it.
 */
export function createSvgDocument(): Document {
  return new SvgDocument() as unknown as Document;
}

/** @internal The XML text of an element made by {@link createSvgDocument}, or of a DOM element. */
export function serializeSvgElement(root: Element): string {
  if ((root as unknown) instanceof SvgNode) return (root as unknown as SvgNode).serialize();
  return new XMLSerializer().serializeToString(root);
}

/**
 * @internal The descendants of `root` (and `root` itself) that carry attribute `name`, in document order,
 * through `children` and `getAttribute` only, so it walks a DOM tree and a {@link SvgDocument}
 * tree alike.
 */
export function elementsWithAttribute(root: Element, name: string): Element[] {
  const out: Element[] = [];
  const visit = (el: Element, depth: number): void => {
    if (depth > MAX_DEPTH) return;
    if (el.getAttribute(name) !== null) out.push(el);
    for (const c of Array.from(el.children)) visit(c, depth + 1);
  };
  visit(root, 0);
  return out;
}
