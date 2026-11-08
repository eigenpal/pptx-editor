// The canonical ordered tree for one XML part: typed kinds where layout needs them, a lossless
// generic node for everything else.
//
// Same contract as docx-editor/packages/core/src/store/package/ooxml-tree.ts, reduced to the
// PresentationML needs of v0: names are EXPANDED (namespace URI + local name), prefixes and
// namespace declarations are retained as fidelity evidence, attribute records are arrays (so a
// file-supplied `__proto__` is inert by construction), and every element and text node has a
// stable id minted deterministically in document order. The same bytes always produce the same
// ids, and an edit never renumbers untouched nodes. DOCX Editor commit adf0e178.

import { readXml, type XmlNode, type XmlLimits, type XmlRejection } from '@eigenpal/ooxml-core';
import { NS } from './namespaces.ts';
import { deriveKind } from './node-kinds.ts';
import { isNcName, namespaceBindingError, qualifiedNameError } from './xml-names.ts';

export { kindOf } from './node-kinds.ts';

/** A node's stable identity within its part. */
export type NodeId = string;

/**
 * The typed kinds. Everything not listed is `generic`, carried verbatim.
 *
 * A kind is a LABEL on the same node shape, assigned from the expanded name and the parent's
 * kind. Layout reads typed kinds; everything else only has to survive save.
 */
export type NodeKind =
  | 'generic'
  | 'spTree'
  | 'grpSp'
  | 'sp'
  | 'pic'
  | 'cxnSp'
  | 'graphicFrame'
  | 'txBody'
  | 'paragraph'
  | 'run'
  | 'break'
  | 'field'
  | 'text';

/** One `xmlns` declaration authored on an element. `prefix` is '' for the default namespace. */
export interface NamespaceDecl {
  readonly prefix: string;
  readonly uri: string;
}

/** One attribute, carried verbatim. */
export interface Attr {
  readonly ns: string;
  readonly local: string;
  readonly prefix?: string;
  readonly value: string;
}

/** An element node. Immutable; edits produce new nodes along the changed path. */
export interface Element {
  readonly type: 'element';
  readonly id: NodeId;
  readonly kind: NodeKind;
  readonly ns: string;
  readonly local: string;
  readonly prefix?: string;
  readonly nsDecls: readonly NamespaceDecl[];
  readonly attrs: readonly Attr[];
  readonly children: readonly Node[];
}

/** A text node. Only retained where it is significant (see {@link isSignificantText}). */
export interface TextNode {
  readonly type: 'text';
  readonly id: NodeId;
  readonly value: string;
}

export type Node = Element | TextNode;

/** One parsed XML part. `nextId` is the id counter for nodes created by later edits. */
export interface XmlPart {
  readonly partName: string;
  readonly root: Element;
  readonly nextId: number;
}

/** Why a part was refused at the trust boundary. */
export type TreeRejection = XmlRejection | 'no-root' | 'unbound-prefix' | 'bad-name';

export type TreeResult =
  | { readonly ok: true; readonly part: XmlPart }
  | { readonly ok: false; readonly reason: TreeRejection; readonly detail?: string };

/** Text content is significant inside `a:t` and in generic leaf elements that hold only text. */
function isSignificantText(parentKind: NodeKind, siblings: readonly XmlNode[]): boolean {
  if (parentKind === 'text') return true;
  return siblings.every((s) => s.type === 'text');
}

class TreeBuildError extends Error {
  constructor(
    readonly reason: TreeRejection,
    detail: string
  ) {
    super(detail);
  }
}

/** Parse one XML part into the canonical tree. Never throws: the bytes are untrusted. */
export function parseXmlPart(partName: string, xml: string, limits?: XmlLimits): TreeResult {
  const read = readXml(xml, limits);
  if (!read.ok) return { ok: false, reason: read.reason };
  const roots = read.nodes.filter((n) => n.type === 'element');
  if (roots.length !== 1 || roots[0].type !== 'element') return { ok: false, reason: 'no-root' };
  const counter = { next: 1 };
  try {
    const root = build(
      roots[0],
      { decls: new Map([['xml', NS.xml]]), parent: null },
      null,
      counter
    );
    if (root.type !== 'element') return { ok: false, reason: 'no-root' };
    return { ok: true, part: { partName, root, nextId: counter.next } };
  } catch (e) {
    if (e instanceof TreeBuildError) return { ok: false, reason: e.reason, detail: e.message };
    return { ok: false, reason: 'parse-error' };
  }
}

function splitName(raw: string): { prefix: string; local: string } {
  const colon = raw.indexOf(':');
  if (colon < 0) return { prefix: '', local: raw };
  return { prefix: raw.slice(0, colon), local: raw.slice(colon + 1) };
}

/**
 * The prefixes in scope: an element's own declarations, then its parent's scope. A chain, not a
 * copy per declaring element: copying a large inherited scope at each of many declaring
 * elements was quadratic. The XML depth limit bounds the walk.
 */
interface NamespaceScope {
  readonly decls: ReadonlyMap<string, string>;
  readonly parent: NamespaceScope | null;
}

function lookupPrefix(scope: NamespaceScope, prefix: string): string | undefined {
  for (let s: NamespaceScope | null = scope; s; s = s.parent) {
    const uri = s.decls.get(prefix);
    if (uri !== undefined) return uri;
  }
  return undefined;
}

function build(
  node: Extract<XmlNode, { type: 'element' }>,
  scope: NamespaceScope,
  parentKind: NodeKind | null,
  counter: { next: number }
): Element {
  const nsDecls: NamespaceDecl[] = [];
  let own: Map<string, string> | null = null;
  const rawAttrs: [string, string][] = [];
  for (const name of Object.keys(node.attributes)) {
    const value = node.attributes[name];
    if (name === 'xmlns' || name.startsWith('xmlns:')) {
      const prefix = name === 'xmlns' ? '' : name.slice(6);
      if (prefix !== '' && !isNcName(prefix)) throw new TreeBuildError('bad-name', name);
      // The same binding rules as an op and the serializer: a deck that opens can be saved.
      const bindingError = namespaceBindingError(prefix, value);
      if (bindingError) throw new TreeBuildError('bad-name', `${name}: ${bindingError.detail}`);
      nsDecls.push({ prefix, uri: value });
      own ??= new Map();
      own.set(prefix, value);
    } else rawAttrs.push([name, value]);
  }
  const local: NamespaceScope = own ? { decls: own, parent: scope } : scope;
  const { prefix, local: localName } = splitName(node.name);
  if (!isNcName(localName) || (prefix && !isNcName(prefix)))
    throw new TreeBuildError('bad-name', node.name);
  const ns = lookupPrefix(local, prefix);
  if (ns === undefined && prefix !== '') throw new TreeBuildError('unbound-prefix', node.name);
  const elementError = qualifiedNameError(ns ?? '', localName, prefix || undefined, false);
  if (elementError) throw new TreeBuildError('bad-name', `${node.name}: ${elementError.detail}`);
  const expanded = new Set<string>();
  const attrs: Attr[] = rawAttrs.map(([name, value]) => {
    const split = splitName(name);
    if (!isNcName(split.local) || (split.prefix && !isNcName(split.prefix)))
      throw new TreeBuildError('bad-name', name);
    const attrNs = split.prefix === '' ? '' : lookupPrefix(local, split.prefix);
    if (attrNs === undefined) throw new TreeBuildError('unbound-prefix', name);
    const attrError = qualifiedNameError(attrNs, split.local, split.prefix || undefined, true);
    if (attrError) throw new TreeBuildError('bad-name', `${name}: ${attrError.detail}`);
    // Two prefixes bound to one URI can spell the same expanded attribute twice.
    const key = `{${attrNs}}${split.local}`;
    if (expanded.has(key)) throw new TreeBuildError('bad-name', `${name}: repeated attribute`);
    expanded.add(key);
    if (split.prefix === '') return { ns: '', local: split.local, value };
    return { ns: attrNs, local: split.local, prefix: split.prefix, value };
  });
  const { kind, childContext } = deriveKind(ns ?? '', localName, parentKind);
  const id = `n${counter.next++}`;
  const children: Node[] = [];
  const keepText = isSignificantText(kind, node.children);
  for (const child of node.children) {
    if (child.type === 'text') {
      if (!keepText && child.value.trim() === '') continue;
      children.push({ type: 'text', id: `n${counter.next++}`, value: child.value });
    } else children.push(build(child, local, childContext, counter));
  }
  return {
    type: 'element',
    id,
    kind,
    ns: ns ?? '',
    local: localName,
    ...(prefix ? { prefix } : {}),
    nsDecls,
    attrs,
    children,
  };
}
