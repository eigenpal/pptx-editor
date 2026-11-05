// Persistent (structurally shared) edits on one XML part, addressed by node id.
//
// An edit rebuilds only the path from the root to the changed node; every untouched subtree is
// shared by reference with the previous revision. That keeps undo a pointer swap and lets
// layout caches keyed by node identity survive edits elsewhere in the slide.

import type { Attr, Element, Node, NodeId, XmlPart } from '../package/tree.ts';
import { deriveKind } from '../package/node-kinds.ts';
import { MAX_XML_DEPTH } from '@eigenpal/ooxml-core';
import { nodeIndex } from './node-index.ts';

export { nodeIndex } from './node-index.ts';

/** Find a node by id. */
export function nodeById(part: XmlPart, id: NodeId): Node | undefined {
  return nodeIndex(part).get(id)?.node;
}

/** The ancestor chain of a node, root first, excluding the node itself. */
export function ancestorsOf(part: XmlPart, id: NodeId): Element[] {
  const index = nodeIndex(part);
  const out: Element[] = [];
  let cur = index.get(id)?.parent ?? null;
  while (cur !== null) {
    const loc = index.get(cur);
    if (!loc || loc.node.type !== 'element') break;
    out.push(loc.node);
    cur = loc.parent;
  }
  return out.reverse();
}

/** Replace the node with `id` by `replace(node)`, rebuilding the path to the root. */
export function updateNode(
  part: XmlPart,
  id: NodeId,
  replace: (node: Node) => Node,
  nextId = part.nextId
): XmlPart | null {
  const index = nodeIndex(part);
  const loc = index.get(id);
  if (!loc) return null;
  let replacement = replace(loc.node);
  let parentId = loc.parent;
  let childIndex = loc.index;
  while (parentId !== null) {
    const parentLoc = index.get(parentId);
    if (!parentLoc || parentLoc.node.type !== 'element') return null;
    const parent = parentLoc.node;
    const kids = parent.children.slice();
    kids[childIndex] = replacement;
    replacement = { ...parent, children: kids };
    childIndex = parentLoc.index;
    parentId = parentLoc.parent;
  }
  if (replacement.type !== 'element') return null;
  return { partName: part.partName, root: replacement, nextId };
}

/** Set or remove an attribute, preserving attribute order. */
export function withAttr(
  el: Element,
  name: { ns: string; local: string; prefix?: string },
  value: string | null
): Element {
  const at = el.attrs.findIndex((a) => a.ns === name.ns && a.local === name.local);
  if (value === null) {
    if (at < 0) return el;
    return { ...el, attrs: el.attrs.filter((_, i) => i !== at) };
  }
  const next: Attr = {
    ns: name.ns,
    local: name.local,
    value,
    ...(name.prefix ? { prefix: name.prefix } : {}),
  };
  const attrs = el.attrs.slice();
  if (at < 0) attrs.push(next);
  else attrs[at] = { ...attrs[at], value };
  return { ...el, attrs };
}

/** A serializable element description, used by ops to create nodes. Ids are minted on apply. */
export interface NodeSpec {
  readonly ns: string;
  readonly local: string;
  readonly prefix?: string;
  readonly attrs?: readonly { ns: string; local: string; prefix?: string; value: string }[];
  readonly nsDecls?: readonly { prefix: string; uri: string }[];
  readonly children?: readonly (NodeSpec | { readonly text: string })[];
}

/**
 * Materialize a spec into nodes with fresh ids, assigning typed kinds under `parentKind` (the kind
 * context of the parent: for a markup-compatibility wrapper, the context it passes through).
 * Kinds follow the reader's rule, so a created node has the kind it has after save and reopen.
 * Only the documented spec fields are copied.
 */
export function materialize(
  spec: NodeSpec,
  parentKind: Element['kind'] | null,
  counter: { next: number },
  depth = 0
): Element {
  if (depth >= MAX_XML_DEPTH) throw new Error('node spec too deep');
  const { kind, childContext } = deriveKind(spec.ns, spec.local, parentKind);
  const id = `n${counter.next++}`;
  const kids: Node[] = [];
  for (const c of spec.children ?? []) {
    if ('text' in c) kids.push({ type: 'text', id: `n${counter.next++}`, value: c.text });
    else kids.push(materialize(c, childContext, counter, depth + 1));
  }
  return {
    type: 'element',
    id,
    kind,
    ns: spec.ns,
    local: spec.local,
    ...(spec.prefix ? { prefix: spec.prefix } : {}),
    nsDecls: (spec.nsDecls ?? []).map((d) => ({ prefix: d.prefix, uri: d.uri })),
    attrs: (spec.attrs ?? []).map((a) => ({
      ns: a.ns,
      local: a.local,
      value: a.value,
      ...(a.prefix ? { prefix: a.prefix } : {}),
    })),
    children: kids,
  };
}

/** Describe an existing subtree as a spec (used by duplicate/copy commands). */
export function specOf(el: Element): NodeSpec {
  return {
    ns: el.ns,
    local: el.local,
    ...(el.prefix ? { prefix: el.prefix } : {}),
    attrs: el.attrs.map((a) => ({ ...a })),
    nsDecls: el.nsDecls,
    children: el.children.map((c) => (c.type === 'text' ? { text: c.value } : specOf(c))),
  };
}
