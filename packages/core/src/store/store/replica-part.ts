// Rebuild an XML part from replicated node records: the trust boundary for collaboration.
//
// A collaboration replica (the pro Yjs binding) keeps each node under a logical key that every
// peer shares, and rebuilds the parts a remote update touched. Peers are untrusted, so the
// rebuilt tree passes the same checks as an `insertNode` spec (`checkSpec`: names, namespace
// bindings, attribute and text validity, depth) plus an element-count cap, before it reaches the
// store. Kinds follow the reader's rule (`deriveKind`), so a rebuilt node has the kind it has
// after save and reopen.
//
// Node ids stay replica-local: a key this replica already knows keeps its store id, so a selection
// on an untouched shape survives a peer's edit. Unknown keys get fresh ids from the part's counter.
// After DOCX Editor's materialize step (`packages/pro/src/collaboration/document/` at adf0e178).

import { MAX_XML_DEPTH, XML_DEFAULT_MAX_ELEMENTS } from '@eigenpal/ooxml-core';
import { deriveKind } from '../package/node-kinds.ts';
import type { Element, Node, NodeId, NodeKind, XmlPart } from '../package/tree.ts';
import { checkSpec } from './op-checks.ts';
import type { NodeSpec } from './tree-edit.ts';

/** A replicated text node: its shared key and its text. @public */
export interface ReplicaTextSpec {
  /** The logical key every peer uses for this node. */
  readonly key: string;
  /** The text content. */
  readonly text: string;
}

/** A replicated element: a node spec whose nodes carry shared keys. @public */
export interface ReplicaElementSpec extends Omit<NodeSpec, 'children'> {
  /** The logical key every peer uses for this node. */
  readonly key: string;
  /** Child elements and text nodes, in document order. */
  readonly children?: readonly (ReplicaElementSpec | ReplicaTextSpec)[];
}

/** The outcome of a rebuild. `ids` maps every key in the part to its store id. @public */
export type ReplicaPartResult =
  | { readonly ok: true; readonly part: XmlPart; readonly ids: ReadonlyMap<string, NodeId> }
  | { readonly ok: false; readonly reason: string };

/** What the replica already knows about the part. @public */
export interface ReplicaPartContext {
  /** The store id this replica uses for a key, if any. */
  idOf(key: string): NodeId | undefined;
  /** The part's id counter: fresh ids start here. */
  readonly nextId: number;
}

function countElements(spec: ReplicaElementSpec, limit: number): number {
  let count = 0;
  const stack: ReplicaElementSpec[] = [spec];
  while (stack.length > 0 && count <= limit) {
    const el = stack.pop()!;
    count += 1;
    for (const c of el.children ?? []) if (!('text' in c)) stack.push(c);
  }
  return count;
}

/**
 * Rebuild one XML part from replicated records. Refuses an invalid tree as a whole; the caller
 * reports the refusal to its session and never publishes a partial part.
 *
 * @public
 */
export function buildReplicaPart(
  partName: string,
  root: ReplicaElementSpec,
  context: ReplicaPartContext
): ReplicaPartResult {
  const height = checkSpec(root, MAX_XML_DEPTH);
  if (typeof height !== 'number') return { ok: false, reason: `invalid node: ${height.code}` };
  if (countElements(root, XML_DEFAULT_MAX_ELEMENTS) > XML_DEFAULT_MAX_ELEMENTS)
    return { ok: false, reason: 'too many elements' };

  const ids = new Map<string, NodeId>();
  const used = new Set<NodeId>();
  // Ids that known keys keep. A fresh id never takes one, even before its key is reached.
  const reserved = new Set(collectKnown(root, context));
  let next = context.nextId;
  const fresh = (): NodeId => {
    let id = `n${next++}`;
    while (used.has(id) || reserved.has(id)) id = `n${next++}`;
    return id;
  };
  const idFor = (key: unknown): NodeId | null => {
    if (typeof key !== 'string' || key.length === 0 || key.length > 256 || ids.has(key))
      return null;
    const known = context.idOf(key);
    const id = known !== undefined && !used.has(known) ? known : fresh();
    ids.set(key, id);
    used.add(id);
    return id;
  };

  const build = (spec: ReplicaElementSpec, parentKind: NodeKind | null): Element | null => {
    const id = idFor(spec.key);
    if (id === null) return null;
    const { kind, childContext } = deriveKind(spec.ns, spec.local, parentKind);
    const children: Node[] = [];
    for (const c of spec.children ?? []) {
      if ('text' in c) {
        const textId = idFor(c.key);
        if (textId === null) return null;
        children.push({ type: 'text', id: textId, value: c.text });
      } else {
        const el = build(c, childContext);
        if (el === null) return null;
        children.push(el);
      }
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
      children,
    };
  };
  const rootEl = build(root, null);
  if (rootEl === null) return { ok: false, reason: 'a key is invalid or appears twice' };
  // The counter stays above every id in the part, kept or fresh.
  let nextId = next;
  for (const id of used) {
    const n = Number(id.slice(1));
    if (id.startsWith('n') && Number.isInteger(n) && n >= nextId) nextId = n + 1;
  }
  return { ok: true, part: { partName, root: rootEl, nextId }, ids };
}

/** The store ids the replica keeps, so fresh ids skip them. */
function collectKnown(root: ReplicaElementSpec, context: ReplicaPartContext): NodeId[] {
  const out: NodeId[] = [];
  const stack: (ReplicaElementSpec | ReplicaTextSpec)[] = [root];
  while (stack.length > 0) {
    const s = stack.pop()!;
    const known = context.idOf(s.key);
    if (known !== undefined) out.push(known);
    if (!('text' in s)) for (const c of s.children ?? []) stack.push(c);
  }
  return out;
}
