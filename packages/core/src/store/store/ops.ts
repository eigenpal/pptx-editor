// The op vocabulary: the ONLY way to change a document.
//
// Every op is plain JSON (no functions, no node references), so a transaction can be logged,
// replayed, or sent to a collaboration peer. Ops address nodes by part name + stable node id.
// An op that cannot apply exactly is REFUSED with a typed rejection; nothing is approximated.
// Because ops arrive as JSON, their shape, names, and namespaces are checked at run time
// (`op-checks.ts`): a malformed op is refused, never a TypeError.

import { isCanonicalPartName, partNameKey } from '@eigenpal/ooxml-core';
import type { PackagePart, PptxPackage } from '../package/package.ts';
import { DerivedParts } from '../package/parts-map.ts';
import { isValidXmlText } from '@eigenpal/ooxml-core';
import { parseXmlPart, type Element } from '../package/tree.ts';
import { nodeIndex } from './node-index.ts';
import { MAX_XML_DEPTH } from '@eigenpal/ooxml-core';
import { checkOpShape, checkSpec } from './op-checks.ts';
import { materialize, withAttr, type NodeSpec } from './tree-edit.ts';
import { rekind, subtreeHeight, WorkingTree } from './working-tree.ts';

/** An expanded attribute name with an optional preferred prefix. */
export interface AttrName {
  readonly ns: string;
  readonly local: string;
  readonly prefix?: string;
}

/** One serializable document operation. */
export type PptxOp =
  | {
      readonly op: 'setAttribute';
      readonly part: string;
      readonly node: string;
      readonly name: AttrName;
      /** `null` removes the attribute. */
      readonly value: string | null;
    }
  | {
      /** Replace all children of an element with one text node (e.g. `a:t`). */
      readonly op: 'setText';
      readonly part: string;
      readonly node: string;
      readonly text: string;
    }
  | {
      readonly op: 'insertNode';
      readonly part: string;
      readonly parent: string;
      readonly index: number;
      readonly spec: NodeSpec;
    }
  | { readonly op: 'removeNode'; readonly part: string; readonly node: string }
  | {
      readonly op: 'moveNode';
      readonly part: string;
      readonly node: string;
      readonly parent: string;
      readonly index: number;
    }
  | {
      /** Add a part. `xml` is the serialized XML text; `base64` is binary content. */
      readonly op: 'addPart';
      readonly part: string;
      readonly xml?: string;
      readonly base64?: string;
    }
  | { readonly op: 'removePart'; readonly part: string };

export type PptxOpKind = PptxOp['op'];

/** Every op kind, for coverage gates. */
export const PPTX_OP_KINDS: readonly PptxOpKind[] = [
  'setAttribute',
  'setText',
  'insertNode',
  'removeNode',
  'moveNode',
  'addPart',
  'removePart',
];

/**
 * A rule of one part's own tree that the store checks when a transaction commits
 * (`tree-invariant`):
 *
 * - `table-grid`: every `a:tr` holds one `a:tc` per `a:gridCol` (ECMA-376 Part 1, 21.1.3.16).
 * - `table-merge`: `gridSpan` and `rowSpan` stay inside the grid, merged blocks do not overlap,
 *   and the `hMerge` and `vMerge` markers agree with the blocks (21.1.3.16).
 * - `shape-id`: `p:cNvPr/@id` is unique among the shapes of one slide, layout, or master.
 *
 * @public
 */
export type TreeInvariantRule = 'table-grid' | 'table-merge' | 'shape-id';

/**
 * A rule that reads more than one part, checked when a transaction commits
 * (`package-invariant`):
 *
 * - `slide-list`: each `p:sldId`, `p:sldMasterId`, and `p:notesMasterId` names, by `r:id`, a
 *   relationship of the presentation part to a part of the package (ECMA-376 Part 1, 19.2.1.33).
 * - `relationship-target`: each internal relationship targets a part of the package (ECMA-376
 *   Part 2, relationships).
 * - `content-type`: each part has a content type, by `Override` or by `Default` (ECMA-376 Part 2,
 *   content types).
 *
 * @public
 */
export type PackageInvariantRule = 'slide-list' | 'relationship-target' | 'content-type';

/**
 * Why an op was refused.
 *
 * Name and namespace refusals: `invalid-name` (not an XML NCName), `reserved-name` (the `xmlns`
 * prefix or namespace, the `xml` prefix bound to another namespace, an unprefixed `xmlns`
 * attribute), `invalid-namespace` (a URI that is not valid XML text, or a binding XML 1.0 cannot
 * express), and `duplicate-name` (one attribute or prefix twice on an element). `too-deep`
 * refuses a node the reader would refuse on reopen. `malformed-op` refuses an op whose JSON
 * shape is wrong.
 *
 * Commit refusals: `tree-invariant` (a touched part's own tree breaks a
 * {@link TreeInvariantRule}) and `package-invariant` (a {@link PackageInvariantRule}, which reads
 * more than one part). The store checks the package a transaction leaves, in the parts it
 * touched, and refuses only a violation the transaction introduced. `part` names the part the
 * violation is in, `rule` the broken rule, and `detail` what is wrong, in English, for logs.
 * Such a refusal means the ops would save a broken package; undo and redo are never refused.
 */
export type OpRejection =
  | { readonly code: 'unknown-part'; readonly part: string }
  | { readonly code: 'not-xml-part'; readonly part: string }
  | { readonly code: 'unknown-node'; readonly part: string; readonly node: string }
  | { readonly code: 'not-an-element'; readonly node: string }
  | { readonly code: 'invalid-text'; readonly detail: string }
  | { readonly code: 'invalid-index'; readonly index: number }
  | { readonly code: 'root-node'; readonly node: string }
  | { readonly code: 'cycle'; readonly node: string }
  | { readonly code: 'part-exists'; readonly part: string }
  | { readonly code: 'unsafe-part-name'; readonly part: string }
  | { readonly code: 'invalid-part-content'; readonly part: string; readonly detail: string }
  | { readonly code: 'invalid-name'; readonly name: string; readonly detail: string }
  | { readonly code: 'reserved-name'; readonly name: string; readonly detail: string }
  | { readonly code: 'invalid-namespace'; readonly uri: string; readonly detail: string }
  | { readonly code: 'duplicate-name'; readonly name: string }
  | { readonly code: 'too-deep'; readonly detail: string }
  | { readonly code: 'malformed-op'; readonly detail: string }
  | { readonly code: 'unsupported'; readonly detail: string }
  | {
      readonly code: 'tree-invariant';
      readonly part: string;
      readonly rule: TreeInvariantRule;
      readonly detail: string;
    }
  | {
      readonly code: 'package-invariant';
      readonly part: string;
      readonly rule: PackageInvariantRule;
      readonly detail: string;
    };

export type OpResult =
  | { readonly ok: true; readonly pkg: PptxPackage }
  | { readonly ok: false; readonly rejection: OpRejection };

function base64ToBytes(b64: string): Uint8Array | null {
  try {
    const bin = atob(b64);
    const out = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i += 1) out[i] = bin.charCodeAt(i);
    return out;
  } catch {
    return null;
  }
}

const validIndex = (index: number, length: number): boolean =>
  Number.isInteger(index) && index >= 0 && index <= length;

/**
 * One transaction against a base package. Parts are copied on first write, and each XML part
 * gets one {@link WorkingTree} that carries its id index across every op.
 */
class Transaction {
  readonly #base: PptxPackage;
  #parts: DerivedParts | null = null;
  readonly #trees = new Map<string, WorkingTree>();
  /** OPC-equivalent part names (`partNameKey`), built on the first `addPart`. */
  #keys: Set<string> | null = null;

  constructor(base: PptxPackage) {
    this.#base = base;
  }

  /** The writable part map: it records changes over the base, never a copy of every part. */
  #partsForWrite(): DerivedParts {
    this.#parts ??= DerivedParts.from(this.#base.parts);
    return this.#parts;
  }

  #tree(part: string): WorkingTree | OpRejection {
    const open = this.#trees.get(part);
    if (open) return open;
    const p = (this.#parts ?? this.#base.parts).get(part);
    if (!p) return { code: 'unknown-part', part };
    if (p.kind !== 'xml') return { code: 'not-xml-part', part };
    const tree = new WorkingTree(p.xml);
    this.#trees.set(part, tree);
    return tree;
  }

  /** The id the next node inserted into `part` gets: the open tree's, else the part's. */
  nextId(part: string): number | undefined {
    const open = this.#trees.get(part);
    if (open) return open.nextId;
    const p = (this.#parts ?? this.#base.parts).get(part);
    return p?.kind === 'xml' ? p.xml.nextId : undefined;
  }

  /** The element `id` of `part`: from the open tree, else from the part's own index. */
  element(part: string, id: string): Element | undefined {
    const open = this.#trees.get(part);
    if (open) return open.element(id);
    const p = (this.#parts ?? this.#base.parts).get(part);
    const node = p?.kind === 'xml' ? nodeIndex(p.xml).get(id)?.node : undefined;
    return node?.type === 'element' ? node : undefined;
  }

  apply(op: PptxOp): OpRejection | null {
    const shape = checkOpShape(op);
    if (shape) return shape;
    switch (op.op) {
      case 'setAttribute': {
        const tree = this.#tree(op.part);
        if (!(tree instanceof WorkingTree)) return tree;
        const loc = tree.get(op.node);
        if (!loc) return { code: 'unknown-node', part: op.part, node: op.node };
        if (loc.node.type !== 'element') return { code: 'not-an-element', node: op.node };
        if (op.value !== null && !isValidXmlText(op.value))
          return { code: 'invalid-text', detail: 'attribute value' };
        const next = withAttr(loc.node, op.name, op.value);
        if (next !== loc.node) tree.replace(op.node, next);
        return null;
      }
      case 'setText': {
        const tree = this.#tree(op.part);
        if (!(tree instanceof WorkingTree)) return tree;
        if (!isValidXmlText(op.text)) return { code: 'invalid-text', detail: 'text' };
        const el = tree.get(op.node)?.node;
        if (!el) return { code: 'unknown-node', part: op.part, node: op.node };
        if (el.type !== 'element') return { code: 'not-an-element', node: op.node };
        const existing = el.children;
        const reuse = existing.length === 1 && existing[0].type === 'text' ? existing[0].id : null;
        const textId = reuse ?? tree.mintId();
        tree.replaceChildren(
          op.node,
          op.text === '' ? [] : [{ type: 'text', id: textId, value: op.text }]
        );
        return null;
      }
      case 'insertNode': {
        const tree = this.#tree(op.part);
        if (!(tree instanceof WorkingTree)) return tree;
        const loc = tree.get(op.parent);
        if (!loc) return { code: 'unknown-node', part: op.part, node: op.parent };
        if (loc.node.type !== 'element') return { code: 'not-an-element', node: op.parent };
        if (!validIndex(op.index, loc.node.children.length))
          return { code: 'invalid-index', index: op.index };
        const height = checkSpec(op.spec, MAX_XML_DEPTH - tree.level(op.parent));
        if (typeof height !== 'number') return height;
        const counter = { next: tree.nextId };
        const node = materialize(op.spec, tree.childContext(op.parent), counter);
        tree.nextId = counter.next;
        tree.insertChild(op.parent, op.index, node);
        return null;
      }
      case 'removeNode': {
        const tree = this.#tree(op.part);
        if (!(tree instanceof WorkingTree)) return tree;
        const loc = tree.get(op.node);
        if (!loc) return { code: 'unknown-node', part: op.part, node: op.node };
        if (loc.parent === null) return { code: 'root-node', node: op.node };
        tree.remove(op.node);
        return null;
      }
      case 'moveNode':
        return this.#move(op);
      case 'addPart':
        return this.#addPart(op);
      case 'removePart': {
        const parts = this.#parts ?? this.#base.parts;
        if (!parts.has(op.part)) return { code: 'unknown-part', part: op.part };
        this.#partsForWrite().delete(op.part);
        this.#trees.delete(op.part);
        this.#keys?.delete(partNameKey(op.part));
        return null;
      }
    }
  }

  #move(op: Extract<PptxOp, { op: 'moveNode' }>): OpRejection | null {
    const tree = this.#tree(op.part);
    if (!(tree instanceof WorkingTree)) return tree;
    const loc = tree.get(op.node);
    if (!loc) return { code: 'unknown-node', part: op.part, node: op.node };
    if (loc.parent === null) return { code: 'root-node', node: op.node };
    const target = tree.get(op.parent);
    if (!target) return { code: 'unknown-node', part: op.part, node: op.parent };
    if (target.node.type !== 'element') return { code: 'not-an-element', node: op.parent };
    if (tree.isSelfOrAncestor(op.node, op.parent)) return { code: 'cycle', node: op.node };
    const node = loc.node;
    if (node.type === 'element') {
      // The moved subtree must still fit under the reader's depth cap at its new place.
      const deepest = tree.level(op.parent) + subtreeHeight(node);
      if (deepest > MAX_XML_DEPTH)
        return { code: 'too-deep', detail: 'the node would be deeper than the reader accepts' };
    }
    const sameParent = loc.parent === op.parent;
    const length = target.node.children.length - (sameParent ? 1 : 0);
    if (!validIndex(op.index, length)) return { code: 'invalid-index', index: op.index };
    tree.remove(op.node);
    // A node's kind depends on its parent: derive it again, as a reopen would.
    const placed = node.type === 'element' ? rekind(node, tree.childContext(op.parent)) : node;
    tree.insertChild(op.parent, op.index, placed);
    return null;
  }

  #addPart(op: Extract<PptxOp, { op: 'addPart' }>): OpRejection | null {
    if (!isCanonicalPartName(op.part)) return { code: 'unsafe-part-name', part: op.part };
    // The writer folds ASCII case, so `/ppt/media/Image1.png` and `/ppt/media/image1.png` are
    // one ZIP entry: refuse a name OPC-equivalent to an existing part.
    if (!this.#keys) {
      this.#keys = new Set();
      for (const name of (this.#parts ?? this.#base.parts).keys())
        this.#keys.add(partNameKey(name));
    }
    const key = partNameKey(op.part);
    if (this.#keys.has(key)) return { code: 'part-exists', part: op.part };
    let part: PackagePart;
    if (op.xml !== undefined) {
      const tree = parseXmlPart(op.part, op.xml);
      if (!tree.ok) return { code: 'invalid-part-content', part: op.part, detail: tree.reason };
      part = { kind: 'xml', partName: op.part, xml: tree.part };
    } else {
      const bytes = op.base64 === undefined ? null : base64ToBytes(op.base64);
      if (!bytes) return { code: 'invalid-part-content', part: op.part, detail: 'base64' };
      part = { kind: 'binary', partName: op.part, bytes };
    }
    this.#partsForWrite().set(op.part, part);
    this.#keys.add(key);
    return null;
  }

  finish(): PptxPackage {
    for (const [name, tree] of this.#trees)
      if (tree.changed)
        this.#partsForWrite().set(name, { kind: 'xml', partName: name, xml: tree.result() });
    return this.#parts ? { parts: this.#parts.seal() } : this.#base;
  }
}

/** A transaction left open across ops: each op checks and applies at once. @internal */
export interface OpenTransaction {
  /** Apply one op; a rejection leaves the transaction unusable. */
  apply(op: PptxOp): OpRejection | null;
  /** The id the next node inserted into `part` gets, or undefined for no such XML part. */
  nextId(part: string): number | undefined;
  /** The element `id` of `part` as the applied ops left it, or undefined. */
  element(part: string, id: string): Element | undefined;
  /** The package after every applied op. The transaction is done after it. */
  finish(): PptxPackage;
}

/**
 * Open a transaction on `pkg` that ops apply to one at a time. A part's id index is copied once
 * per transaction, not once per op, so n ops on one part cost O(n + part), not O(n × part).
 * Pure: never mutates `pkg`. @internal
 */
export function openTransaction(pkg: PptxPackage): OpenTransaction {
  const tx = new Transaction(pkg);
  return {
    apply: (op) => tx.apply(op),
    nextId: (part) => tx.nextId(part),
    element: (part, id) => tx.element(part, id),
    finish: () => tx.finish(),
  };
}

/** Apply ops in order against one working copy. All or nothing. Pure: never mutates `pkg`. */
export function applyOps(pkg: PptxPackage, ops: readonly PptxOp[]): OpResult {
  if (!Array.isArray(ops))
    return { ok: false, rejection: { code: 'malformed-op', detail: 'ops is not an array' } };
  const tx = new Transaction(pkg);
  for (const op of ops) {
    const rejection = tx.apply(op);
    if (rejection) return { ok: false, rejection };
  }
  return { ok: true, pkg: tx.finish() };
}

/** Apply one op. Pure: returns a new package or a typed rejection. */
export function applyOp(pkg: PptxPackage, op: PptxOp): OpResult {
  return applyOps(pkg, [op]);
}
