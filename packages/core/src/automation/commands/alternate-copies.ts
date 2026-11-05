// The copies of a shape that an edit leaves out of date in `mc:AlternateContent`.
//
// A writer puts one object in several branches (ECMA-376 Part 3, 10.2.1): each `mc:Choice` is
// for a reader that understands the namespaces it requires, and `mc:Fallback` is for any other
// reader. The editor draws and edits one copy, the branch of `alternateBranch`. When that copy is
// in a `mc:Choice`, a reader that selects another branch would show the object as it was before
// the edit. So the transaction of the edit also removes each copy in another branch that the edit
// did not change too. Copies match by kind and `p:cNvPr/@id`. Hide, rename, and the aspect lock
// change every copy (`arrange-commands.ts`), so those copies stay. When the edit deletes the
// drawn copy, every other copy goes too, whichever branch the editor draws, and a wrapper that
// keeps no content goes with them.
//
// The rule removes, and does not copy, for two reasons. A copy of the `mc:Choice` shape can hold
// markup in the namespaces that the Choice requires, which a reader of `mc:Fallback` does not
// understand. And a removal converges in a collaboration room: when two peers edit the same
// shape, both remove the same node, but two copies would both stay. DOCX Editor rewrites the
// fallback of a text box from its DrawingML story on save (`textbox-fallback-export.ts`). A
// PresentationML fallback is a whole shape that can differ from its Choice by design, so that
// rule does not carry over.
//
// The decision reads the package before the edit and the ops only, never editor state, so every
// replica that applies the same ops gets the same result.

import { alternateBranch, isAlternateContent } from '../../cascade/index.ts';
import {
  attr,
  child,
  elements,
  nodeIndex,
  NS,
  xmlPartOf,
  type Element,
  type PptxOp,
  type PptxPackage,
} from '../../store/index.ts';
import { OpBatch } from './op-batch.ts';
import { removeShapes } from './shape-removal.ts';

/** The deepest nesting the walk up from an edited node follows: the reader's XML depth cap. */
const MAX_DEPTH = 256;

const isBranch = (el: Element): boolean =>
  el.ns === NS.mc && (el.local === 'Choice' || el.local === 'Fallback');

/** A copy that an edit touched: the copy, its branch, and their `mc:AlternateContent`. */
interface Touched {
  readonly copy: Element;
  readonly branch: Element;
  readonly wrapper: Element;
}

/** The `p:cNvPr/@id` of a shape (`p:nvSpPr`, `p:nvPicPr`, and the other `p:nv…Pr`). */
function drawingId(el: Element): string | undefined {
  const nv = elements(el).find((c) => c.ns === NS.p && c.local.startsWith('nv'));
  return attr(child(nv, 'p:cNvPr'), 'id');
}

/** The node ids that an op edits, in the package before the edit. */
function targetsOf(op: PptxOp): string[] {
  switch (op.op) {
    case 'setAttribute':
    case 'setText':
    case 'removeNode':
      return [op.node];
    case 'insertNode':
      return [op.parent];
    case 'moveNode':
      return [op.node, op.parent];
    default:
      return [];
  }
}

/** The copies in `mc:AlternateContent` branches that the ops touch, by part. */
function touchedCopies(base: PptxPackage, ops: readonly PptxOp[]): Map<string, Touched[]> {
  const out = new Map<string, Touched[]>();
  const seen = new Set<string>();
  for (const op of ops)
    for (const id of targetsOf(op)) {
      const key = `${op.part}\u0000${id}`;
      if (seen.has(key)) continue;
      seen.add(key);
      const xml = xmlPartOf(base, op.part);
      if (!xml) continue;
      const index = nodeIndex(xml);
      // A node that an earlier op of the same edit created is not in the index: the op that
      // created it touched one of its ancestors.
      let at = index.get(id);
      let below: Element | undefined;
      for (let depth = 0; at && depth < MAX_DEPTH; depth += 1) {
        const node = at.node;
        if (node.type !== 'element') {
          at = at.parent ? index.get(at.parent) : undefined;
          continue;
        }
        const up = at.parent ? index.get(at.parent) : undefined;
        const wrapper = up?.node;
        if (below && isBranch(node) && wrapper?.type === 'element' && isAlternateContent(wrapper)) {
          const list = out.get(op.part) ?? [];
          list.push({ copy: below, branch: node, wrapper });
          out.set(op.part, list);
        }
        below = node;
        at = up;
      }
    }
  return out;
}

/**
 * `ops` with, at the end, the removal of each copy that they leave out of date in another
 * branch of an `mc:AlternateContent` (the rules are at the top of this file). Returns `ops`
 * itself when no op touches a copy in a branch. @internal
 */
export function withoutStaleCopies(base: PptxPackage, ops: readonly PptxOp[]): readonly PptxOp[] {
  const touched = touchedCopies(base, ops);
  if (touched.size === 0) return ops;
  const batch = new OpBatch(base);
  for (const op of ops) batch.push(op);
  for (const [part, list] of touched) {
    const root = xmlPartOf(base, part)?.root;
    const after = xmlPartOf(batch.pkg, part);
    if (!root || !after) continue;
    const exists = (id: string): boolean => nodeIndex(after).has(id);
    const changed = new Set(list.map((t) => t.copy.id));
    const stale = new Set<string>();
    const emptied = new Set<Element>();
    for (const { copy, branch, wrapper } of list) {
      // Only an edit of the drawn copy makes the others out of date.
      if (alternateBranch(wrapper, root) !== branch) continue;
      const deleted = !exists(copy.id);
      if (!deleted && branch.local !== 'Choice') continue;
      const id = drawingId(copy);
      if (id === undefined) continue;
      for (const other of elements(wrapper)) {
        if (other === branch || !isBranch(other)) continue;
        for (const twin of elements(other))
          if (
            twin.ns === copy.ns &&
            twin.local === copy.local &&
            drawingId(twin) === id &&
            !changed.has(twin.id) &&
            exists(twin.id)
          )
            stale.add(twin.id);
      }
      if (deleted) emptied.add(wrapper);
    }
    removeShapes(batch, part, [...stale]);
    // A wrapper whose branches all lost their content draws nothing for any reader.
    for (const wrapper of emptied) {
      const now = xmlPartOf(batch.pkg, part);
      const at = now ? nodeIndex(now).get(wrapper.id)?.node : undefined;
      if (at?.type !== 'element') continue;
      if (elements(at).every((b) => !isBranch(b) || elements(b).length === 0))
        batch.push({ op: 'removeNode', part, node: wrapper.id });
    }
  }
  return batch.ops;
}
