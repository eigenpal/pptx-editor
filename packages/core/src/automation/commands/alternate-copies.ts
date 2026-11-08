// The copies of a shape that an edit leaves out of date in `mc:AlternateContent`.
//
// A writer puts one object in several branches (ECMA-376 Part 3, 10.2.1): each `mc:Choice` is
// for a reader that understands the namespaces it requires, and `mc:Fallback` is for any other
// reader. The editor draws and edits one copy, the branch of `alternateBranch`. When that copy is
// in a `mc:Choice`, a reader that selects another branch would show the object as it was before
// the edit. So the transaction of the edit also removes each copy in another branch that the edit
// did not change too. Copies match by kind and `p:cNvPr/@id`. In nested `mc:AlternateContent`, the
// rule holds at each level: the copies in the other branches of every enclosing one go too, also
// those inside a nested `mc:AlternateContent` of such a branch. Hide, rename, and the aspect lock
// change every copy (`arrange-commands.ts`), so those copies stay. When the edit deletes the
// drawn copy, every other copy goes too, whichever branch the editor draws, and a wrapper that
// keeps no content goes with them.
//
// An edit that only sets attributes of the copy's transform (`a:xfrm` in `p:spPr` or
// `p:grpSpPr`, and the `p:xfrm` of a frame: a move, a resize, a rotation, or a flip) sets the
// same attributes in each other copy that has the same element, in every branch. Those copies
// stay current, so no copy goes, and a reader of another branch shows the object where the
// editor shows it. An attribute write converges in a collaboration room as the edit does.
//
// Other edits remove, and do not copy, for two reasons. A copy of the `mc:Choice` shape can hold
// markup in the namespaces that the Choice requires, which a reader of `mc:Fallback` does not
// understand. And a removal converges in a collaboration room: when two peers edit the same shape,
// both remove the same node, but two copies would both stay. DOCX Editor rewrites the fallback of a
// text box from its DrawingML story on save
// (`packages/core/src/store/package/textbox-fallback-export.ts` at adf0e178). A PresentationML
// fallback is a whole shape that can differ from its Choice by design, so that rule does not carry
// over.
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

/**
 * A copy that an edit touched: the copy, its branch, and their `mc:AlternateContent`. In an
 * enclosing `mc:AlternateContent`, the copy is still the shape in the innermost branch.
 */
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

/** One step down a copy: an element name and its place among the siblings of that name. */
interface Step {
  readonly ns: string;
  readonly local: string;
  readonly nth: number;
}

/** The steps from `top` down to the element with node id `id`, or undefined. */
function pathTo(top: Element, id: string, depth = 0): Step[] | undefined {
  if (top.id === id) return [];
  if (depth > MAX_DEPTH) return undefined;
  const seen = new Map<string, number>();
  for (const c of elements(top)) {
    const key = `${c.ns} ${c.local}`;
    const nth = seen.get(key) ?? 0;
    seen.set(key, nth + 1);
    const rest = pathTo(c, id, depth + 1);
    if (rest) return [{ ns: c.ns, local: c.local, nth }, ...rest];
  }
  return undefined;
}

/** The element that `path` reaches from `top`, or undefined. */
function follow(top: Element, path: readonly Step[]): Element | undefined {
  let at: Element | undefined = top;
  for (const step of path) {
    if (!at) return undefined;
    at = elements(at).filter((c) => c.ns === step.ns && c.local === step.local)[step.nth];
  }
  return at;
}

/** Whether `path` reaches the transform of a shape, a group, or a frame, or an element in it. */
function isTransform(path: readonly Step[]): boolean {
  const [first, second] = path;
  if (!first) return false;
  if (first.ns === NS.p && first.local === 'xfrm') return true;
  return (
    first.ns === NS.p &&
    (first.local === 'spPr' || first.local === 'grpSpPr') &&
    second?.ns === NS.a &&
    second.local === 'xfrm'
  );
}

/**
 * The ops that set the same transform attributes in `twin` as `ops` set in `copy`, or
 * undefined when an op of `ops` inside `copy` is not a transform attribute or `twin` lacks
 * its element. `inside` holds the node ids of `copy`.
 */
function mirroredOps(
  copy: Element,
  twin: Element,
  ops: readonly PptxOp[],
  inside: ReadonlySet<string>
): PptxOp[] | undefined {
  const out: PptxOp[] = [];
  for (const op of ops) {
    if (!targetsOf(op).some((id) => inside.has(id))) continue;
    if (op.op !== 'setAttribute' || op.name.ns !== '') return undefined;
    const path = pathTo(copy, op.node);
    if (!path || !isTransform(path)) return undefined;
    const target = follow(twin, path);
    if (!target) return undefined;
    out.push({ ...op, node: target.id });
  }
  return out;
}

/** The node ids of `el` and its descendants. */
function idsOf(el: Element, out = new Set<string>(), depth = 0): Set<string> {
  out.add(el.id);
  if (depth < MAX_DEPTH)
    for (const c of el.children) if (c.type === 'element') idsOf(c, out, depth + 1);
  return out;
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
      // The shape that the last branch on the way up holds.
      let copy: Element | undefined;
      for (let depth = 0; at && depth < MAX_DEPTH; depth += 1) {
        const node = at.node;
        if (node.type !== 'element') {
          at = at.parent ? index.get(at.parent) : undefined;
          continue;
        }
        const up = at.parent ? index.get(at.parent) : undefined;
        const wrapper = up?.node;
        if (below && isBranch(node) && wrapper?.type === 'element' && isAlternateContent(wrapper)) {
          // A nested `mc:AlternateContent` is not a copy: the shape in it is.
          copy = copy && isAlternateContent(below) ? copy : below;
          const list = out.get(op.part) ?? [];
          list.push({ copy, branch: node, wrapper });
          out.set(op.part, list);
        }
        below = node;
        at = up;
      }
    }
  return out;
}

/** Whether `branch` draws nothing: no content, or only nested wrappers that draw nothing. */
function drawsNothing(branch: Element, depth = 0): boolean {
  return elements(branch).every(
    (el) =>
      isAlternateContent(el) &&
      depth < MAX_DEPTH &&
      elements(el).every((b) => !isBranch(b) || drawsNothing(b, depth + 1))
  );
}

/** The elements of `branch`, with the content of every branch of a nested `mc:AlternateContent`. */
function branchContent(branch: Element, depth = 0, out: Element[] = []): Element[] {
  for (const el of elements(branch)) {
    if (!isAlternateContent(el)) out.push(el);
    else if (depth < MAX_DEPTH)
      for (const inner of elements(el)) if (isBranch(inner)) branchContent(inner, depth + 1, out);
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
    const done = new Set<string>();
    for (const { copy, branch, wrapper } of list) {
      // Only an edit of the drawn copy makes the others out of date.
      // A copy in nested wrappers has one entry for each wrapper around it.
      const key = `${copy.id} ${wrapper.id}`;
      if (done.has(key) || alternateBranch(wrapper, root) !== branch) continue;
      done.add(key);
      const deleted = !exists(copy.id);
      const id = drawingId(copy);
      if (id === undefined) continue;
      const twins: Element[] = [];
      for (const other of elements(wrapper)) {
        if (other === branch || !isBranch(other)) continue;
        for (const twin of branchContent(other))
          if (
            twin.ns === copy.ns &&
            twin.local === copy.local &&
            drawingId(twin) === id &&
            !changed.has(twin.id) &&
            exists(twin.id)
          )
            twins.push(twin);
      }
      // A transform edit keeps every copy current.
      const inside = idsOf(copy);
      const mirrors = deleted ? [] : twins.map((t) => mirroredOps(copy, t, ops, inside));
      if (!deleted && mirrors.every((m) => m !== undefined)) {
        for (const m of mirrors) for (const op of m!) batch.push(op);
        continue;
      }
      if (!deleted && branch.local !== 'Choice') continue;
      for (const twin of twins) stale.add(twin.id);
      if (deleted) emptied.add(wrapper);
    }
    removeShapes(batch, part, [...stale]);
    // A wrapper whose branches all lost their content draws nothing for any reader.
    for (const wrapper of emptied) {
      const now = xmlPartOf(batch.pkg, part);
      const at = now ? nodeIndex(now).get(wrapper.id)?.node : undefined;
      if (at?.type !== 'element') continue;
      if (elements(at).every((b) => !isBranch(b) || drawsNothing(b)))
        batch.push({ op: 'removeNode', part, node: wrapper.id });
    }
  }
  return batch.ops;
}
