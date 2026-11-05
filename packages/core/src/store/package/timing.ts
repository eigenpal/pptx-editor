// The animation targets of a slide (ECMA-376 Part 1, §19.5 `p:timing`), and the timing entries
// that must go when shapes go.
//
// `p:timing` names a shape by its `p:cNvPr/@id` in a `spid` attribute: `p:spTgt` (§19.5.72,
// `CT_TLShapeTargetElement`), its `p:subSp`, `p:inkTgt` (`CT_TLSubShapeId`), and the build
// entries `p:bldP`, `p:bldDgm`, `p:bldOleChart`, and `p:bldGraphic` (`AG_TLBuild`).
// `timingRemovals` lists what goes when the shapes of `gone` go, so the rest stays valid
// against `CT_SlideTiming`:
//
// - A time node (a child of `p:tnLst`, `p:childTnLst`, or `p:subTnLst`; `CT_TimeNodeList`) that
//   names a removed shape outside its own child lists. A behavior (`p:set`, `p:anim`, …) and a
//   media node must have exactly one `p:tgtEl` (`CT_TLCommonBehaviorData`,
//   `CT_TLCommonMediaNodeData`), so the target cannot go alone. A container whose `p:cond`
//   triggers on the shape (an interactive sequence) goes too: its trigger is gone.
// - A time node whose condition (`p:stCondLst`, `p:endCondLst`, `p:endSync`, or a sequence's
//   `p:prevCondLst` and `p:nextCondLst`) names (`p:tn`, `CT_TLTriggerTimeNodeID`) the
//   `p:cTn/@id` of a time node that went, or of a time node inside it: it waits for an event
//   that can never come. This repeats until nothing more goes.
// - A container whose `p:childTnLst` has no time node left. `p:childTnLst` is optional in
//   `CT_TLCommonTimeNodeData`, but a click step or sequence with no effect plays nothing, and
//   `CT_TimeNodeList` needs at least one child. A `p:subTnLst` with none left goes alone.
// - A build entry (a child of `p:bldLst`, `CT_BuildList`) that names a removed shape anywhere
//   inside it (its `p:tmplLst` too), and one whose effects all went: the entry pairs with the
//   effects (`p:cTn` with the same `grpId`) that target its `spid`.
// - `p:tnLst` and `p:bldLst` once empty (each needs one child), and the whole `p:timing` once
//   neither is left: a `p:extLst` alone describes nothing.
// - In a top-level `mc:AlternateContent` (ECMA-376 Part 3, §10.2.1), each branch's `p:timing`
//   by the same rules, and the whole `mc:AlternateContent` once every branch is empty.
//
// The timing is attacker-controlled: one pass over it builds the time nodes, their lists, and an
// index from each `p:cTn/@id` to the nodes that wait for it, and a worklist (not recursion)
// removes the nodes. The cost is linear in the size of `p:timing`, and a long `p:tn` chain
// cannot exhaust the stack.
//
// A removal takes only what it empties: a list that the file left empty stays. The writer's
// repair (`repairTimingRoot`) takes the lists that are already empty, by the same rules with no
// shape gone. Two collaborators who each delete one of two animated shapes leave such a list:
// each replica kept the list for the other shape's step, and the merge removes both steps. The
// repair touches only a `p:timing` that the session changed (an edit, or a collaborator's
// different copy): a `p:timing` exactly as the file wrote it saves unchanged, empty lists
// included (`markFileTimings`, `carryFileTimings`). A `spid` that names no shape is not
// repaired: only an edit that removes the shape removes its entries.

import { MAX_XML_DEPTH } from '@eigenpal/ooxml-core';
import { NS } from './namespaces.ts';
import { attr, child, elements } from './query.ts';
import type { Element, Node } from './tree.ts';

const NONE: ReadonlySet<string> = new Set();

const isP = (el: Element, local: string): boolean => el.ns === NS.p && el.local === local;
const isAlternate = (el: Element): boolean => el.ns === NS.mc && el.local === 'AlternateContent';
const isBranch = (el: Element): boolean =>
  el.ns === NS.mc && (el.local === 'Choice' || el.local === 'Fallback');

/** The slide's `p:timing` elements: its own, and each one in a top-level `mc:AlternateContent`. */
function timingsOf(root: Element | undefined): Element[] {
  const out: Element[] = [];
  const own = child(root, 'p:timing');
  if (own) out.push(own);
  for (const ac of elements(root).filter(isAlternate))
    for (const branch of elements(ac).filter(isBranch))
      for (const t of elements(branch)) if (isP(t, 'timing')) out.push(t);
  return out;
}

// ---------------------------------------------------------------- file timings

/**
 * The `p:timing` elements as the reader built them from a file. The tree is persistent, so an
 * edit inside a `p:timing` makes a new element and an edit elsewhere on the slide keeps it.
 */
const fileTimings = new WeakSet<Element>();

/** Record the `p:timing` elements of a slide root that the reader built from a file. */
export function markFileTimings(root: Element): void {
  for (const t of timingsOf(root)) fileTimings.add(t);
}

/**
 * Keep the file mark of a slide's `p:timing` when a collaborator's copy of the part replaces
 * it with the same content, so each replica saves the same bytes. Only a `p:timing` at the same
 * place that equals the marked one keeps the mark.
 */
export function carryFileTimings(before: Element, after: Element): void {
  const old = timingsOf(before);
  const next = timingsOf(after);
  if (old.length !== next.length) return;
  for (let i = 0; i < old.length; i += 1)
    if (fileTimings.has(old[i]) && old[i] !== next[i] && sameTree(old[i], next[i], 0))
      fileTimings.add(next[i]);
}

/** Whether two trees hold the same XML, node ids and attribute order aside. */
function sameTree(a: Node, b: Node, depth: number): boolean {
  if (a.type !== b.type || depth > MAX_XML_DEPTH) return false;
  if (a.type === 'text' || b.type === 'text')
    return a.type === 'text' && b.type === 'text' && a.value === b.value;
  if (a.ns !== b.ns || a.local !== b.local || a.attrs.length !== b.attrs.length) return false;
  if (a.children.length !== b.children.length) return false;
  if (a.attrs.length > 0) {
    const values = new Map(b.attrs.map((y) => [`${y.ns} ${y.local}`, y.value]));
    for (const x of a.attrs) if (values.get(`${x.ns} ${x.local}`) !== x.value) return false;
  }
  for (let i = 0; i < a.children.length; i += 1)
    if (!sameTree(a.children[i], b.children[i], depth + 1)) return false;
  return true;
}

// ---------------------------------------------------------------- public helpers

/**
 * The `p:cNvPr/@id` values that a slide's animations name (§19.5.72 `spTgt`, §19.5.12 `bldP`):
 * in `p:timing`, and in a `p:timing` inside top-level `mc:AlternateContent` (every `mc:Choice`
 * and the `mc:Fallback`, ECMA-376 Part 3, §10.2.1), where a writer puts animations that need an
 * extension. @internal
 */
export function animatedShapeIds(root: Element | undefined): Set<string> {
  const ids = new Set<string>();
  const walk = (el: Element, depth: number): void => {
    if (depth > MAX_XML_DEPTH) return;
    const spid = attr(el, 'spid');
    if (spid !== undefined) ids.add(spid);
    for (const c of elements(el)) walk(c, depth + 1);
  };
  for (const t of timingsOf(root)) walk(t, 0);
  return ids;
}

/**
 * The elements of a slide root to remove when the shapes whose `p:cNvPr/@id` is in `gone` go
 * (see the file comment): parts of `p:timing`, a whole `p:timing`, or a whole top-level
 * `mc:AlternateContent`. None of them is inside another. @internal
 */
export function timingRemovals(root: Element, gone: ReadonlySet<string>): Element[] {
  return removalsIn(root, gone, () => true, false);
}

/**
 * A slide root as the writer saves it: each `p:timing` that the session changed loses the time
 * node lists and build lists that are empty, by the rules of `timingRemovals` with no shape
 * gone. The root itself when nothing changes.
 */
export function repairTimingRoot(root: Element): Element {
  const timings = timingsOf(root);
  if (timings.every((t) => fileTimings.has(t))) return root;
  const remove = new Set(removalsIn(root, NONE, (t) => !fileTimings.has(t), true));
  if (remove.size === 0) return root;
  const strip = (el: Element, depth: number): Element => {
    if (depth > MAX_XML_DEPTH) return el;
    let changed = false;
    const kids: Node[] = [];
    for (const c of el.children) {
      if (c.type === 'element' && remove.has(c)) {
        changed = true;
        continue;
      }
      const next = c.type === 'element' ? strip(c, depth + 1) : c;
      if (next !== c) changed = true;
      kids.push(next);
    }
    return changed ? { ...el, children: kids } : el;
  };
  return strip(root, 0);
}

/**
 * `timingRemovals` over the `p:timing` elements that `include` accepts. With `repair`, the lists
 * that are already empty go too; without it, only what the removal of `gone` empties goes.
 */
function removalsIn(
  root: Element,
  gone: ReadonlySet<string>,
  include: (timing: Element) => boolean,
  repair: boolean
): Element[] {
  const out: Element[] = [];
  const own = child(root, 'p:timing');
  if (own && include(own)) topmost(own, deadIn(own, gone, repair), out);
  for (const ac of elements(root).filter(isAlternate)) {
    const branches = elements(ac).filter(isBranch);
    const timings = branches.flatMap((b) => elements(b).filter((t) => isP(t, 'timing')));
    if (timings.length === 0) continue;
    const dead = new Set<Element>();
    for (const t of timings) if (include(t)) for (const el of deadIn(t, gone, repair)) dead.add(el);
    if (dead.size === 0) continue;
    if (branches.every((b) => elements(b).every((c) => dead.has(c)))) out.push(ac);
    else for (const t of timings) topmost(t, dead, out);
  }
  return out;
}

/** Add to `out` the marked elements under `root` (itself included) with no marked ancestor. */
function topmost(root: Element, dead: ReadonlySet<Element>, out: Element[]): void {
  const walk = (el: Element, depth: number): void => {
    if (dead.has(el)) {
      out.push(el);
      return;
    }
    if (depth > MAX_XML_DEPTH) return;
    for (const c of elements(el)) walk(c, depth + 1);
  };
  walk(root, 0);
}

// ---------------------------------------------------------------- the time node graph

const TIME_LISTS: ReadonlySet<string> = new Set(['tnLst', 'childTnLst', 'subTnLst']);
const isTimeList = (el: Element): boolean => el.ns === NS.p && TIME_LISTS.has(el.local);

/** A time node's `p:cTn`: its own, or the one in its `p:cBhvr` or `p:cMediaNode`. */
const cTnOf = (node: Element): Element | undefined =>
  child(node, 'p:cTn') ??
  child(child(node, 'p:cBhvr'), 'p:cTn') ??
  child(child(node, 'p:cMediaNode'), 'p:cTn');

/** One time node list: `p:tnLst`, or a node's `p:childTnLst` or `p:subTnLst`. */
interface TimeList {
  readonly el: Element;
  /** The node whose `p:cTn` holds the list; none for `p:tnLst`. */
  readonly owner: TimeNode | undefined;
  /** Its nodes that have not gone. */
  alive: number;
}

/** One time node, and what its own content (outside its time node lists) names. */
interface TimeNode {
  readonly el: Element;
  readonly list: TimeList;
  /** The `p:cTn/@id` of the node. */
  readonly id: string | undefined;
  /** The `grpId` of the nearest effect (`p:cTn` with a `grpId`): the node or one around it. */
  readonly effect: string | undefined;
  /** The nodes in its `p:childTnLst` and `p:subTnLst`. */
  readonly kids: TimeNode[];
  /** The `spid` values in its own content. */
  readonly spids: string[];
  gone: boolean;
  /** Its id and the ids of the nodes inside it count as gone. */
  released: boolean;
  /** It goes, or an element around it goes. */
  lost: boolean;
}

/** The time nodes of one `p:tnLst`, in document order, and their lists. */
interface TimeGraph {
  readonly nodes: TimeNode[];
  readonly lists: TimeList[];
  /** For each `p:cTn/@id`, the nodes whose own content names it in a `p:tn`. */
  readonly waiters: Map<string, TimeNode[]>;
}

function readGraph(tnLst: Element | undefined): TimeGraph {
  const graph: TimeGraph = { nodes: [], lists: [], waiters: new Map() };
  const { nodes, lists, waiters } = graph;
  // What a node's own content names: every element under it outside its time node lists.
  const scanOwn = (el: Element, n: TimeNode, depth: number): void => {
    if (depth > MAX_XML_DEPTH) return;
    const spid = attr(el, 'spid');
    if (spid !== undefined) n.spids.push(spid);
    if (isP(el, 'tn')) {
      const val = attr(el, 'val');
      const at = val === undefined ? undefined : waiters.get(val);
      if (at) at.push(n);
      else if (val !== undefined) waiters.set(val, [n]);
    }
    for (const c of elements(el)) if (!isTimeList(c)) scanOwn(c, n, depth + 1);
  };
  const readList = (
    el: Element,
    owner: TimeNode | undefined,
    effect: string | undefined,
    depth: number
  ): void => {
    const list: TimeList = { el, owner, alive: 0 };
    lists.push(list);
    if (depth > MAX_XML_DEPTH) return;
    for (const nodeEl of elements(el)) {
      const cTn = cTnOf(nodeEl);
      const n: TimeNode = {
        el: nodeEl,
        list,
        id: attr(cTn, 'id'),
        effect: attr(cTn, 'grpId') ?? effect,
        kids: [],
        spids: [],
        gone: false,
        released: false,
        lost: false,
      };
      list.alive += 1;
      nodes.push(n);
      owner?.kids.push(n);
      scanOwn(nodeEl, n, depth + 1);
      for (const sub of [child(cTn, 'p:childTnLst'), child(cTn, 'p:subTnLst')])
        if (sub) readList(sub, n, n.effect, depth + 3);
    }
  };
  if (tnLst) readList(tnLst, undefined, undefined, 0);
  return graph;
}

/**
 * Mark in `dead` the time nodes and lists that go when the shapes in `gone` go. A worklist,
 * not recursion, follows the removals, so a long chain of `p:tn` waits uses no stack, and each
 * node and each `p:tn` is handled once.
 */
function removeTimeNodes(
  graph: TimeGraph,
  gone: ReadonlySet<string>,
  repair: boolean,
  dead: Set<Element>
): void {
  const work: TimeNode[] = [];
  const emptied = (list: TimeList): void => {
    if (list.owner && list.el.local === 'childTnLst') work.push(list.owner);
    else dead.add(list.el);
  };
  const release = (top: TimeNode): void => {
    const stack = [top];
    while (stack.length > 0) {
      const n = stack.pop()!;
      if (n.released) continue;
      n.released = true;
      if (n.id !== undefined) for (const w of graph.waiters.get(n.id) ?? []) work.push(w);
      for (const k of n.kids) stack.push(k);
    }
  };
  if (repair) for (const list of graph.lists) if (list.alive === 0) emptied(list);
  for (const n of graph.nodes) if (n.spids.some((s) => gone.has(s))) work.push(n);
  while (work.length > 0) {
    const n = work.pop()!;
    if (n.gone) continue;
    n.gone = true;
    dead.add(n.el);
    release(n);
    n.list.alive -= 1;
    if (n.list.alive === 0) emptied(n.list);
  }
  // Document order puts each node after the node that holds it.
  for (const n of graph.nodes)
    n.lost = n.gone || dead.has(n.list.el) || (n.list.owner?.lost ?? false);
}

/** `spid grpId` of each effect target: of every node, or of the nodes that stay. */
function effectTargets(graph: TimeGraph, staying: boolean): Set<string> {
  const out = new Set<string>();
  for (const n of graph.nodes)
    if (n.effect !== undefined && !(staying && n.lost))
      for (const spid of n.spids) out.add(`${spid} ${n.effect}`);
  return out;
}

/** Whether `el` or an element inside it has a `spid` in `gone`. */
function spidIn(el: Element, gone: ReadonlySet<string>, depth: number): boolean {
  if (depth > MAX_XML_DEPTH) return false;
  const spid = attr(el, 'spid');
  if (spid !== undefined && gone.has(spid)) return true;
  return elements(el).some((c) => spidIn(c, gone, depth + 1));
}

/**
 * The elements of one `p:timing` (itself included) that go when the shapes in `gone` go, and
 * with `repair` the lists that are already empty.
 */
function deadIn(timing: Element, gone: ReadonlySet<string>, repair: boolean): Set<Element> {
  const dead = new Set<Element>();
  const tnLst = child(timing, 'p:tnLst');
  const graph = readGraph(tnLst);
  removeTimeNodes(graph, gone, repair, dead);
  const bldLst = child(timing, 'p:bldLst');
  const builds = elements(bldLst);
  // A build entry pairs with the effects that target its `spid` with its `grpId`: when every
  // one of them went, the entry builds nothing. One with no effect in the file stays.
  const before = builds.length > 0 && dead.size > 0 ? effectTargets(graph, false) : undefined;
  const after = before ? effectTargets(graph, true) : undefined;
  for (const b of builds) {
    const pair = `${attr(b, 'spid') ?? ''} ${attr(b, 'grpId') ?? ''}`;
    if (spidIn(b, gone, 0) || (before?.has(pair) && !after?.has(pair))) dead.add(b);
  }
  if (bldLst && (builds.length > 0 || repair) && builds.every((b) => dead.has(b))) dead.add(bldLst);
  if (dead.size > 0 && (!tnLst || dead.has(tnLst)) && (!bldLst || dead.has(bldLst)))
    dead.add(timing);
  return dead;
}
