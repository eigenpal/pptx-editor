// Shape removal that leaves no reference to a removed shape.
//
// Two parts of a slide name a shape by its `p:cNvPr/@id`: the slide's animations (ECMA-376
// Part 1, §19.5 `p:timing`, `spid`) and connector ends (`a:stCxn` and `a:endCxn`, `CT_Connection`
// in `CT_NonVisualConnectorProperties`). An edit that removes shapes removes, in the same transaction, the timing entries
// that name them or a shape inside them (`timingRemovals` in `store/package/timing.ts` holds the
// rules), and the connector ends glued to them. A connector whose end goes stays where it is,
// unglued: `CT_NonVisualConnectorProperties` makes both ends optional. A new shape takes the
// next free id, which can be the id of a removed shape, so a kept end would glue the connector
// to an unrelated shape. Only ids that the edit removed count: a reference that a file already
// left dangling stays as the file wrote it.

import {
  attr,
  find,
  nodeIndex,
  NS,
  timingRemovals,
  xmlPartOf,
  type Element,
} from '../../store/index.ts';
import type { OpBatch } from './op-batch.ts';

const MAX_DEPTH = 64;

/** The drawing ids of a slide's shape tree, and its connector ends. */
interface Drawing {
  /** Every `p:cNvPr/@id`, groups and `mc:AlternateContent` included. */
  readonly ids: Set<string>;
  /** Every `a:stCxn` and `a:endCxn`. */
  readonly ends: Element[];
}

function drawingOf(batch: OpBatch, part: string): Drawing {
  const ids = new Set<string>();
  const ends: Element[] = [];
  const walk = (el: Element, depth: number): void => {
    if (depth > MAX_DEPTH) return;
    if (el.ns === NS.p && el.local === 'cNvPr') {
      const id = attr(el, 'id');
      if (id !== undefined) ids.add(id);
      return;
    }
    if (el.ns === NS.a && (el.local === 'stCxn' || el.local === 'endCxn')) {
      ends.push(el);
      return;
    }
    // No `p:cNvPr` lives in text or in a frame's graphic data.
    if ((el.ns === NS.p && el.local === 'txBody') || (el.ns === NS.a && el.local === 'graphic'))
      return;
    for (const c of el.children) if (c.type === 'element') walk(c, depth + 1);
  };
  const tree = find(xmlPartOf(batch.pkg, part)?.root, 'p:spTree');
  if (tree) walk(tree, 0);
  return { ids, ends };
}

/**
 * Remove shapes (`p:sp`, `p:pic`, `p:grpSp`, `p:cxnSp`, `p:graphicFrame`, by node id) from a
 * slide part, with the animation entries and connector ends that name them or a shape inside
 * them. Every command that removes a shape removes it here. A shape moved out of a removed
 * group (ungroup) keeps its id, so it keeps its animations and its connectors.
 */
export function removeShapes(batch: OpBatch, part: string, ids: readonly string[]): void {
  if (ids.length === 0) return;
  const before = drawingOf(batch, part).ids;
  for (const id of lastFirst(batch, part, ids)) batch.push({ op: 'removeNode', part, node: id });
  const root = xmlPartOf(batch.pkg, part)?.root;
  if (!root) return;
  const now = drawingOf(batch, part);
  const gone = new Set([...before].filter((id) => !now.ids.has(id)));
  if (gone.size === 0) return;
  for (const end of now.ends)
    if (gone.has(attr(end, 'id') ?? '')) batch.push({ op: 'removeNode', part, node: end.id });
  for (const el of timingRemovals(root, gone)) batch.push({ op: 'removeNode', part, node: el.id });
}

/**
 * `ids` in removal order. Shapes of one parent go from the last place to the first: a removal
 * then leaves the places of the shapes still to remove as they were, and removing thousands of
 * shapes stays linear. Shapes of several parents keep their order.
 */
function lastFirst(batch: OpBatch, part: string, ids: readonly string[]): readonly string[] {
  const xml = xmlPartOf(batch.pkg, part);
  if (!xml || ids.length < 2) return ids;
  const index = nodeIndex(xml);
  const locs = ids.map((id) => index.get(id));
  const parent = locs[0]?.parent;
  if (locs.some((l) => !l || l.parent !== parent)) return ids;
  return ids
    .map((id, i) => ({ id, at: locs[i]!.index }))
    .sort((a, b) => b.at - a.at)
    .map((x) => x.id);
}
