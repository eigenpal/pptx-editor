// Structure edits of a diagram's data model (ECMA-376 Part 1, 21.4.3): add a node before or
// after a node at its level or as its last child, and delete a node with its subtree.
//
// A new node is a content point with an empty text body, a parent transition point and a sibling
// transition point (`parTrans`, `sibTrans`), and the parent-of connection (`dgm:cxn`) that names
// all three. The connections of one parent number their children from 0 (`srcOrd`), in order.
// A new node gets no presentation points: a presentation application builds them from the
// layout definition when it lays the diagram out. A deleted node takes its subtree, their
// transition points and connections, and the presentation points and connections that belong to
// them, so no point or connection names a point that is gone.

import { attr, child, children, elements, type Element } from '../../store/index.ts';
import {
  isContentPoint,
  type DiagramModel,
  readDiagramModel,
} from '../../cascade/diagram-model.ts';
import { CommandRefusal } from './op-batch.ts';
import { dgm, dml, mapElements, withAttr, withChildren } from './diagram-xml.ts';

/** Where a new node goes, next to the target node. @public */
export type DiagramInsertPosition = 'after' | 'before' | 'below';

/** The most points one data model may hold after an edit. */
export const MAX_DIAGRAM_POINTS = 4096;

/** The largest connection list an edit reads. */
const MAX_DIAGRAM_CONNECTIONS = 4 * MAX_DIAGRAM_POINTS;

/** Refuse a data model larger than an edit reads. */
function checkSize(model: ReturnType<typeof readDiagramModel>, added = 0): void {
  if (
    model.points.size + added > MAX_DIAGRAM_POINTS ||
    model.connections.length > MAX_DIAGRAM_CONNECTIONS
  )
    throw new CommandRefusal('errors.diagram.tooLarge');
}

/** FNV-1a over a string, from `seed`. */
function fnv(text: string, seed: number): number {
  let h = seed >>> 0;
  for (let i = 0; i < text.length; i += 1) {
    h ^= text.charCodeAt(i);
    h = Math.imul(h, 0x01000193) >>> 0;
  }
  return h >>> 0;
}

/**
 * New model ids (`{8-4-4-4-12}` GUIDs) that no point or connection of the model uses. They
 * derive from the model's ids, so the same edit of the same model gives the same ids: a saved
 * copy is the same package each time.
 */
export function idSource(model: DiagramModel): () => string {
  const used = new Set<string>([...model.points.keys(), ...model.connections.map((c) => c.id)]);
  const seed = [...used].join('|');
  let n = 0;
  return () => {
    for (;;) {
      n += 1;
      const hex = [0x811c9dc5, 0x01000193, 0x9e3779b9, 0x85ebca6b]
        .map((s) =>
          fnv(`${seed}#${n}`, s ^ n)
            .toString(16)
            .padStart(8, '0')
        )
        .join('')
        .toUpperCase();
      const id = `{${hex.slice(0, 8)}-${hex.slice(8, 12)}-4${hex.slice(13, 16)}-8${hex.slice(17, 20)}-${hex.slice(20, 32)}}`;
      if (!used.has(id)) {
        used.add(id);
        return id;
      }
    }
  };
}

/** The language of a point's first run or end of paragraph, for a new point's text. */
function langOf(point: Element | undefined): string | undefined {
  for (const p of children(child(point, 'dgm:t'), 'a:p'))
    for (const r of elements(p)) {
      const lang = attr(r.local === 'endParaRPr' ? r : child(r, 'a:rPr'), 'lang');
      if (lang) return lang;
    }
  return undefined;
}

/** An empty text body, as a data point holds it. */
const emptyText = (lang: string | undefined): Element =>
  dgm('t', {}, [
    dml('bodyPr'),
    dml('lstStyle'),
    dml('p', {}, [dml('endParaRPr', lang ? { lang } : {})]),
  ]);

/** A point with no properties of its own. */
const point = (attrs: Record<string, string>, lang: string | undefined): Element =>
  dgm('pt', attrs, [dgm('prSet'), dgm('spPr'), emptyText(lang)]);

/** `dataModel` with its `ptLst` and `cxnLst` mapped. */
function withLists(
  dataModel: Element,
  pts: (list: Element) => Element,
  cxns: (list: Element) => Element
): Element {
  return mapElements(dataModel, (e) =>
    e.local === 'ptLst' ? pts(e) : e.local === 'cxnLst' ? cxns(e) : e
  );
}

/** Set the `srcOrd` of the connections in `parentLinks` (connection id to order). */
function renumber(cxnLst: Element, parentLinks: ReadonlyMap<string, number>): Element {
  return mapElements(cxnLst, (c) => {
    const id = attr(c, 'modelId');
    const ord = id === undefined ? undefined : parentLinks.get(id);
    return ord === undefined ? c : withAttr(c, 'srcOrd', String(ord));
  });
}

/** The content node `id`, or a refusal. */
function contentNode(model: DiagramModel, id: string): void {
  if (!isContentPoint(model.points.get(id))) throw new CommandRefusal('errors.diagram.noNode');
}

/** Add a node next to `target` (`position`); returns the new data model and the new node. */
export function insertDiagramNode(
  dataModel: Element,
  target: string,
  position: DiagramInsertPosition
): { readonly dataModel: Element; readonly node: string } {
  const model = readDiagramModel(dataModel);
  contentNode(model, target);
  checkSize(model, 3);
  const link = model.parentLink.get(target);
  const parent = position === 'below' ? target : link?.src;
  if (!parent) throw new CommandRefusal('errors.diagram.noNode');
  const siblings = [...model.childrenOf(parent)];
  const at =
    position === 'below'
      ? siblings.length
      : siblings.indexOf(target) + (position === 'after' ? 1 : 0);
  const next = idSource(model);
  const [node, parTrans, sibTrans, cxn] = [next(), next(), next(), next()];
  siblings.splice(at, 0, node);
  const lang = langOf(model.points.get(target)?.el);
  const added = [
    point({ modelId: node }, lang),
    point({ modelId: parTrans, type: 'parTrans', cxnId: cxn }, lang),
    point({ modelId: sibTrans, type: 'sibTrans', cxnId: cxn }, lang),
  ];
  const newLink = dgm('cxn', {
    modelId: cxn,
    srcId: parent,
    destId: node,
    srcOrd: String(at),
    destOrd: '0',
    parTransId: parTrans,
    sibTransId: sibTrans,
  });
  const order = new Map<string, number>();
  siblings.forEach((id, i) => {
    const c = model.parentLink.get(id);
    if (c) order.set(c.id, i);
  });
  const targetEl = model.points.get(target)!.el;
  const out = withLists(
    dataModel,
    (ptLst) => {
      const kids = [...ptLst.children];
      const i = kids.indexOf(targetEl);
      const at = position === 'before' ? i : i + 1;
      return withChildren(ptLst, [...kids.slice(0, at), ...added, ...kids.slice(at)]);
    },
    (cxnLst) => {
      const list = renumber(cxnLst, order);
      const kids = [...list.children];
      // After the last parent-of connection, ahead of the presentation connections.
      let i = kids.length;
      for (let k = kids.length - 1; k >= 0; k -= 1) {
        const c = kids[k];
        if (c?.type === 'element' && (attr(c, 'type') ?? 'parOf') === 'parOf') {
          i = k + 1;
          break;
        }
      }
      kids.splice(i, 0, newLink);
      return withChildren(list, kids);
    }
  );
  return { dataModel: out, node };
}

/** The subtree of `id` under parent-of connections, `id` first. */
export function subtreeOf(model: DiagramModel, id: string): string[] {
  const out: string[] = [];
  const queue = [id];
  const seen = new Set<string>();
  for (let i = 0; i < queue.length; i += 1) {
    const next = queue[i]!;
    if (seen.has(next)) continue;
    seen.add(next);
    out.push(next);
    for (const c of model.childrenOf(next)) queue.push(c);
  }
  return out;
}

/** Delete `target` and its subtree; returns the new data model and every removed point id. */
export function deleteDiagramNode(
  dataModel: Element,
  target: string
): { readonly dataModel: Element; readonly removed: ReadonlySet<string> } {
  const model = readDiagramModel(dataModel);
  contentNode(model, target);
  checkSize(model);
  const link = model.parentLink.get(target);
  if (!link) throw new CommandRefusal('errors.diagram.noNode');
  const content = (id: string) => isContentPoint(model.points.get(id));
  if (link.src === model.root && model.childrenOf(link.src).filter(content).length <= 1)
    throw new CommandRefusal('errors.diagram.lastShape');
  const removed = new Set<string>(subtreeOf(model, target));
  const cxns = new Set<string>();
  for (const id of [...removed]) {
    const c = model.parentLink.get(id);
    if (!c) continue;
    cxns.add(c.id);
    if (c.parTrans) removed.add(c.parTrans);
    if (c.sibTrans) removed.add(c.sibTrans);
  }
  for (const p of model.points.values()) {
    const cxnId = attr(p.el, 'cxnId');
    if ((p.type === 'parTrans' || p.type === 'sibTrans') && cxnId && cxns.has(cxnId))
      removed.add(p.id);
  }
  // Presentation points of removed points, and the presentation points under them.
  for (const p of model.points.values())
    if (p.type === 'pres' && removed.has(attr(child(p.el, 'dgm:prSet'), 'presAssocID') ?? ''))
      removed.add(p.id);
  const presKids = new Map<string, string[]>();
  for (const c of model.connections)
    if (c.type === 'presParOf') presKids.set(c.src, [...(presKids.get(c.src) ?? []), c.dest]);
  const queue = [...removed];
  for (let i = 0; i < queue.length; i += 1)
    for (const dest of presKids.get(queue[i]!) ?? [])
      if (!removed.has(dest)) {
        removed.add(dest);
        queue.push(dest);
      }
  const order = new Map<string, number>();
  model
    .childrenOf(link.src)
    .filter((id) => id !== target)
    .forEach((id, i) => {
      const c = model.parentLink.get(id);
      if (c) order.set(c.id, i);
    });
  const out = withLists(
    dataModel,
    (ptLst) => mapElements(ptLst, (p) => (removed.has(attr(p, 'modelId') ?? '') ? null : p)),
    (cxnLst) =>
      renumber(
        mapElements(cxnLst, (c) => {
          const id = attr(c, 'modelId') ?? '';
          const touches =
            removed.has(attr(c, 'srcId') ?? '') || removed.has(attr(c, 'destId') ?? '');
          return cxns.has(id) || touches ? null : c;
        }),
        order
      )
  );
  return { dataModel: out, removed };
}
