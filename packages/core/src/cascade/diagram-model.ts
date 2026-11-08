// The data model of a diagram (ECMA-376 Part 1, 21.4.3): its points (`dgm:pt`) and the
// connections between them (`dgm:cxn`). The parent-of connections (`parOf`, the default type)
// make the node tree under the document point. Presentation connections (`presOf`) tell which
// point's text a presentation shape shows; a saved drawing shape (`dsp:sp/@modelId`) names its
// presentation point, so `presOf` links a drawing shape to the node whose text it shows.
//
// Every value here comes from the file: ids are compared as strings and never parsed, and an
// order (`srcOrd`) that does not read as an integer sorts last.

import { attr, child, children, type Element } from '../store/index.ts';

/** One point of the data model. `type` is `node` when the file leaves it out. */
export interface DiagramPoint {
  readonly id: string;
  readonly type: string;
  readonly el: Element;
}

/** One connection. `type` is `parOf` when the file leaves it out. */
export interface DiagramConnection {
  readonly id: string;
  readonly type: string;
  readonly src: string;
  readonly dest: string;
  readonly srcOrd: number;
  readonly parTrans?: string;
  readonly sibTrans?: string;
  readonly el: Element;
}

/** The read data model: points by id, connections, and the node tree under the document. */
export interface DiagramModel {
  readonly points: ReadonlyMap<string, DiagramPoint>;
  readonly connections: readonly DiagramConnection[];
  /** The document point (`type="doc"`), the root of the node tree. */
  readonly root: string | undefined;
  /** The parent-of connection that leads to each point. */
  readonly parentLink: ReadonlyMap<string, DiagramConnection>;
  /** Each point's children, in `srcOrd` order. */
  readonly childrenOf: (id: string) => readonly string[];
}

/** Whether a point holds content: a node or an assistant node (21.4.7.4 ST_PtType). */
export function isContentPoint(point: DiagramPoint | undefined): boolean {
  return point?.type === 'node' || point?.type === 'asst';
}

function order(value: string | undefined): number {
  return value !== undefined && /^\d{1,9}$/.test(value) ? Number(value) : Number.MAX_SAFE_INTEGER;
}

/** Read the points and connections of a data model root (`dgm:dataModel`). */
export function readDiagramModel(dataModel: Element): DiagramModel {
  const points = new Map<string, DiagramPoint>();
  let root: string | undefined;
  for (const pt of children(child(dataModel, 'dgm:ptLst'), 'dgm:pt')) {
    const id = attr(pt, 'modelId');
    if (!id || points.has(id)) continue;
    const type = attr(pt, 'type') ?? 'node';
    points.set(id, { id, type, el: pt });
    if (type === 'doc' && root === undefined) root = id;
  }
  const connections: DiagramConnection[] = [];
  for (const cxn of children(child(dataModel, 'dgm:cxnLst'), 'dgm:cxn')) {
    const id = attr(cxn, 'modelId');
    const src = attr(cxn, 'srcId');
    const dest = attr(cxn, 'destId');
    if (!id || !src || !dest) continue;
    const parTrans = attr(cxn, 'parTransId');
    const sibTrans = attr(cxn, 'sibTransId');
    connections.push({
      id,
      type: attr(cxn, 'type') ?? 'parOf',
      src,
      dest,
      srcOrd: order(attr(cxn, 'srcOrd')),
      ...(parTrans ? { parTrans } : {}),
      ...(sibTrans ? { sibTrans } : {}),
      el: cxn,
    });
  }
  const parentLink = new Map<string, DiagramConnection>();
  const kids = new Map<string, DiagramConnection[]>();
  for (const c of connections) {
    if (c.type !== 'parOf' || parentLink.has(c.dest)) continue;
    parentLink.set(c.dest, c);
    const list = kids.get(c.src) ?? [];
    list.push(c);
    kids.set(c.src, list);
  }
  const sorted = new Map<string, readonly string[]>();
  for (const [src, list] of kids)
    sorted.set(
      src,
      [...list].sort((a, b) => a.srcOrd - b.srcOrd).map((c) => c.dest)
    );
  return {
    points,
    connections,
    root,
    parentLink,
    childrenOf: (id) => sorted.get(id) ?? [],
  };
}

/**
 * The node whose text each drawing shape shows, by the shape's `modelId`: the one content point
 * that a `presOf` connection names for the shape's presentation point, or the content point
 * whose id the shape carries itself. A shape that shows the text of more than one point (a
 * shape with its children as bullets) is left out: its text is not one point's text.
 */
export function diagramTextNodes(model: DiagramModel): ReadonlyMap<string, string> {
  const sources = new Map<string, Set<string>>();
  for (const c of model.connections) {
    if (c.type !== 'presOf') continue;
    const set = sources.get(c.dest) ?? new Set<string>();
    set.add(c.src);
    sources.set(c.dest, set);
  }
  const out = new Map<string, string>();
  for (const [dest, set] of sources) {
    if (set.size !== 1) continue;
    const [src] = set;
    if (src && isContentPoint(model.points.get(src))) out.set(dest, src);
  }
  for (const p of model.points.values())
    if (isContentPoint(p) && !sources.has(p.id)) out.set(p.id, p.id);
  return out;
}

const textNodeCache = new WeakMap<Element, ReadonlyMap<string, string>>();

/** `diagramTextNodes` of a data model root, read once per root. */
export function textNodesOf(dataModel: Element): ReadonlyMap<string, string> {
  let nodes = textNodeCache.get(dataModel);
  if (!nodes) textNodeCache.set(dataModel, (nodes = diagramTextNodes(readDiagramModel(dataModel))));
  return nodes;
}
