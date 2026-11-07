// The writes of connectors glued to shapes (see `connectors.ts` for the routes).
//
// A transform with a `connector` route writes its preset and adjust values
// (`a:prstGeom/a:avLst`) and its glued ends (`p:cNvCxnSpPr/a:stCxn` and `a:endCxn`,
// ECMA-376 Part 1, 19.3.1.13 and 20.1.2.2.4). A connector that moves with no route keeps each
// end whose shape moves with it and lets the other ends go: it no longer sits on their sites.
// The connectors glued to the moved shapes follow them in the same transaction, so a move and
// the connectors it reroutes are one undo step and one set of ops for collaboration.

import type { ConnectorRoute, TransformItem } from '../../contracts/neutral.ts';
import { presetAdjustments, type LayoutItem } from '../../layout/index.ts';
import { attr, child, elements, NS, path, type Element } from '../../store/index.ts';
import { CONNECTOR_PRESET, byDrawingId, followingConnectors } from './connectors.ts';
import { setTransforms } from './group-commands.ts';
import { CommandRefusal, type OpBatch } from './op-batch.ts';
import { MAX_UNSIGNED_INT } from './package-edits.ts';
import { specFromXml } from './shape-commands.ts';

const GUIDE_LIMIT = 2_147_483_647;
/** The shapes a connector end glues to. */
const GLUE_TARGETS = new Set(['sp', 'pic', 'graphicFrame']);

/** The `p:cNvPr/@id` of a shape element. */
function drawingIdOf(el: Element): string | undefined {
  const nv = elements(el).find((c) => c.ns === NS.p && c.local.startsWith('nv'));
  return attr(child(nv, 'p:cNvPr'), 'id');
}

/** The connector's `p:cNvCxnSpPr`, made when the file has none. */
function connectionProps(batch: OpBatch, part: string, cxn: Element): Element {
  const nv = child(cxn, 'p:nvCxnSpPr');
  if (!nv) throw new CommandRefusal('errors.malformed');
  const props = child(nv, 'p:cNvCxnSpPr');
  if (props) return props;
  // CT_ConnectorNonVisual: `p:cNvPr`, then `p:cNvCxnSpPr`.
  const id = batch.insert({
    op: 'insertNode',
    part,
    parent: nv.id,
    index: 1,
    spec: specFromXml('<p:cNvCxnSpPr/>'),
  });
  return batch.element(part, id);
}

/** Glue one end of a connector to a site, or let it go (`target` null). */
function writeEnd(
  batch: OpBatch,
  part: string,
  cxn: Element,
  name: 'stCxn' | 'endCxn',
  target: { readonly nodeId: string; readonly index: number } | null
): void {
  let glue: { id: string; idx: number } | null = null;
  if (target) {
    let el: Element;
    try {
      el = batch.element(part, target.nodeId);
    } catch {
      throw new CommandRefusal('errors.unsupported', 'connection');
    }
    const id = drawingIdOf(el);
    if (
      el.id === cxn.id ||
      !GLUE_TARGETS.has(el.kind) ||
      id === undefined ||
      !Number.isInteger(target.index) ||
      target.index < 0 ||
      target.index > MAX_UNSIGNED_INT
    )
      throw new CommandRefusal('errors.unsupported', 'connection');
    glue = { id, idx: target.index };
  }
  const props = connectionProps(batch, part, batch.element(part, cxn.id));
  const old = elements(props).find((c) => c.ns === NS.a && c.local === name);
  if (!glue) {
    if (old) batch.push({ op: 'removeNode', part, node: old.id });
    return;
  }
  if (old) {
    for (const [local, value] of [
      ['id', glue.id],
      ['idx', String(glue.idx)],
    ] as const)
      if (attr(old, local) !== value)
        batch.push({ op: 'setAttribute', part, node: old.id, name: { ns: '', local }, value });
    return;
  }
  // CT_NonVisualConnectorProperties: `a:cxnSpLocks`, `a:stCxn`, `a:endCxn`, `a:extLst`.
  const before = name === 'stCxn' ? ['endCxn', 'extLst'] : ['extLst'];
  const index = props.children.findIndex(
    (c) => c.type === 'element' && c.ns === NS.a && before.includes(c.local)
  );
  batch.insert({
    op: 'insertNode',
    part,
    parent: props.id,
    index: index < 0 ? props.children.length : index,
    spec: specFromXml(`<a:${name} id="${glue.id}" idx="${glue.idx}"/>`),
  });
}

/** Write a connector's preset and adjust values, when they change. */
function writeGeometry(batch: OpBatch, part: string, cxn: Element, route: ConnectorRoute): void {
  if (route.preset === undefined && route.adjust === undefined) return;
  const spPr = child(cxn, 'p:spPr');
  const geom = child(spPr, 'a:prstGeom');
  if (!spPr || !geom) throw new CommandRefusal('errors.unsupported', 'connector geometry');
  const preset = route.preset ?? attr(geom, 'prst') ?? '';
  if (!CONNECTOR_PRESET.test(preset))
    throw new CommandRefusal('errors.unsupported', 'connector geometry');
  const names = (presetAdjustments(preset, 0, 0, new Map()) ?? []).map((a) => a.name);
  const values = route.adjust ?? {};
  for (const [name, v] of Object.entries(values))
    if (!names.includes(name) || !Number.isInteger(v) || Math.abs(v) > GUIDE_LIMIT)
      throw new CommandRefusal('errors.unsupported', 'adjust value');
  const gds = names
    .filter((n) => Object.hasOwn(values, n))
    .map((n) => `<a:gd name="${n}" fmla="val ${values[n]}"/>`)
    .join('');
  const now = elements(child(geom, 'a:avLst'))
    .map((g) => `<a:gd name="${attr(g, 'name')}" fmla="${attr(g, 'fmla')}"/>`)
    .join('');
  if (attr(geom, 'prst') === preset && now === gds) return;
  const index = spPr.children.indexOf(geom);
  batch.push({ op: 'removeNode', part, node: geom.id });
  batch.insert({
    op: 'insertNode',
    part,
    parent: spPr.id,
    index,
    spec: specFromXml(`<a:prstGeom prst="${preset}"><a:avLst>${gds}</a:avLst></a:prstGeom>`),
  });
}

/** Write a connector's route: its preset, its adjust values, and its glued ends. */
export function writeConnectorRoute(
  batch: OpBatch,
  part: string,
  nodeId: string,
  route: ConnectorRoute
): void {
  const cxn = batch.element(part, nodeId);
  if (cxn.kind !== 'cxnSp') throw new CommandRefusal('disabled.unsupportedElement');
  writeGeometry(batch, part, cxn, route);
  if (route.start !== undefined) writeEnd(batch, part, cxn, 'stCxn', route.start);
  if (route.end !== undefined) writeEnd(batch, part, cxn, 'endCxn', route.end);
}

/**
 * After the transforms `items` are written: write their connector routes, let go the ends of
 * moved connectors whose shapes stay, and route the connectors glued to the moved shapes.
 * `layout` is the slide's layout before the transforms.
 */
export function writeConnectors(
  batch: OpBatch,
  part: string,
  items: readonly TransformItem[],
  layout: readonly LayoutItem[]
): void {
  const moved = new Set(items.map((t) => t.nodeId));
  let targets: Map<string, LayoutItem> | undefined;
  for (const t of items) {
    if (t.connector) {
      writeConnectorRoute(batch, part, t.nodeId, t.connector);
      continue;
    }
    const el = batch.element(part, t.nodeId);
    if (el.kind !== 'cxnSp') continue;
    const props = path(el, 'p:nvCxnSpPr', 'p:cNvCxnSpPr');
    for (const end of elements(props)) {
      if (end.ns !== NS.a || (end.local !== 'stCxn' && end.local !== 'endCxn')) continue;
      targets ??= byDrawingId(layout);
      const target = targets.get(attr(end, 'id') ?? '');
      if (target && !moved.has(target.element.nodeId))
        batch.push({ op: 'removeNode', part, node: end.id });
    }
  }
  const follow = followingConnectors(layout, items);
  if (follow.length === 0) return;
  setTransforms(batch, part, follow, layout);
  for (const t of follow) if (t.connector) writeConnectorRoute(batch, part, t.nodeId, t.connector);
}

/**
 * Write transforms (`setTransforms`), then the connectors they move or reroute
 * (`writeConnectors`).
 */
export function placeShapes(
  batch: OpBatch,
  part: string,
  items: readonly TransformItem[],
  layout: readonly LayoutItem[]
): void {
  setTransforms(batch, part, items, layout);
  writeConnectors(batch, part, items, layout);
}
