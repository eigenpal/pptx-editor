// The saved drawing of a diagram after a structure edit (`diagram-structure.ts`). The editor
// draws a diagram from its saved drawing (`dsp:drawing`), so after a node is added or deleted
// the drawing is laid out again here, from the drawing itself: a simple linear layout, not the
// layout definition's algorithms.
//
// The drawing must read as a line of blocks. Each top-level node is one block: one shape for the
// node, and one shape for each of its children, stacked across the line. A shape between two
// blocks may show the sibling transition of the first one (an arrow of a process). The blocks
// share the line's length evenly, with the gap ratio they had; children keep their size and
// spacing, and shrink only when they would leave the frame. A drawing that does not read this
// way (a cycle, a hierarchy, a grid, nested groups, a third level) refuses the edit: its
// layout is not supported, and nothing is written.

import { attr, child, children, elements, intAttr, type Element } from '../../store/index.ts';
import {
  diagramTextNodes,
  isContentPoint,
  type DiagramModel,
} from '../../cascade/diagram-model.ts';
import { MAX_DIAGRAM_SHAPES } from '../../cascade/diagram.ts';
import { CommandRefusal } from './op-batch.ts';
import { dml, mapElements, updatePath, withAttr, withChildren } from './diagram-xml.ts';

interface Box {
  readonly x: number;
  readonly y: number;
  readonly cx: number;
  readonly cy: number;
}

type Axis = 'x' | 'y';

/** The largest coordinate the schema allows (`ST_Coordinate`). */
const MAX_COORDINATE = 27273042316900;

const unsupported = (): never => {
  throw new CommandRefusal('errors.diagram.layoutUnsupported');
};

const startOf = (b: Box, ax: Axis): number => (ax === 'x' ? b.x : b.y);
const sizeOf = (b: Box, ax: Axis): number => (ax === 'x' ? b.cx : b.cy);
const endOf = (b: Box, ax: Axis): number => startOf(b, ax) + sizeOf(b, ax);
const centerOf = (b: Box, ax: Axis): number => startOf(b, ax) + sizeOf(b, ax) / 2;
const other = (ax: Axis): Axis => (ax === 'x' ? 'y' : 'x');

/** `b` with its start and size on `ax` replaced. */
function along(b: Box, ax: Axis, start: number, size: number): Box {
  return ax === 'x' ? { ...b, x: start, cx: size } : { ...b, y: start, cy: size };
}

function boxOf(xfrm: Element | undefined): Box | undefined {
  const off = child(xfrm, 'a:off');
  const ext = child(xfrm, 'a:ext');
  const x = intAttr(off, 'x');
  const y = intAttr(off, 'y');
  const cx = intAttr(ext, 'cx');
  const cy = intAttr(ext, 'cy');
  if (x === undefined || y === undefined || cx === undefined || cy === undefined) return undefined;
  if (cx < 0 || cy < 0) return undefined;
  return { x, y, cx, cy };
}

/** One drawing shape and its box. */
interface Drawn {
  readonly el: Element;
  readonly box: Box;
}

/** `xfrm` (an `a:xfrm` or `dsp:txXfrm`) with its box set to `box`. */
function writeBox(xfrm: Element, box: Box): Element {
  const r = (v: number) => {
    if (!(Math.abs(v) <= MAX_COORDINATE)) return unsupported();
    return String(Math.round(v));
  };
  if (box.cx < 0 || box.cy < 0) return unsupported();
  return mapElements(xfrm, (e) =>
    e.local === 'off'
      ? withAttr(withAttr(e, 'x', r(box.x)), 'y', r(box.y))
      : e.local === 'ext'
        ? withAttr(withAttr(e, 'cx', r(box.cx)), 'cy', r(box.cy))
        : e
  );
}

/** A shape moved from `from` to `to`: its box, and its text box in proportion. */
function placed(shape: Element, from: Box, to: Box): Element {
  const sx = from.cx ? to.cx / from.cx : 1;
  const sy = from.cy ? to.cy / from.cy : 1;
  let out = updatePath(shape, ['dsp:spPr', 'a:xfrm'], (x) => writeBox(x, to));
  const tx = boxOf(child(shape, 'dsp:txXfrm'));
  if (tx)
    out = updatePath(out, ['dsp:txXfrm'], (x) =>
      writeBox(x, {
        x: to.x + (tx.x - from.x) * sx,
        y: to.y + (tx.y - from.y) * sy,
        cx: tx.cx * sx,
        cy: tx.cy * sy,
      })
    );
  return out;
}

/** A copy of `shape` for a new node: its own model id, and one empty paragraph of text. */
function cloned(shape: Element, modelId: string): Element {
  const out = withAttr(shape, 'modelId', modelId);
  return updatePath(out, ['dsp:txBody'], (body) => {
    const first = children(body, 'a:p')[0];
    const pPr = child(first, 'a:pPr');
    const run = elements(first).find((e) => e.local === 'r');
    const props = child(run, 'a:rPr') ?? child(first, 'a:endParaRPr');
    const end = props ? { ...props, local: 'endParaRPr' } : dml('endParaRPr');
    const kids = body.children.filter((c) => c.type !== 'element' || c.local !== 'p');
    return withChildren(body, [...kids, dml('p', {}, pPr ? [pPr, end] : [end])]);
  });
}

/** What a drawing shape shows: a node's text, or a top-level node's sibling transition. */
function ownerOf(model: DiagramModel, modelId: string): string | undefined {
  const p = model.points.get(modelId);
  if (p?.type === 'pres') return attr(child(p.el, 'dgm:prSet'), 'presAssocID');
  if (p) return modelId;
  return model.connections.find((c) => c.type === 'presOf' && c.dest === modelId)?.src;
}

/** How deep a node is under the document point (1 for a top-level node), or 0. */
function depthOf(model: DiagramModel, id: string): number {
  let at = id;
  for (let depth = 1; depth <= 64; depth += 1) {
    const link = model.parentLink.get(at);
    if (!link) return 0;
    if (link.src === model.root) return depth;
    at = link.src;
  }
  return 0;
}

/** The drawing read as a line of blocks. */
interface Line {
  readonly heads: ReadonlyMap<string, Drawn>;
  readonly kids: ReadonlyMap<string, Drawn>;
  /** Connector shapes, by the top-level node whose sibling transition they show. */
  readonly connectors: ReadonlyMap<string, Drawn>;
}

function readLine(shapes: readonly Element[], model: DiagramModel): Line {
  const text = diagramTextNodes(model);
  const heads = new Map<string, Drawn>();
  const kids = new Map<string, Drawn>();
  const connectors = new Map<string, Drawn>();
  const top = model.root ? model.childrenOf(model.root) : [];
  const bySibTrans = new Map<string, string>();
  for (const t of top) {
    const s = model.parentLink.get(t)?.sibTrans;
    if (s) bySibTrans.set(s, t);
  }
  for (const el of shapes) {
    const box = boxOf(child(child(el, 'dsp:spPr'), 'a:xfrm'));
    const id = attr(el, 'modelId') ?? '';
    if (!box) return unsupported();
    const node = text.get(id);
    if (node) {
      const depth = depthOf(model, node);
      const into = depth === 1 ? heads : depth === 2 ? kids : undefined;
      if (!into || into.has(node)) return unsupported();
      into.set(node, { el, box });
      continue;
    }
    const owner = bySibTrans.get(ownerOf(model, id) ?? '');
    if (!owner || connectors.has(owner)) return unsupported();
    connectors.set(owner, { el, box });
  }
  for (const t of top) {
    if (!heads.has(t)) return unsupported();
    for (const k of model.childrenOf(t)) {
      if (!kids.has(k)) return unsupported();
      if (model.childrenOf(k).length > 0) return unsupported();
    }
  }
  if (top.length > 0 && connectors.size >= top.length) return unsupported();
  return { heads, kids, connectors };
}

/** What changed: the node a new block or child is copied from, by new node id. */
export interface RelayoutInput {
  readonly drawing: Element;
  readonly before: DiagramModel;
  readonly after: DiagramModel;
  /** The frame's extent, the drawing's space. */
  readonly extent: { readonly cx: number; readonly cy: number };
  /** A new node and the node next to which it was added. */
  readonly inserted?: { readonly node: string; readonly target: string };
}

/** Lay the saved drawing out again for the data model `after`. */
export function relayoutDrawing(input: RelayoutInput): Element {
  const { drawing, before, after, extent } = input;
  if (!(
    extent.cx > 0 &&
    extent.cy > 0 &&
    extent.cx <= MAX_COORDINATE &&
    extent.cy <= MAX_COORDINATE
  ))
    return unsupported();
  const tree = child(drawing, 'dsp:spTree');
  if (!tree || child(child(tree, 'dsp:grpSpPr'), 'a:xfrm')) return unsupported();
  const els = elements(tree);
  if (els.some((e) => e.local === 'grpSp')) return unsupported();
  const shapes = children(tree, 'dsp:sp');
  const line = readLine(shapes, before);
  const topBefore = before.root ? before.childrenOf(before.root) : [];
  const topAfter = (after.root ? after.childrenOf(after.root) : []).filter((t) =>
    isContentPoint(after.points.get(t))
  );
  if (topBefore.length === 0 || topAfter.length === 0) return unsupported();
  // The line draws two levels: a node below a child has no place in it.
  for (const t of topAfter)
    for (const k of after.childrenOf(t))
      if (after.childrenOf(k).length > 0) throw new CommandRefusal('errors.diagram.noChildren');
  const head = (t: string) => line.heads.get(t)!;
  const kidsOf = (t: string) => before.childrenOf(t).map((k) => line.kids.get(k)!);

  // The line's axis and direction, from the top-level shapes.
  const centers = topBefore.map((t) => head(t).box);
  let outer: Axis;
  let dir = 1;
  if (centers.length >= 2) {
    const a = centers[0]!;
    const b = centers[centers.length - 1]!;
    const dx = centerOf(b, 'x') - centerOf(a, 'x');
    const dy = centerOf(b, 'y') - centerOf(a, 'y');
    outer = Math.abs(dx) >= Math.abs(dy) ? 'x' : 'y';
    dir = (outer === 'x' ? dx : dy) >= 0 ? 1 : -1;
    const cross = other(outer);
    const tolerance = Math.max(1, 0.02 * (cross === 'x' ? extent.cx : extent.cy));
    for (let i = 1; i < centers.length; i += 1) {
      const step = (centerOf(centers[i]!, outer) - centerOf(centers[i - 1]!, outer)) * dir;
      if (step <= 0) return unsupported();
      if (Math.abs(centerOf(centers[i]!, cross) - centerOf(a, cross)) > tolerance)
        return unsupported();
    }
  } else {
    const k = kidsOf(topBefore[0]!)[0];
    const h = centers[0]!;
    if (k)
      outer =
        Math.abs(centerOf(k.box, 'x') - centerOf(h, 'x')) >=
        Math.abs(centerOf(k.box, 'y') - centerOf(h, 'y'))
          ? 'y'
          : 'x';
    else outer = extent.cx >= extent.cy ? 'x' : 'y';
  }
  const inner = other(outer);

  // Each block's span on the line.
  const blockOf = (t: string) => {
    const boxes = [head(t).box, ...kidsOf(t).map((k) => k.box)];
    let s = Infinity;
    let e = -Infinity;
    for (const b of boxes) {
      s = Math.min(s, startOf(b, outer));
      e = Math.max(e, endOf(b, outer));
    }
    return { s, e };
  };
  const blocks = new Map(topBefore.map((t) => [t, blockOf(t)]));
  const visual = [...blocks.values()].sort((p, q) => p.s - q.s);
  const size = visual.reduce((m, b) => Math.max(m, b.e - b.s), 0);
  if (!(size > 0)) return unsupported();
  const spanStart = visual[0]!.s;
  const spanEnd = visual.reduce((m, b) => Math.max(m, b.e), -Infinity);
  let gap = 0.1 * size;
  if (visual.length >= 2) {
    let sum = 0;
    for (let i = 1; i < visual.length; i += 1) sum += Math.max(0, visual[i]!.s - visual[i - 1]!.e);
    gap = sum / (visual.length - 1);
  }
  const n = topAfter.length;
  const ratio = gap / size;
  const size2 = (spanEnd - spanStart) / (n + (n - 1) * ratio);
  const gap2 = size2 * ratio;
  const f = size2 / size;
  const startAt = (j: number) =>
    dir > 0 ? spanStart + j * (size2 + gap2) : spanEnd - size2 - j * (size2 + gap2);

  // The child template: the first child of any block, and its place under its block's shape.
  const anyParent = topBefore.find((t) => before.childrenOf(t).length > 0);
  const childTemplate = (t: string | undefined) => {
    const from = t && before.childrenOf(t).length > 0 ? t : anyParent;
    if (!from) return undefined;
    const ks = kidsOf(from);
    const h = head(from).box;
    const k0 = ks[0]!.box;
    const k1 = ks[1]?.box;
    const step = k1
      ? startOf(k1, inner) - startOf(k0, inner)
      : sizeOf(k0, inner) + Math.max(0, startOf(k0, inner) - endOf(h, inner));
    return {
      from,
      drawn: ks[0]!,
      offset: startOf(k0, inner) - startOf(h, inner),
      step: step > 0 ? step : sizeOf(k0, inner) * 1.1,
    };
  };

  const inserted = input.inserted;
  const out: Element[] = [];
  const mapOuter = (b: Box, oldStart: number, newStart: number): Box =>
    along(b, outer, newStart + (startOf(b, outer) - oldStart) * f, sizeOf(b, outer) * f);
  const bound = inner === 'x' ? extent.cx : extent.cy;
  topAfter.forEach((t, j) => {
    const newStart = startAt(j);
    const fresh = !line.heads.has(t);
    if (fresh && inserted?.node !== t) return unsupported();
    const templateTop = fresh ? inserted!.target : t;
    if (!line.heads.has(templateTop)) return unsupported();
    const block = blocks.get(templateTop)!;
    const h = head(templateTop);
    const headBox = mapOuter(h.box, block.s, newStart);
    out.push(placed(fresh ? cloned(h.el, t) : h.el, h.box, headBox));
    const kidsAfter = after.childrenOf(t);
    const kidsBefore = fresh ? [] : before.childrenOf(t);
    const same =
      kidsAfter.length === kidsBefore.length && kidsAfter.every((k, i) => k === kidsBefore[i]);
    if (same) {
      for (const k of kidsAfter) {
        const d = line.kids.get(k)!;
        out.push(placed(d.el, d.box, mapOuter(d.box, block.s, newStart)));
      }
      return;
    }
    if (kidsAfter.length === 0) return;
    // The block's children changed: stack them again from the template.
    const tpl = childTemplate(fresh ? undefined : t);
    if (!tpl) throw new CommandRefusal('errors.diagram.noChildren');
    const tplBlock = blocks.get(tpl.from)!;
    let childSize = sizeOf(tpl.drawn.box, inner);
    let step = tpl.step;
    const first = startOf(h.box, inner) + tpl.offset;
    const need = (kidsAfter.length - 1) * step + childSize;
    if (first + need > bound && bound > first) {
      const fit = (bound - first) / need;
      childSize *= fit;
      step *= fit;
    }
    kidsAfter.forEach((k, i) => {
      const own = line.kids.get(k);
      const src = own ?? tpl.drawn;
      const srcBlockStart = own ? block.s : tplBlock.s;
      const box = along(
        mapOuter(src.box, srcBlockStart, newStart),
        inner,
        first + i * step,
        childSize
      );
      out.push(placed(own ? own.el : cloned(src.el, k), src.box, box));
    });
  });

  // Connectors between neighbouring blocks, as the template connector shows them.
  const tplConnector = topBefore.map((t) => line.connectors.get(t)).find(Boolean);
  if (tplConnector) {
    const s = Math.min(1, f);
    const cBox = tplConnector.box;
    for (let j = 0; j + 1 < n; j += 1) {
      const t = topAfter[j]!;
      const sib = after.parentLink.get(t)?.sibTrans;
      const own = line.connectors.get(t);
      if (!own && !sib) continue;
      const a = startAt(j);
      const b = startAt(j + 1);
      const mid = (Math.min(a, b) + size2 + Math.max(a, b)) / 2;
      const cx = cBox.cx * s;
      const cy = cBox.cy * s;
      const crossCenter = centerOf(cBox, inner);
      const box: Box =
        outer === 'x'
          ? { x: mid - cx / 2, y: crossCenter - cy / 2, cx, cy }
          : { x: crossCenter - cx / 2, y: mid - cy / 2, cx, cy };
      const src = own ?? tplConnector;
      out.push(placed(own ? own.el : cloned(src.el, sib!), src.box, box));
    }
  }

  if (out.length > MAX_DIAGRAM_SHAPES) throw new CommandRefusal('errors.diagram.tooLarge');
  const firstShape = tree.children.findIndex((c) => c.type === 'element' && c.local === 'sp');
  const lead = tree.children.filter(
    (c, i) => (firstShape < 0 || i < firstShape) && (c.type !== 'element' || c.local !== 'sp')
  );
  const tail = tree.children.filter(
    (c, i) => firstShape >= 0 && i > firstShape && (c.type !== 'element' || c.local !== 'sp')
  );
  const nextTree = withChildren(tree, [...lead, ...out, ...tail]);
  return mapElements(drawing, (e) => (e === tree ? nextTree : e));
}
