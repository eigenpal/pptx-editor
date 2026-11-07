// Re-fit after text edits and resizes (ECMA-376 Part 1, 21.1.2.1.1 bodyPr). The sizes and
// scales come from `fitShape` (layout/autofit.ts), which describes the rules; this file writes
// them.

import { attr, child, NS, type Element } from '../../store/index.ts';
import {
  resolveSlide,
  type ResolvedElement,
  type ResolvedShape,
  type ResolvedSlide,
} from '../../cascade/index.ts';
import type { TextMeasurer } from '../../layout/index.ts';
import { fitShape } from '../../layout/autofit.ts';
import type { OpBatch } from './op-batch.ts';
import { setTransform, specFromXml, transformHost } from './shape-commands.ts';
import { noteFit, refitTable } from './table-fit.ts';

export { autofitStepsFor, columnFlowHeight } from '../../layout/autofit.ts';

/** A shape and the scale its groups draw it at (child space to slide space). */
interface Found {
  readonly shape: ResolvedShape;
  readonly sx: number;
  readonly sy: number;
}

function findShape(
  elements: readonly ResolvedElement[],
  nodeId: string,
  sx = 1,
  sy = 1
): Found | undefined {
  for (const el of elements) {
    if (el.nodeId === nodeId && el.kind === 'shape') return { shape: el, sx, sy };
    if (el.kind === 'group') {
      // As the layout places a child (`placeInGroup`): its box scales with the group.
      const ext = el.childExtent;
      const gx = ext.cx ? el.xfrm.cx / ext.cx : 1;
      const gy = ext.cy ? el.xfrm.cy / ext.cy : 1;
      const hit = findShape(el.children, nodeId, sx * gx, sy * gy);
      if (hit) return hit;
    }
  }
  return undefined;
}

/** The shapes of a resolved slide by node id, built once: a re-fit of many shapes reads it. */
const shapeTables = new WeakMap<ResolvedSlide, Map<string, Found>>();

function shapeTable(slide: ResolvedSlide): Map<string, Found> {
  let table = shapeTables.get(slide);
  if (table) return table;
  const out = new Map<string, Found>();
  const walk = (elements: readonly ResolvedElement[], sx: number, sy: number): void => {
    for (const el of elements) {
      if (el.kind === 'shape' && !out.has(el.nodeId)) out.set(el.nodeId, { shape: el, sx, sy });
      if (el.kind === 'group') {
        // As `findShape` scales a group child.
        const ext = el.childExtent;
        const gx = ext.cx ? el.xfrm.cx / ext.cx : 1;
        const gy = ext.cy ? el.xfrm.cy / ext.cy : 1;
        walk(el.children, sx * gx, sy * gy);
      }
    }
  };
  walk(slide.elements, 1, 1);
  shapeTables.set(slide, (table = out));
  return table;
}

/**
 * Re-fit one shape on slide `index` after its text changed. `resolved` is the slide as resolved
 * after every text change and before any re-fit: a re-fit changes only its own shape, so many
 * shapes re-fit from one resolution (`refitAll`).
 */
export function refit(
  batch: OpBatch,
  index: number,
  part: string,
  nodeId: string,
  measurer: TextMeasurer,
  resolved?: ResolvedSlide
): void {
  const slide = resolved ?? resolveSlide(batch.pkg, index);
  // One shape: a walk to it. Many shapes share `resolved` and its table.
  const found = !slide
    ? undefined
    : resolved
      ? shapeTable(resolved).get(nodeId)
      : findShape(slide.elements, nodeId);
  // A table frame: its size follows its rows, which grow to fit cell text.
  if (!found) return void refitTable(batch, index, part, nodeId, measurer);
  const { shape, sx, sy } = found;
  if (!shape.text || shape.layer !== 'slide') return;
  noteFit(measurer, part, nodeId);
  // A shape with no transform host (a broken file with no p:spPr) keeps its size.
  const host = transformHost(batch.element(part, nodeId));
  if (!host) return;
  const { xfrm } = shape;
  // A refit never changes the rotation: the shape's own a:xfrm keeps its rot as the file has it
  // (one with no a:off or a:ext is not read by the cascade); a shape with no a:xfrm takes the
  // rotation and flips it inherits from its layout placeholder, as they are drawn.
  const own = child(host, 'a:xfrm');
  const rot = own ? (Number(attr(own, 'rot')) || 0) / 60_000 : xfrm.rot;
  const flips = own ? undefined : { h: xfrm.flipH, v: xfrm.flipV };
  // A group child is fitted at the size its groups draw it at, with its fonts unscaled, as it
  // is laid out; a new size goes back to the child space.
  const drawn = { ...xfrm, cx: xfrm.cx * sx, cy: xfrm.cy * sy };
  const fit = fitShape(sx === 1 && sy === 1 ? shape : { ...shape, xfrm: drawn }, measurer);
  if (fit?.kind === 'size') {
    const cx = sx === 1 ? fit.cx : Math.round(fit.cx / (sx || 1));
    const cy = sy === 1 ? fit.cy : Math.round(fit.cy / (sy || 1));
    if (Math.abs(cx - xfrm.cx) + Math.abs(cy - xfrm.cy) === 0) return;
    setTransform(batch, part, nodeId, xfrm.x, xfrm.y, cx, cy, rot, flips);
  } else if (fit) writeNormAutofit(batch, part, nodeId, fit.scale, fit.reduction);
}

/**
 * Re-fit each of `nodeIds` on slide `index` from one resolution of the slide. A resolution for
 * each shape made a format change on many shapes quadratic.
 */
export function refitAll(
  batch: OpBatch,
  index: number,
  part: string,
  nodeIds: Iterable<string>,
  measurer: TextMeasurer
): void {
  const ids = [...nodeIds];
  if (ids.length === 0) return;
  const resolved = ids.length > 1 ? resolveSlide(batch.pkg, index) : undefined;
  for (const id of ids) refit(batch, index, part, id, measurer, resolved);
}

function writeNormAutofit(
  batch: OpBatch,
  part: string,
  nodeId: string,
  scale: number,
  reduction: number
): void {
  const sp = batch.element(part, nodeId);
  const bodyPr = child(child(sp, 'p:txBody'), 'a:bodyPr');
  if (!bodyPr) return;
  let norm: Element | undefined = child(bodyPr, 'a:normAutofit');
  if (!norm) {
    for (const other of ['a:noAutofit', 'a:spAutoFit']) {
      const el = child(bodyPr, other);
      if (el) batch.push({ op: 'removeNode', part, node: el.id });
    }
    // CT_TextBodyProperties: `a:prstTxWarp` first, then the autofit choice.
    const warp = bodyPr.children.findIndex((c) => c.type === 'element' && c.local === 'prstTxWarp');
    const id = batch.insert({
      op: 'insertNode',
      part,
      parent: bodyPr.id,
      index: warp + 1,
      spec: specFromXml('<a:normAutofit/>'),
    });
    norm = batch.element(part, id);
  }
  // A value the file already has writes no op, so peers and the saved file see no change.
  const set = (name: string, value: string | null): void => {
    if ((attr(norm, name) ?? null) === value) return;
    batch.push({ op: 'setAttribute', part, node: norm!.id, name: { ns: '', local: name }, value });
  };
  set('fontScale', scale < 1 ? String(Math.round(scale * 100_000)) : null);
  set('lnSpcReduction', reduction > 0 ? String(Math.round(reduction * 100_000)) : null);
}

export { NS };
