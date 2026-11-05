// Re-fit after text edits (ECMA-376 Part 1, 21.1.2.1.1 bodyPr):
//
// - `a:spAutoFit`: the shape grows or shrinks to the text, across its lines: the height for
//   horizontal text (exact: same layout), the width for rotated text, and with columns the
//   smallest height at which every column fits, from one layout of the lines and the column
//   flow rule of `text-layout.ts`. The size comes from the shape's geometry, so an ellipse
//   grows until its text rectangle holds the text, and a rounded rectangle is laid out again
//   at its new width (three times at most). Rotated text, columns, and shapes that are not
//   rectangles are APPROXIMATIONS, and a rotated box that does not wrap keeps its size; all
//   are listed as approximations on the public fidelity page. A size stays within the bound the cascade reads back
//   (`XFRM_BOUND`).
// - `a:normAutofit`: shrink text on overflow. The spec does not say how to pick `fontScale` and
//   `lnSpcReduction`. The search starts from 100 % on every edit, not from the shape's last
//   state: a deck saved at `fontScale="62500"` whose text is cut back fits again at 100 %
//   (measured). The steps differ by placeholder type (measured on a 90 pt high title box with
//   text from 20 pt to 72 pt, saved after each edit):
//   - A title (`title`, `ctrTitle`) takes one step at most: `fontScale="90000"` with no line
//     spacing reduction. Text that still overflows at 90 % wraps and overflows the box.
//   - Other text (a body placeholder, a text box): 10 % line-spacing reduction, then a 92.5 %
//     font scale with it, then the 20 % reduction, then font scale in 7.5 % steps with the
//     20 % reduction, down to 25 %. There is no 20 % step at full size. Measured on a text box
//     of three 72 pt lines at 21 heights from 80 pt to 280 pt: every step from 100 % to 55 %
//     appeared in this order. The steps below 55 % follow the same 7.5 % rule and are not
//     measured. This is an APPROXIMATION, listed on the public fidelity page.

import { attr, child, NS, type Element } from '../../store/index.ts';
import {
  resolveSlide,
  XFRM_BOUND,
  type ResolvedElement,
  type ResolvedShape,
} from '../../cascade/index.ts';
import { evaluateGeometry, layoutText, type TextMeasurer } from '../../layout/index.ts';
import type { OpBatch } from './op-batch.ts';
import { setTransform, specFromXml, transformHost } from './shape-commands.ts';
import { MAX_COORDINATE } from './package-edits.ts';
import { refitTable } from './table-fit.ts';

function findShape(
  elements: readonly ResolvedElement[],
  nodeId: string
): ResolvedShape | undefined {
  for (const el of elements) {
    if (el.nodeId === nodeId && el.kind === 'shape') return el;
    if (el.kind === 'group') {
      const hit = findShape(el.children, nodeId);
      if (hit) return hit;
    }
  }
  return undefined;
}

interface AutofitStep {
  readonly scale: number;
  readonly reduction: number;
}

/** The `normAutofit` steps, largest text first, in the measured order (see the header). */
const AUTOFIT_STEPS: readonly AutofitStep[] = [
  { scale: 1, reduction: 0 },
  { scale: 1, reduction: 0.1 },
  { scale: 0.925, reduction: 0.1 },
  ...[0.925, 0.85, 0.775, 0.7, 0.625, 0.55, 0.475, 0.4, 0.325, 0.25].map((scale) => ({
    scale,
    reduction: 0.2,
  })),
];

/** A title's steps: 90 % at most, with no line spacing reduction (measured; see the header). */
const TITLE_AUTOFIT_STEPS: readonly AutofitStep[] = [
  { scale: 1, reduction: 0 },
  { scale: 0.9, reduction: 0 },
];

/** The steps that a shape's `normAutofit` searches: a title's own, or the general ones. */
export function autofitStepsFor(placeholder: ResolvedShape['placeholder']): readonly AutofitStep[] {
  const type = placeholder?.type;
  return type === 'title' || type === 'ctrTitle' ? TITLE_AUTOFIT_STEPS : AUTOFIT_STEPS;
}

/** Re-fit one shape on slide `index` after its text changed. */
export function refit(
  batch: OpBatch,
  index: number,
  part: string,
  nodeId: string,
  measurer: TextMeasurer
): void {
  const slide = resolveSlide(batch.pkg, index);
  const shape = slide ? findShape(slide.elements, nodeId) : undefined;
  // A table frame: its size follows its rows, which grow to fit cell text.
  if (!shape) return void refitTable(batch, index, part, nodeId, measurer);
  if (!shape.text || shape.layer !== 'slide') return;
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
  const geometry = evaluateGeometry(shape.geometry, xfrm.cx, xfrm.cy);
  const r = geometry.textRect;
  const rect = { x: r.l, y: r.t, width: r.r - r.l, height: r.b - r.t };
  const autofit = shape.text.body.autofit;
  if (autofit.kind === 'shape') {
    let across = acrossLines(shape.text, rect, measurer);
    const { swapped } = across;
    // A rotated box that does not wrap keeps its size: its lines run along the box, and
    // sizing only across them would shrink it to one line (an approximation).
    if (swapped && shape.text.body.wrap === 'none') return;
    const old = swapped ? xfrm.cx : xfrm.cy;
    // The text rectangle at a shape size: the other side of the shape stays as it is.
    const rectAt = (size: number): Rect => {
      const t = evaluateGeometry(
        shape.geometry,
        swapped ? size : xfrm.cx,
        swapped ? xfrm.cy : size
      ).textRect;
      return { x: t.l, y: t.t, width: t.r - t.l, height: t.b - t.t };
    };
    const extentOf = (r: Rect): number => (swapped ? r.width : r.height);
    const lineSide = (r: Rect): number => (swapped ? r.height : r.width);
    const inner = (size: number): number => extentOf(rectAt(size));
    let exact = solveSize(inner, old, extentOf(rect), across.extent);
    // The lines' own side can change with the size too (a rounded rectangle's or an octagon's
    // inset follows the shorter side): lay the text out again at the new size's rectangle, and
    // solve again, three times at most.
    let laid = rect;
    for (let k = 0; k < 3; k += 1) {
      const at = Math.round(exact);
      const next = rectAt(at);
      if (Math.abs(lineSide(next) - lineSide(laid)) < 0.5) break;
      laid = next;
      across = acrossLines(shape.text, next, measurer);
      exact = solveSize(inner, at, extentOf(next), across.extent);
    }
    // With columns, up to a whole EMU, so a column break that needs every EMU still fits; a
    // change of one EMU or less is left alone otherwise, so rounding cannot make it wobble.
    const columns = shape.text.body.numCol > 1;
    let size = Math.min(BOUND, columns ? Math.ceil(exact - 0.001) : Math.round(exact));
    // Only while a step helps: a text rectangle that does not grow keeps the size.
    for (
      let i = 0;
      columns &&
      i < 4 &&
      size < BOUND &&
      inner(size) < across.extent &&
      inner(size + 1) > inner(size);
      i += 1
    )
      size += 1;
    if (size === old || (Math.abs(size - old) <= 1 && (!columns || size < old))) return;
    if (swapped) setTransform(batch, part, nodeId, xfrm.x, xfrm.y, size, xfrm.cy, rot, flips);
    else setTransform(batch, part, nodeId, xfrm.x, xfrm.y, xfrm.cx, size, rot, flips);
    return;
  }
  if (autofit.kind !== 'normal') return;
  const chosen = chooseNormAutofit(shape.text, rect, measurer, autofitStepsFor(shape.placeholder));
  writeNormAutofit(batch, part, nodeId, chosen.scale, chosen.reduction);
}

/** The largest size a refit writes: one the cascade reads back as written. */
const BOUND = Math.min(MAX_COORDINATE, XFRM_BOUND);

/**
 * The shape size whose text rectangle extent is `target`. A rectangle's text rectangle is the
 * shape less a fixed border, so the first step is exact; an ellipse's grows with the shape at
 * its own rate (preset guide formulas), so secant steps follow, a few at most. Bounded by
 * `BOUND`, so it always ends.
 */
function solveSize(
  inner: (size: number) => number,
  size: number,
  extent: number,
  target: number
): number {
  let [s0, f0] = [size, extent];
  let s1 = Math.max(0, Math.min(BOUND, size + target - extent));
  for (let i = 0; i < 6; i += 1) {
    const f1 = inner(s1);
    if (Math.abs(f1 - target) <= 0.25 || s1 === s0) return s1;
    const slope = (f1 - f0) / (s1 - s0);
    // A text rectangle that does not grow with the shape: no other size fits better, so the
    // last good size stays (the shape's own size on the first step).
    if (!(slope > 0)) return s0;
    [s0, f0] = [s1, f1];
    s1 = Math.max(0, Math.min(BOUND, s1 + (target - f1) / slope));
  }
  return s1;
}

type Text = NonNullable<ResolvedShape['text']>;
type Rect = Parameters<typeof layoutText>[1];

/** Whether the text fits its frame: the frame is the swapped one for rotated text. */
const fits = (text: Text, rect: Rect, measurer: TextMeasurer): boolean => {
  const laid = layoutText(text, rect, measurer);
  return laid.contentHeight <= laid.box.height + 1;
};

/**
 * The extent the text needs across its lines (`spAutoFit`): the box height for horizontal
 * text, the box width for rotated text (`vert`, `vert270`, `eaVert`, `wordArtVert`). With
 * columns, the smallest extent at which every column fits.
 */
function acrossLines(
  text: Text,
  rect: Rect,
  measurer: TextMeasurer
): { swapped: boolean; extent: number } {
  const laid = layoutText(text, rect, measurer);
  const swapped = Math.round(laid.rotation - text.body.rot) % 180 !== 0;
  const { numCol, tIns, bIns } = text.body;
  if (numCol <= 1) return { swapped, extent: laid.contentHeight };
  // The line breaks do not depend on the extent (only on the other side of the box), so one
  // layout tall enough for every line in the first column gives the lines; the column flow
  // then runs on their tops and heights alone. Top anchoring keeps the tops as laid.
  const tall = layoutText(
    { ...text, body: { ...text.body, anchor: 't' } },
    swapped ? { ...rect, width: MAX_COORDINATE } : { ...rect, height: MAX_COORDINATE },
    measurer
  );
  const lines = tall.lines.map((l) => ({ top: l.top - tIns, height: l.height }));
  return { swapped, extent: columnFlowHeight(lines, numCol) + tIns + bIns };
}

/**
 * The smallest inner height at which every line fits in `columns` columns, by the flow rule of
 * `text-layout.ts`: a line moves to the next column when it would end below the box and is not
 * the first line of its column. `lines` are laid out in one column. A binary search over whole
 * EMU, bounded by the coordinate limit, so it always ends.
 */
export function columnFlowHeight(
  lines: readonly { readonly top: number; readonly height: number }[],
  columns: number
): number {
  const fitsIn = (h: number): boolean => {
    let col = 0;
    let offset = 0;
    for (const l of lines) {
      if (l.top - offset + l.height > h && col < columns - 1 && l.top - offset > 0) {
        col += 1;
        offset = l.top;
      }
      if (l.top - offset + l.height > h) return false;
    }
    return true;
  };
  let hi = 0;
  for (const l of lines) hi = Math.max(hi, l.top + l.height);
  hi = Math.min(MAX_COORDINATE, Math.ceil(hi));
  let lo = 0;
  if (fitsIn(lo)) return lo;
  // `fitsIn(hi)` holds: every line in the first column.
  while (hi - lo > 1) {
    const mid = lo + Math.floor((hi - lo) / 2);
    if (fitsIn(mid)) hi = mid;
    else lo = mid;
  }
  return hi;
}

/** The first step at which the text fits; the last step when none does (the text overflows). */
function chooseNormAutofit(
  text: Text,
  rect: Rect,
  measurer: TextMeasurer,
  steps: readonly AutofitStep[]
): AutofitStep {
  for (const { scale, reduction } of steps) {
    const body = {
      ...text.body,
      autofit: { kind: 'normal' as const, fontScale: scale, lineSpaceReduction: reduction },
    };
    if (fits({ ...text, body }, rect, measurer)) return { scale, reduction };
  }
  return steps[steps.length - 1] ?? AUTOFIT_STEPS[0]!;
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
  const set = (name: string, value: string | null): void =>
    batch.push({ op: 'setAttribute', part, node: norm!.id, name: { ns: '', local: name }, value });
  set('fontScale', scale < 1 ? String(Math.round(scale * 100_000)) : null);
  set('lnSpcReduction', reduction > 0 ? String(Math.round(reduction * 100_000)) : null);
}

export { NS };
