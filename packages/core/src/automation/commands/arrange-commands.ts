// Arrange commands beyond order and align: distribute, flip, rotate, lock aspect ratio, and the
// selection pane's show, hide, and rename. Transforms go through the shared transform writer;
// the others set one attribute of the element's non-visual properties.

import { attr, child, elements, NS, type Element } from '../../store/index.ts';
import type { ObjectAxis } from '../../contracts/neutral.ts';
import { CommandRefusal, type OpBatch } from './op-batch.ts';
import { parentOf } from './package-edits.ts';

/** One box in its parent's space. */
interface Box {
  readonly id: string;
  readonly x: number;
  readonly y: number;
  readonly cx: number;
  readonly cy: number;
}

/**
 * Distribute boxes so the gaps between them are equal, as Distribute does with Align Selected
 * Objects: the boxes keep their order along the axis (by their leading edge),
 * the first stays, and the last ends where the selection ended. One box centers in the space
 * (`size`, the slide in the box's parent space); two have nothing to distribute.
 */
export function distributeBoxes(
  boxes: readonly Box[],
  axis: ObjectAxis,
  size: number
): { id: string; x: number; y: number }[] {
  if (axis !== 'horizontal' && axis !== 'vertical')
    throw new CommandRefusal('errors.unsupported', 'axis');
  const h = axis === 'horizontal';
  const start = (b: Box): number => (h ? b.x : b.y);
  const length = (b: Box): number => (h ? b.cx : b.cy);
  const place = (b: Box, at: number): { id: string; x: number; y: number } =>
    h ? { id: b.id, x: at, y: b.y } : { id: b.id, x: b.x, y: at };
  if (boxes.length === 1) return [place(boxes[0], (size - length(boxes[0])) / 2)];
  if (boxes.length < 3) throw new CommandRefusal('disabled.distributeNeedsThree');
  const sorted = [...boxes].sort((a, b) => start(a) - start(b));
  const from = start(sorted[0]);
  let to = -Infinity;
  for (const b of sorted) to = Math.max(to, start(b) + length(b));
  let used = 0;
  for (const b of sorted) used += length(b);
  const gap = (to - from - used) / (sorted.length - 1);
  const out = new Map<string, { id: string; x: number; y: number }>();
  let at = from;
  for (const b of sorted) {
    out.set(b.id, place(b, at));
    at += length(b) + gap;
  }
  return boxes.map((b) => out.get(b.id)!);
}

/** How an element's parent space sits on the slide: its turn in degrees, and a mirror. */
export interface ParentFrame {
  readonly turn: number;
  readonly mirrored: boolean;
}

/**
 * The parent frame of a parent matrix `[a, b, c, d, e, f]` (x' = a·x + c·y + e,
 * y' = b·x + d·y + f). A non-uniform group scale is read as its turn only.
 */
export function parentFrame(m: readonly number[]): ParentFrame {
  const [a = 1, b = 0, c = 0, d = 1] = m;
  const mirrored = a * d - b * c < 0;
  const turn = (Math.atan2(mirrored ? -b : b, mirrored ? -a : a) * 180) / Math.PI;
  return { turn, mirrored };
}

const norm = (deg: number): number => {
  const r = ((deg % 360) + 360) % 360;
  return Math.abs(r - 360) < 1e-9 ? 0 : r;
};

/**
 * The rotation and flips after a flip on the slide: the element mirrors about its center
 * along the slide's axis. At the top level a turned element turns
 * the other way (`rot` becomes `360 - rot`). In a group, the mirror is taken into the group's
 * space: in a group turned by θ, `rot` becomes `-2θ - rot` (`2θ - rot` in a mirrored group).
 */
export function flipped(
  box: { readonly rot: number; readonly flipH: boolean; readonly flipV: boolean },
  axis: ObjectAxis,
  parent: ParentFrame = { turn: 0, mirrored: false }
): { rot: number; flipH: boolean; flipV: boolean } {
  if (axis !== 'horizontal' && axis !== 'vertical')
    throw new CommandRefusal('errors.unsupported', 'axis');
  const phi = parent.mirrored ? 2 * parent.turn : -2 * parent.turn;
  return {
    rot: norm(phi - box.rot),
    flipH: axis === 'horizontal' ? !box.flipH : box.flipH,
    flipV: axis === 'vertical' ? !box.flipV : box.flipV,
  };
}

/**
 * A rotation turned by `degrees` clockwise on the slide, normalized to 0 … 360. In a mirrored
 * group, a clockwise turn on the slide is a counterclockwise turn in the group's space.
 */
export function rotated(
  rot: number,
  degrees: number,
  parent: ParentFrame = { turn: 0, mirrored: false }
): number {
  if (!Number.isFinite(degrees) || Math.abs(degrees) > 3600)
    throw new CommandRefusal('errors.unsupported', 'rot');
  return norm(rot + (parent.mirrored ? -degrees : degrees));
}

/**
 * The non-visual property names of each element kind: its `p:nv…Pr`, the `p:cNv…Pr` inside
 * it, and the locks element there (CT_NonVisualDrawingShapeProps and its siblings, where the
 * locks are the first child).
 */
const NON_VISUAL: Readonly<Record<string, readonly [string, string, string]>> = {
  sp: ['nvSpPr', 'cNvSpPr', 'spLocks'],
  pic: ['nvPicPr', 'cNvPicPr', 'picLocks'],
  grpSp: ['nvGrpSpPr', 'cNvGrpSpPr', 'grpSpLocks'],
  cxnSp: ['nvCxnSpPr', 'cNvCxnSpPr', 'cxnSpLocks'],
  graphicFrame: ['nvGraphicFramePr', 'cNvGraphicFramePr', 'graphicFrameLocks'],
};

/** An element of the shape tree: a shape, picture, group, connector, or frame. */
function treeElement(batch: OpBatch, part: string, id: string): Element {
  const el = batch.element(part, id);
  if (!Object.prototype.hasOwnProperty.call(NON_VISUAL, el.kind))
    throw new CommandRefusal('disabled.unsupportedElement');
  return el;
}

/** The element's `p:cNvPr` (name, id, hidden). */
function drawingProps(el: Element): Element {
  const cNvPr = child(child(el, `p:${NON_VISUAL[el.kind][0]}`), 'p:cNvPr');
  if (!cNvPr) throw new CommandRefusal('errors.malformed');
  return cNvPr;
}

/** The element's locks element, or undefined; and the `p:cNv…Pr` that holds it. */
export function locksOf(el: Element): { host?: Element; locks?: Element } {
  const names = NON_VISUAL[el.kind];
  if (!names) return {};
  const host = child(child(el, `p:${names[0]}`), `p:${names[1]}`);
  return host ? { host, locks: child(host, `a:${names[2]}`) } : {};
}

/** Whether an element's aspect ratio is locked (`noChangeAspect`). */
export function aspectLocked(el: Element | undefined): boolean {
  const v = el ? locksOf(el).locks : undefined;
  const a = v?.attrs.find((x) => x.ns === '' && x.local === 'noChangeAspect')?.value;
  return a === '1' || a === 'true';
}

/**
 * An element and its copies in the other branches of the `mc:AlternateContent` it sits in.
 * The editor shows the copy of `alternateBranch` (`cascade/alternate-content.ts`); another
 * consumer can show another copy (ECMA-376 Part 3, Markup Compatibility), so a hide, a rename,
 * or a lock goes to every copy (matched by kind and `p:cNvPr/@id`).
 */
function withCopies(batch: OpBatch, part: string, el: Element): Element[] {
  const out = [el];
  let branch: Element | undefined;
  for (let p = parentOf(batch, part, el.id), depth = 0; depth < 32; depth += 1) {
    if (p.ns === NS.mc && (p.local === 'Fallback' || p.local === 'Choice')) branch = p;
    if (p.kind === 'spTree' || branch) break;
    p = parentOf(batch, part, p.id);
  }
  if (!branch) return out;
  const wrapper = parentOf(batch, part, branch.id);
  const id = attr(drawingProps(el), 'id');
  const find = (node: Element, depth: number): Element | undefined => {
    for (const c of elements(node)) {
      if (c.kind === el.kind && attr(drawingProps(c), 'id') === id) return c;
      const inner = depth < 32 ? find(c, depth + 1) : undefined;
      if (inner) return inner;
    }
    return undefined;
  };
  for (const other of elements(wrapper)) {
    if (other.id === branch.id || other.ns !== NS.mc) continue;
    const copy = find(other, 0);
    if (copy) out.push(copy);
  }
  return out;
}

/** Lock or unlock the aspect ratio of elements. */
export function setLockAspectRatio(
  batch: OpBatch,
  part: string,
  ids: readonly string[],
  lock: boolean
): void {
  if (ids.length === 0) throw new CommandRefusal('disabled.noSelection');
  if (typeof lock !== 'boolean') throw new CommandRefusal('errors.unsupported', 'lock');
  let changed = 0;
  for (const el of ids.flatMap((id) => withCopies(batch, part, treeElement(batch, part, id)))) {
    const { host, locks } = locksOf(el);
    if (!host) throw new CommandRefusal('errors.malformed');
    // An element that already has that state is left alone: no undo step without a change.
    if (aspectLocked(el) === lock) continue;
    changed += 1;
    const value = lock ? '1' : null;
    if (locks) {
      batch.push({
        op: 'setAttribute',
        part,
        node: locks.id,
        name: { ns: '', local: 'noChangeAspect' },
        value,
      });
      continue;
    }
    batch.insert({
      op: 'insertNode',
      part,
      parent: host.id,
      index: 0,
      spec: {
        ns: NS.a,
        local: NON_VISUAL[el.kind][2],
        prefix: 'a',
        attrs: [{ ns: '', local: 'noChangeAspect', value: '1' }],
      },
    });
  }
  if (changed === 0) throw new CommandRefusal('disabled.noLockChange');
}

/**
 * Show or hide elements (`p:cNvPr/@hidden`). A hidden element does not draw. Elements that
 * already have that state are left alone; when none changes, the command has nothing to do.
 */
export function setShapesVisible(
  batch: OpBatch,
  part: string,
  ids: readonly string[],
  visible: boolean
): void {
  if (!Array.isArray(ids) || ids.length === 0) throw new CommandRefusal('disabled.noSelection');
  if (typeof visible !== 'boolean') throw new CommandRefusal('errors.unsupported', 'visible');
  const changed = ids
    .flatMap((id) => withCopies(batch, part, treeElement(batch, part, id)))
    .map((el) => drawingProps(el))
    .filter((cNvPr) => {
      const v = cNvPr.attrs.find((a) => a.ns === '' && a.local === 'hidden')?.value;
      return (v === '1' || v === 'true') === visible;
    });
  if (changed.length === 0) throw new CommandRefusal('disabled.noVisibilityChange');
  for (const cNvPr of changed) {
    batch.push({
      op: 'setAttribute',
      part,
      node: cNvPr.id,
      name: { ns: '', local: 'hidden' },
      value: visible ? null : '1',
    });
  }
}

/** The longest name the selection pane writes: 255 characters. */
export const SHAPE_NAME_MAX = 255;

/** Rename an element (`p:cNvPr/@name`). Control characters and empty names are refused. */
export function renameShape(batch: OpBatch, part: string, id: string, name: string): void {
  const value = typeof name === 'string' ? name.trim() : '';
  // No control characters, noncharacters, or lone surrogates: the store refuses those in text.
  if (
    !value ||
    value.length > SHAPE_NAME_MAX ||
    /[\p{Cc}\uFFFE\uFFFF]/u.test(value) ||
    /[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?<![\uD800-\uDBFF])[\uDC00-\uDFFF]/.test(value)
  )
    throw new CommandRefusal('errors.invalidName');
  const copies = withCopies(batch, part, treeElement(batch, part, id));
  const nameOf = (el: Element): string | undefined =>
    drawingProps(el).attrs.find((a) => a.ns === '' && a.local === 'name')?.value;
  if (copies.every((el) => nameOf(el) === value)) throw new CommandRefusal('disabled.noNameChange');
  for (const el of copies) {
    const cNvPr = drawingProps(el);
    batch.push({
      op: 'setAttribute',
      part,
      node: cNvPr.id,
      name: { ns: '', local: 'name' },
      value,
    });
  }
}
