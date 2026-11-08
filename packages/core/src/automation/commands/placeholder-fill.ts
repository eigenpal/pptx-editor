// A new table, chart, or picture takes the place of an empty placeholder of the slide: its box and its
// index, so the slide layout still positions it (ECMA-376 Part 1, 19.3.1.36 ph). DOM-free, so
// the editor commands and the surface chrome read the same rules.

import { layoutPlaceholderOf, resolveSlide } from '../../cascade/index.ts';
import {
  attr,
  boolAttr,
  child,
  elements,
  escapeXmlAttribute,
  path,
  textOf,
  xmlPartOf,
  type Element,
} from '../../store/index.ts';
import type { EditorSelection } from '../../contracts/neutral.ts';
import type { OpBatch } from './op-batch.ts';
import { spTreeOf } from './package-edits.ts';
import { removeShapes } from './shape-removal.ts';

/** The placeholder types a table fills (§19.7.10 `ST_PlaceholderType`): content and table. */
export const TABLE_PLACEHOLDERS: ReadonlySet<string> = new Set(['obj', 'tbl']);

/** The placeholder types a chart fills: content and chart. */
export const CHART_PLACEHOLDERS: ReadonlySet<string> = new Set(['obj', 'chart']);

/** The placeholder types a picture fills: content, picture, and clip art. */
export const PICTURE_PLACEHOLDERS: ReadonlySet<string> = new Set(['obj', 'pic', 'clipArt']);

/** The `p:ph` attributes a new frame keeps from the placeholder it fills. */
const KEPT_PH_ATTRS = ['type', 'orient', 'sz', 'idx'] as const;

/** Whether a text body holds no character (a field counts as one). @internal */
export function emptyText(txBody: Element | undefined): boolean {
  if (!txBody) return true;
  const walk = (el: Element, depth: number): boolean =>
    depth < 64 &&
    elements(el).every(
      (c) => !((c.local === 't' && textOf(c).length > 0) || c.local === 'fld') && walk(c, depth + 1)
    );
  return walk(txBody, 0);
}

/** The shape properties that make a placeholder hold more than its box: a fill, line, effect. */
const BOX_ONLY = new Set(['xfrm', 'prstGeom', 'custGeom']);

/**
 * The placeholder type of a top-level `p:sp` that holds nothing: no text, and no fill, line, or
 * effect of its own, which a new object would drop. Undefined for anything else. @internal
 */
export function emptyPlaceholderType(el: Element): string | undefined {
  if (el.local !== 'sp') return undefined;
  const ph = path(el, 'p:nvSpPr', 'p:nvPr', 'p:ph');
  if (!ph || boolAttr(path(el, 'p:nvSpPr', 'p:cNvPr'), 'hidden')) return undefined;
  if (elements(child(el, 'p:spPr')).some((c) => !BOX_ONLY.has(c.local))) return undefined;
  return emptyText(child(el, 'p:txBody')) ? (attr(ph, 'type') ?? 'obj') : undefined;
}

/**
 * The placeholder a selection points at: one selected shape, or the shape that holds the caret.
 * @internal
 */
export function selectedPlaceholder(sel: EditorSelection): string | undefined {
  if (sel.kind === 'shapes') return sel.nodeIds.length === 1 ? sel.nodeIds[0] : undefined;
  return sel.kind === 'text' && !sel.cell ? sel.nodeId : undefined;
}

/** An empty placeholder of the slide that a new table or picture can take the place of. */
export interface EmptyPlaceholder {
  /** The `p:sp` node id. */
  readonly nodeId: string;
  /**
   * Its `p:cNvPr` id: the animations and connector ends that name it go on to name the new
   * object (`removeFilledPlaceholder`).
   */
  readonly shapeId: string;
  /** Its `p:cNvPr` name, which the new object keeps. */
  readonly name: string;
  /** Its geometry (`a:prstGeom` or `a:custGeom`), from the slide or the layout, if any. */
  readonly geometry?: Element;
  /** Its placeholder type; `obj` when the file names none. */
  readonly type: string;
  /** Its box on the slide, EMU, from the slide, layout, or master. */
  readonly x: number;
  readonly y: number;
  readonly cx: number;
  readonly cy: number;
  /** The `p:ph` element of the new frame, with the placeholder's own type, size, and index. */
  readonly ph: string;
}

/**
 * The empty placeholder of one of `types` that a new table or picture fills: `preferred` when it
 * is one, otherwise the first on the slide. A hidden placeholder, one with text or formatting of
 * its own, and a turned or flipped one are left alone. Undefined when the slide has none.
 */
export function emptyPlaceholderOf(
  batch: OpBatch,
  part: string,
  slideIndex: number,
  types: ReadonlySet<string>,
  preferred?: string
): EmptyPlaceholder | undefined {
  const resolved = resolveSlide(batch.pkg, slideIndex)?.elements;
  // A layout or master element can reuse a slide element's node id: only the slide's counts.
  const boxOf = (id: string) => resolved?.find((r) => r.layer === 'slide' && r.nodeId === id)?.xfrm;
  const candidates = elements(spTreeOf(batch, part)).filter((el) => {
    const type = emptyPlaceholderType(el);
    if (type === undefined || !types.has(type)) return false;
    const box = boxOf(el.id);
    return !!box && box.rot === 0 && !box.flipH && !box.flipV && box.cx > 0 && box.cy > 0;
  });
  if (!candidates.length) return undefined;
  const el = candidates.find((c) => c.id === preferred) ?? candidates[0]!;
  const box = boxOf(el.id)!;
  const ph = path(el, 'p:nvSpPr', 'p:nvPr', 'p:ph')!;
  const nv = path(el, 'p:nvSpPr', 'p:cNvPr');
  const geometryOf = (sp: Element | undefined) =>
    elements(child(sp, 'p:spPr')).find((c) => c.local === 'prstGeom' || c.local === 'custGeom');
  const geometry = geometryOf(el) ?? geometryOf(layoutPlaceholderOf(batch.pkg, slideIndex, el));
  const kept = KEPT_PH_ATTRS.map((name) => [name, attr(ph, name)] as const)
    .filter(([, v]) => v !== undefined)
    .map(([name, v]) => ` ${name}="${escapeXmlAttribute(v!)}"`)
    .join('');
  return {
    nodeId: el.id,
    shapeId: attr(nv, 'id') ?? '',
    name: attr(nv, 'name') ?? '',
    type: attr(ph, 'type') ?? 'obj',
    x: box.x,
    y: box.y,
    cx: box.cx,
    cy: box.cy,
    ph: `<p:ph${kept}/>`,
    ...(geometry ? { geometry } : {}),
  };
}

/** The step a new frame moves by when its place is taken: 5 % of the slide, down and right. */
const INSERT_STEP = 0.05;

/**
 * Where a new frame of `cx` by `cy` goes, from the place (`x`, `y`). When a top-level frame or
 * picture covers more than half of it, it starts one step past that frame's corner, and it
 * moves one step more for each frame that starts there: objects inserted one after another
 * cascade, and each one shows its own corner. Text shapes do not count, so a new object can sit
 * on an empty text placeholder. It stays on the slide.
 */
export function freeInsertPosition(
  tree: Element,
  x: number,
  y: number,
  size: {
    readonly cx: number;
    readonly cy: number;
    readonly slideW: number;
    readonly slideH: number;
  }
): { x: number; y: number } {
  const boxes: { x: number; y: number; cx: number; cy: number }[] = [];
  for (const el of elements(tree)) {
    if (el.local !== 'graphicFrame' && el.local !== 'pic') continue;
    const xfrm = path(el, 'p:xfrm') ?? path(el, 'p:spPr', 'a:xfrm');
    const off = path(xfrm, 'a:off');
    const ext = path(xfrm, 'a:ext');
    if (!off || !ext) continue;
    const n = (v: string | undefined) => Number(v ?? 0) || 0;
    boxes.push({
      x: n(attr(off, 'x')),
      y: n(attr(off, 'y')),
      cx: n(attr(ext, 'cx')),
      cy: n(attr(ext, 'cy')),
    });
  }
  const area = Math.max(1, size.cx * size.cy);
  const dx = Math.round(size.slideW * INSERT_STEP);
  const dy = Math.round(size.slideH * INSERT_STEP);
  const fits = (px: number, py: number) =>
    px + size.cx <= size.slideW && py + size.cy <= size.slideH;
  // A frame that covers more than half of the new one: start one step past its corner.
  const cover = boxes.find((b) => {
    const w = Math.min(x + size.cx, b.x + b.cx) - Math.max(x, b.x);
    const h = Math.min(y + size.cy, b.y + b.cy) - Math.max(y, b.y);
    return w > 0 && h > 0 && w * h > area / 2;
  });
  if (cover && fits(cover.x + dx, cover.y + dy)) {
    x = cover.x + dx;
    y = cover.y + dy;
  }
  // Then one step more for each frame that starts at about the same point: a cascade.
  for (let i = 0; i < boxes.length; i++) {
    const near = boxes.some((b) => Math.abs(b.x - x) < dx / 2 && Math.abs(b.y - y) < dy / 2);
    if (!near || !fits(x + dx, y + dy)) break;
    x += dx;
    y += dy;
  }
  return { x, y };
}

/**
 * The tree index where the frame that fills `into` goes: the placeholder's own place, so the
 * z-order stays. -1 when the placeholder is gone.
 */
export function placeholderIndex(batch: OpBatch, part: string, into: EmptyPlaceholder): number {
  return spTreeOf(batch, part).children.findIndex(
    (c) => c.type === 'element' && c.id === into.nodeId
  );
}

/**
 * Remove the placeholder that the new object `shapeId` filled. The animations and connector
 * ends that name the placeholder name the object instead, so they stay. A text build by
 * paragraph (`p:bldP`) goes: the object has no paragraphs. The object has its own id: two
 * people who fill one placeholder at once make two objects with two ids.
 */
export function removeFilledPlaceholder(
  batch: OpBatch,
  part: string,
  into: EmptyPlaceholder,
  shapeId: string
): void {
  const root = xmlPartOf(batch.pkg, part)?.root;
  const set = (el: Element, name: string): void =>
    batch.push({
      op: 'setAttribute',
      part,
      node: el.id,
      name: { ns: '', local: name },
      value: shapeId,
    });
  // Shape ids are unsigned integers: "0005" names shape 5.
  const old = Number(into.shapeId);
  const names = (v: string | undefined) => v !== undefined && v !== '' && Number(v) === old;
  // Only animations (`p:timing`, also in `mc:AlternateContent` copies) and connector ends name
  // a shape by id here; other `spid` attributes (`p:oleObj`) mean something else.
  const walk = (el: Element, depth: number, timing: boolean): void => {
    if (depth > 64) return;
    for (const c of elements(el)) {
      const inTiming = timing || c.local === 'timing';
      if (inTiming && c.local === 'bldP' && names(attr(c, 'spid'))) {
        batch.push({ op: 'removeNode', part, node: c.id });
        continue;
      }
      if (inTiming && names(attr(c, 'spid'))) set(c, 'spid');
      if ((c.local === 'stCxn' || c.local === 'endCxn') && names(attr(c, 'id'))) set(c, 'id');
      walk(c, depth + 1, inTiming);
    }
  };
  if (root && into.shapeId) walk(root, 0, false);
  removeShapes(batch, part, [into.nodeId]);
}

/**
 * The empty placeholder a new table, chart, or picture fills on slide `slideIndex`: the one that
 * is selected or holds the caret, or else the first one that takes it. @internal
 */
export function placeholderFor(
  batch: OpBatch,
  part: string,
  slideIndex: number,
  selection: EditorSelection,
  what: 'table' | 'chart' | 'picture'
): EmptyPlaceholder | undefined {
  const types =
    what === 'table'
      ? TABLE_PLACEHOLDERS
      : what === 'chart'
        ? CHART_PLACEHOLDERS
        : PICTURE_PLACEHOLDERS;
  return emptyPlaceholderOf(batch, part, slideIndex, types, selectedPlaceholder(selection));
}
