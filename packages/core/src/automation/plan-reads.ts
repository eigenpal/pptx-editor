// The read operations of presentations, slides, layouts, masters, and shapes.
//
// Effective values come from the cascade (`resolveSlide`), so a placeholder that inherits its
// position, fill, or line from its layout reads as it is drawn. A shape on a layout or master
// reads its own values from the file: the cascade resolves slides only.

import { lineEndType } from './commands/style-commands.ts';
import {
  resolveSlide,
  toHex,
  type ResolvedElement,
  type ResolvedFill,
  type ResolvedLine,
  type ResolvedSlide,
  type Xfrm,
} from '../cascade/index.ts';
import {
  attr,
  child,
  elements,
  intAttr,
  path,
  presentationInfo,
  shapeTreeOf,
  xmlPartOf,
  type Element,
  type PptxPackage,
} from '../store/index.ts';
import type { AutomationItemKey } from './operations.ts';
import type {
  AutomationFillRead,
  AutomationHandle,
  AutomationLineRead,
  AutomationShapeRead,
  AutomationValue,
} from './protocol.ts';
import {
  AutomationFailure,
  cNvPrOf,
  containerAt,
  layoutsOf,
  mastersOf,
  masterOfLayout,
  shapeAt,
  shapeHandle,
  shapesIn,
  slideOf,
  targetOf,
  toPoints,
  type PlanContext,
  type ShapeLocation,
} from './targets.ts';
import { corePropertiesOf } from './plan-properties.ts';

const resolvedCache = new WeakMap<PptxPackage, Map<number, ResolvedSlide | undefined>>();

/** The resolved slide at `index`, memoized per working package. */
export function resolvedSlide(pkg: PptxPackage, index: number): ResolvedSlide | undefined {
  let byIndex = resolvedCache.get(pkg);
  if (!byIndex) resolvedCache.set(pkg, (byIndex = new Map()));
  if (!byIndex.has(index)) byIndex.set(index, resolveSlide(pkg, index));
  return byIndex.get(index);
}

/** The resolved element of a slide shape, or undefined on a layout or master. */
export function resolvedShape(ctx: PlanContext, at: ShapeLocation): ResolvedElement | undefined {
  if (at.owner.kind !== 'slide') return undefined;
  const slide = resolvedSlide(ctx.batch.pkg, slideOf(ctx.batch.pkg, at.owner.sldId).index);
  const visit = (list: readonly ResolvedElement[]): ResolvedElement | undefined => {
    for (const el of list) {
      if (el.nodeId === at.el.id && el.partName === at.part) return el;
      if (el.kind === 'group') {
        const hit = visit(el.children);
        if (hit) return hit;
      }
    }
    return undefined;
  };
  return slide ? visit(slide.elements) : undefined;
}

const handles = (list: readonly AutomationHandle[]): AutomationValue => ({
  kind: 'handles',
  handles: list,
});

/** Pick an item by index or id, or answer null / fail. */
function pick<T>(
  items: readonly T[],
  key: AutomationItemKey,
  idOf: (item: T) => string,
  orNull: boolean | undefined
): T | null {
  const hit =
    'index' in key
      ? Number.isInteger(key.index) && key.index >= 0
        ? items[key.index]
        : undefined
      : typeof key.id === 'string'
        ? items.find((item) => idOf(item) === key.id)
        : undefined;
  if (hit !== undefined) return hit;
  if (orNull) return null;
  throw new AutomationFailure('item-not-found', 'no such item', 'index' in key ? 'index' : 'id');
}

export function listSlides(ctx: PlanContext): AutomationValue {
  return handles(
    presentationInfo(ctx.batch.pkg).slides.map((s) =>
      ctx.handles.mint({ kind: 'slide', sldId: s.sldId })
    )
  );
}

export function getSlide(
  ctx: PlanContext,
  key: AutomationItemKey,
  orNull?: boolean
): AutomationValue {
  const slide = pick(presentationInfo(ctx.batch.pkg).slides, key, (s) => s.sldId, orNull);
  return slide
    ? { kind: 'handle', handle: ctx.handles.mint({ kind: 'slide', sldId: slide.sldId }) }
    : { kind: 'null' };
}

export function readSlide(ctx: PlanContext, ref: Parameters<typeof targetOf>[1]): AutomationValue {
  const { sldId } = targetOf(ctx, ref, 'slide');
  const { ref: slide, index } = slideOf(ctx.batch.pkg, sldId);
  const master =
    slide.masterPart ?? (slide.layoutPart && masterOfLayout(ctx.batch.pkg, slide.layoutPart));
  return {
    kind: 'slide',
    slide: {
      id: sldId,
      index,
      hidden: slide.hidden,
      layout: slide.layoutPart
        ? ctx.handles.mint({ kind: 'slideLayout', part: slide.layoutPart })
        : null,
      master: master ? ctx.handles.mint({ kind: 'slideMaster', part: master }) : null,
    },
  };
}

export function readPresentation(ctx: PlanContext, id: string): AutomationValue {
  const info = presentationInfo(ctx.batch.pkg);
  return {
    kind: 'presentation',
    presentation: {
      id,
      title: corePropertiesOf(ctx.batch.pkg).title,
      slideWidth: toPoints(info.width),
      slideHeight: toPoints(info.height),
    },
  };
}

export function listMasters(ctx: PlanContext): AutomationValue {
  return handles(
    mastersOf(ctx.batch.pkg).map((m) => ctx.handles.mint({ kind: 'slideMaster', part: m.part }))
  );
}

export function getMaster(
  ctx: PlanContext,
  key: AutomationItemKey,
  orNull?: boolean
): AutomationValue {
  const master = pick(mastersOf(ctx.batch.pkg), key, (m) => m.id, orNull);
  return master
    ? { kind: 'handle', handle: ctx.handles.mint({ kind: 'slideMaster', part: master.part }) }
    : { kind: 'null' };
}

const nameOf = (pkg: PptxPackage, part: string): string =>
  attr(path(ctxRoot(pkg, part), 'p:cSld'), 'name') ?? '';

function ctxRoot(pkg: PptxPackage, part: string): Element | undefined {
  return xmlPartOf(pkg, part)?.root;
}

export function readMaster(ctx: PlanContext, ref: Parameters<typeof targetOf>[1]): AutomationValue {
  const { part } = targetOf(ctx, ref, 'slideMaster');
  const entry = mastersOf(ctx.batch.pkg).find((m) => m.part === part);
  if (!entry) throw new AutomationFailure('invalid-handle', 'that master is gone', 'gone');
  return {
    kind: 'layout',
    layout: { id: entry.id, name: nameOf(ctx.batch.pkg, part), type: null, master: null },
  };
}

export function listLayouts(
  ctx: PlanContext,
  ref: Parameters<typeof targetOf>[1]
): AutomationValue {
  const { part } = targetOf(ctx, ref, 'slideMaster');
  return handles(
    layoutsOf(ctx.batch.pkg, part).map((l) =>
      ctx.handles.mint({ kind: 'slideLayout', part: l.part })
    )
  );
}

export function getLayout(
  ctx: PlanContext,
  ref: Parameters<typeof targetOf>[1],
  key: AutomationItemKey,
  orNull?: boolean
): AutomationValue {
  const { part } = targetOf(ctx, ref, 'slideMaster');
  const layout = pick(layoutsOf(ctx.batch.pkg, part), key, (l) => l.id, orNull);
  return layout
    ? { kind: 'handle', handle: ctx.handles.mint({ kind: 'slideLayout', part: layout.part }) }
    : { kind: 'null' };
}

export function readLayout(ctx: PlanContext, ref: Parameters<typeof targetOf>[1]): AutomationValue {
  const { part } = targetOf(ctx, ref, 'slideLayout');
  const master = masterOfLayout(ctx.batch.pkg, part);
  const entry = master ? layoutsOf(ctx.batch.pkg, master).find((l) => l.part === part) : undefined;
  const root = ctxRoot(ctx.batch.pkg, part);
  if (!root) throw new AutomationFailure('invalid-handle', 'that layout is gone', 'gone');
  return {
    kind: 'layout',
    layout: {
      id: entry?.id ?? '',
      name: nameOf(ctx.batch.pkg, part),
      type: attr(root, 'type') ?? 'cust',
      master: master ? ctx.handles.mint({ kind: 'slideMaster', part: master }) : null,
    },
  };
}

export function listShapes(
  ctx: PlanContext,
  ref: Parameters<typeof targetOf>[1],
  orNull?: boolean
): AutomationValue {
  if (
    orNull &&
    targetOf(ctx, ref, 'slide', 'slideLayout', 'slideMaster', 'shape').kind === 'shape' &&
    !isGroup(ctx, ref, true)
  )
    return { kind: 'null' };
  const { owner, tree } = containerAt(ctx, ref);
  return handles(shapesIn(tree).map((el) => shapeHandle(ctx, owner, el)));
}

/**
 * Whether a shape is a group (`p:grpSp`). A shape that is not one refuses (`not-group`) unless
 * `orNull` is set, which answers false.
 */
export function isGroup(
  ctx: PlanContext,
  ref: Parameters<typeof targetOf>[1],
  orNull?: boolean
): boolean {
  if (shapeAt(ctx, ref).el.kind === 'grpSp') return true;
  if (orNull) return false;
  throw new AutomationFailure('invalid-handle', 'that shape is not a group', 'not-group');
}

export function getShape(
  ctx: PlanContext,
  ref: Parameters<typeof targetOf>[1],
  key: AutomationItemKey,
  orNull?: boolean
): AutomationValue {
  const { owner, tree } = containerAt(ctx, ref);
  const shape = pick(shapesIn(tree), key, (el) => attr(cNvPrOf(el), 'id') ?? '', orNull);
  return shape ? { kind: 'handle', handle: shapeHandle(ctx, owner, shape) } : { kind: 'null' };
}

/** Graphic frame kinds by `a:graphicData/@uri`. Null prototype: the uri comes from the file. */
const FRAME_TYPES: Readonly<Record<string, string>> = Object.assign(Object.create(null), {
  'http://schemas.openxmlformats.org/drawingml/2006/table': 'table',
  'http://schemas.openxmlformats.org/drawingml/2006/chart': 'chart',
  'http://schemas.openxmlformats.org/drawingml/2006/diagram': 'diagram',
  'http://schemas.openxmlformats.org/presentationml/2006/ole': 'ole',
});

/** The object model's shape type of an element. */
export function shapeType(el: Element): string {
  switch (el.kind) {
    case 'grpSp':
      return 'group';
    case 'cxnSp':
      return 'line';
    case 'pic': {
      const nvPr = path(el, 'p:nvPicPr', 'p:nvPr');
      const media = elements(nvPr).some((c) =>
        ['videoFile', 'audioFile', 'quickTimeFile', 'wavAudioFile'].includes(c.local)
      );
      return media ? 'media' : 'image';
    }
    case 'graphicFrame':
      return (
        FRAME_TYPES[attr(path(el, 'a:graphic', 'a:graphicData'), 'uri') ?? ''] ?? 'unsupported'
      );
    case 'sp': {
      if (path(el, 'p:nvSpPr', 'p:nvPr', 'p:ph')) return 'placeholder';
      if (attr(path(el, 'p:nvSpPr', 'p:cNvSpPr'), 'txBox') === '1') return 'textBox';
      const spPr = child(el, 'p:spPr');
      if (child(spPr, 'a:custGeom')) return 'freeform';
      if (attr(child(spPr, 'a:prstGeom'), 'prst') === 'line') return 'line';
      return 'geometricShape';
    }
    default:
      return 'unsupported';
  }
}

/** The transform element a shape stores (`a:xfrm` in its properties, or `p:xfrm`). */
export function ownXfrm(el: Element): Element | undefined {
  if (el.kind === 'graphicFrame') return child(el, 'p:xfrm');
  const props = el.kind === 'grpSp' ? child(el, 'p:grpSpPr') : child(el, 'p:spPr');
  return child(props, 'a:xfrm');
}

/** A shape's box in its parent's coordinates, in EMU, from the file or the cascade. */
export function fileBox(ctx: PlanContext, at: ShapeLocation): Xfrm {
  const xfrm = ownXfrm(at.el);
  const off = child(xfrm, 'a:off');
  const ext = child(xfrm, 'a:ext');
  if (off && ext)
    return {
      x: intAttr(off, 'x') ?? 0,
      y: intAttr(off, 'y') ?? 0,
      cx: intAttr(ext, 'cx') ?? 0,
      cy: intAttr(ext, 'cy') ?? 0,
      rot: (intAttr(xfrm, 'rot') ?? 0) / 60_000,
      flipH: attr(xfrm, 'flipH') === '1',
      flipV: attr(xfrm, 'flipV') === '1',
    };
  const resolved = resolvedShape(ctx, at);
  return resolved?.xfrm ?? { x: 0, y: 0, cx: 0, cy: 0, rot: 0, flipH: false, flipV: false };
}

/** The group transform from a group's child space to its parent's space. */
function groupMap(group: Element): (x: number, y: number, cx: number, cy: number) => number[] {
  const xfrm = child(child(group, 'p:grpSpPr'), 'a:xfrm');
  const n = (el: string, a: string): number => intAttr(child(xfrm, el), a) ?? 0;
  const sx = n('a:chExt', 'cx') ? n('a:ext', 'cx') / n('a:chExt', 'cx') : 1;
  const sy = n('a:chExt', 'cy') ? n('a:ext', 'cy') / n('a:chExt', 'cy') : 1;
  return (x, y, cx, cy) => [
    n('a:off', 'x') + (x - n('a:chOff', 'x')) * sx,
    n('a:off', 'y') + (y - n('a:chOff', 'y')) * sy,
    cx * sx,
    cy * sy,
  ];
}

/** The groups that hold a shape, innermost first. */
export function ancestorsOf(at: ShapeLocation, tree: Element | undefined): Element[] {
  const out: Element[] = [];
  const visit = (parent: Element, chain: Element[]): boolean => {
    for (const el of shapesIn(parent)) {
      if (el === at.el) {
        for (let i = chain.length - 1; i >= 0; i -= 1) out.push(chain[i]!);
        return true;
      }
      if (el.kind === 'grpSp' && visit(el, [...chain, el])) return true;
    }
    return false;
  };
  if (tree) visit(tree, []);
  return out;
}

/**
 * A shape's box on the slide, in EMU: the file box mapped out of every group that holds it.
 * Group rotation is not applied (the object model reports an unrotated box).
 */
export function slideBox(ctx: PlanContext, at: ShapeLocation, tree: Element | undefined): Xfrm {
  const box = fileBox(ctx, at);
  let [x, y, cx, cy] = [box.x, box.y, box.cx, box.cy];
  for (const group of ancestorsOf(at, tree)) [x, y, cx, cy] = groupMap(group)(x!, y!, cx!, cy!);
  return { ...box, x: x!, y: y!, cx: cx!, cy: cy! };
}

const DECORATIVE_EXT = '{C183D7F6-B498-43B3-948B-1728B52AA6E4}';

/** The decorative flag (`adec:decorative/@val`) in `p:cNvPr/a:extLst`. */
export function isDecorative(cNvPr: Element | undefined): boolean {
  const ext = elements(child(cNvPr, 'a:extLst')).find((e) => attr(e, 'uri') === DECORATIVE_EXT);
  const flag = elements(ext).find((e) => e.local === 'decorative');
  return attr(flag, 'val') === '1';
}

export function readShape(ctx: PlanContext, ref: Parameters<typeof targetOf>[1]): AutomationValue {
  const at = shapeAt(ctx, ref);
  const cNvPr = cNvPrOf(at.el);
  const box = slideBox(ctx, at, shapeTreeOf(ctx.batch.pkg, at.part));
  const ph = path(at.el, 'p:nvSpPr', 'p:nvPr', 'p:ph') ?? path(nvPrOf(at.el), 'p:ph');
  const type = shapeType(at.el);
  const shape: AutomationShapeRead = {
    id: at.shapeId,
    name: attr(cNvPr, 'name') ?? '',
    type,
    left: toPoints(box.x),
    top: toPoints(box.y),
    width: toPoints(box.cx),
    height: toPoints(box.cy),
    rotation: Math.round(box.rot * 1000) / 1000,
    visible: attr(cNvPr, 'hidden') !== '1' && attr(cNvPr, 'hidden') !== 'true',
    zOrderPosition: shapesIn(at.parent).indexOf(at.el),
    level: at.level,
    geometry: attr(path(at.el, 'p:spPr', 'a:prstGeom'), 'prst') ?? null,
    altTextTitle: attr(cNvPr, 'title') ?? '',
    altTextDescription: attr(cNvPr, 'descr') ?? '',
    isDecorative: isDecorative(cNvPr),
    hasTextFrame: Boolean(child(at.el, 'p:txBody')),
    placeholder: ph
      ? {
          type: attr(ph, 'type') ?? 'obj',
          containedType:
            at.el.kind === 'sp' ? null : type === 'placeholder' ? null : shapeType(at.el),
        }
      : null,
    parentGroup: at.parent.kind === 'grpSp' ? shapeHandle(ctx, at.owner, at.parent) : null,
  };
  return { kind: 'shape', shape };
}

/** The `p:nvPr` of any shape kind. */
function nvPrOf(el: Element): Element | undefined {
  return child(
    elements(el).find((c) => c.local.startsWith('nv') && c.local.endsWith('Pr')),
    'p:nvPr'
  );
}

const OPAQUE = 100_000;

/** The `a:alpha` of a file color as a transparency from 0 to 1, or null. */
function fileTransparency(color: Element | undefined): number | null {
  const alpha = intAttr(child(color, 'a:alpha'), 'val');
  return alpha === undefined ? 0 : Math.round((1 - alpha / OPAQUE) * 1000) / 1000;
}

/** The first color element of a fill element (`a:srgbClr`, `a:schemeClr`, …). */
function colorElement(fill: Element | undefined): Element | undefined {
  return elements(fill).find((c) => c.local.endsWith('Clr'));
}

const hex = (rgba: { r: number; g: number; b: number }): string =>
  toHex({ ...rgba, a: 1 }).toUpperCase();

const transparencyOf = (a: number): number => Math.round((1 - a) * 1000) / 1000;

/** The object model's view of a resolved fill. */
export function fillRead(fill: ResolvedFill): AutomationFillRead {
  switch (fill.type) {
    case 'none':
      return { type: 'noFill', foregroundColor: null, transparency: null };
    case 'solid':
      return {
        type: 'solid',
        foregroundColor: hex(fill.color),
        transparency: transparencyOf(fill.color.a),
      };
    case 'linear':
    case 'path':
      return {
        type: 'gradient',
        foregroundColor: fill.stops[0] ? hex(fill.stops[0].color) : null,
        transparency: null,
      };
    case 'pattern':
      return {
        type: 'pattern',
        foregroundColor: hex(fill.fg),
        transparency: transparencyOf(fill.fg.a),
      };
    case 'picture':
      return { type: 'pictureAndTexture', foregroundColor: null, transparency: null };
    case 'group':
      return { type: 'group', foregroundColor: null, transparency: null };
  }
}

/** The fill a shape stores in its properties, for a shape the cascade does not resolve. */
function fileFill(el: Element): AutomationFillRead {
  const props = el.kind === 'grpSp' ? child(el, 'p:grpSpPr') : child(el, 'p:spPr');
  for (const c of elements(props)) {
    if (c.local === 'noFill') return { type: 'noFill', foregroundColor: null, transparency: null };
    if (c.local === 'solidFill') {
      const color = colorElement(c);
      const val = attr(color, 'val');
      return {
        type: 'solid',
        foregroundColor: color?.local === 'srgbClr' && val ? `#${val.toUpperCase()}` : null,
        transparency: fileTransparency(color),
      };
    }
    if (c.local === 'gradFill')
      return { type: 'gradient', foregroundColor: null, transparency: null };
    if (c.local === 'pattFill')
      return { type: 'pattern', foregroundColor: null, transparency: null };
    if (c.local === 'blipFill')
      return { type: 'pictureAndTexture', foregroundColor: null, transparency: null };
    if (c.local === 'grpFill') return { type: 'group', foregroundColor: null, transparency: null };
  }
  return { type: 'noFill', foregroundColor: null, transparency: null };
}

export function readFill(ctx: PlanContext, ref: Parameters<typeof targetOf>[1]): AutomationValue {
  const at = shapeAt(ctx, ref);
  const resolved = resolvedShape(ctx, at);
  const fill =
    resolved && resolved.kind !== 'graphicFrame' ? fillRead(resolved.fill) : fileFill(at.el);
  return { kind: 'fill', fill };
}

/** The object model's view of a resolved line. */
export function lineRead(line: ResolvedLine | undefined): AutomationLineRead {
  if (!line || line.fill.type === 'none')
    return {
      head: lineEndType(line?.head?.type),
      tail: lineEndType(line?.tail?.type),
      visible: false,
      color: null,
      weight: line ? toPoints(line.width) : null,
      dashStyle: line?.dash ?? null,
      style: line?.compound ?? null,
      transparency: null,
    };
  const solid = line.fill.type === 'solid' ? line.fill : null;
  return {
    head: lineEndType(line.head?.type),
    tail: lineEndType(line.tail?.type),
    visible: true,
    color: solid ? hex(solid.color) : null,
    weight: toPoints(line.width),
    dashStyle: line.dash,
    style: line.compound,
    transparency: solid ? transparencyOf(solid.color.a) : null,
  };
}

/** The line a shape stores, for a shape the cascade does not resolve. */
function fileLine(el: Element): AutomationLineRead {
  return lineElementRead(path(el, 'p:spPr', 'a:ln'));
}

/** A stored `a:ln` (or a cell's `a:lnL` and the like, CT_LineProperties) as plain values. */
export function lineElementRead(ln: Element | undefined): AutomationLineRead {
  const solid = child(ln, 'a:solidFill');
  const color = colorElement(solid);
  const val = attr(color, 'val');
  const w = intAttr(ln, 'w');
  return {
    head: lineEndType(attr(child(ln, 'a:headEnd'), 'type')),
    tail: lineEndType(attr(child(ln, 'a:tailEnd'), 'type')),
    visible: Boolean(ln) && !child(ln, 'a:noFill'),
    color: color?.local === 'srgbClr' && val ? `#${val.toUpperCase()}` : null,
    weight: w === undefined ? null : toPoints(w),
    dashStyle: attr(child(ln, 'a:prstDash'), 'val') ?? null,
    style: attr(ln, 'cmpd') ?? null,
    transparency: solid ? fileTransparency(color) : null,
  };
}

export function readLine(ctx: PlanContext, ref: Parameters<typeof targetOf>[1]): AutomationValue {
  const at = shapeAt(ctx, ref);
  const resolved = resolvedShape(ctx, at);
  const line =
    resolved &&
    (resolved.kind === 'shape' || resolved.kind === 'connector' || resolved.kind === 'picture')
      ? lineRead(resolved.line)
      : fileLine(at.el);
  return { kind: 'line', line };
}
