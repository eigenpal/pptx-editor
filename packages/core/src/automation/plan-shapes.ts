// Shape writes: add, delete, position and size, name, visibility, alt text, z-order, groups,
// fill, and outline. Each lowers to the editor's own edits (`commands/`) where one exists.

import { layoutSlide, createFixedMeasurer, type TextMeasurer } from '../layout/index.ts';
import {
  attr,
  child,
  elements,
  imageSize,
  isValidXmlText,
  presentationInfo,
  REL,
  sanitizeHref,
  shapeTreeOf,
  type Element,
} from '../store/index.ts';
import { renameShape, setShapesVisible } from './commands/arrange-commands.ts';
import { imageType } from './commands/fill-commands.ts';
import { groupShapes, ungroupShapes } from './commands/group-commands.ts';
import {
  addExternalRelationship,
  removeUnusedHyperlink,
  xmlAttr,
} from './commands/package-edits.ts';
import { fromBase64 } from './commands/base64.ts';
import {
  addImagePart,
  deleteShapes,
  insertLine,
  insertShape,
  insertTextBox,
  MAX_IMAGE_BYTES,
  orderShapes,
  setTransform,
  specFromXml,
} from './commands/shape-commands.ts';
import { insertTableFrame } from './commands/table-insert.ts';
import { resolveSlide } from '../cascade/index.ts';
import type {
  AutomationBox,
  AutomationFillWrite,
  AutomationLineWrite,
  AutomationNewShape,
  AutomationShapeWrite,
  AutomationZOrder,
} from './operations.ts';
import type { AutomationHandle, AutomationRef, AutomationValue } from './protocol.ts';
import { ancestorsOf, fileBox, isDecorative, readFill, readLine, slideBox } from './plan-reads.ts';
import { colorSpec, setText, type TextPlanContext } from './plan-text.ts';
import { writeCells } from './plan-tables.ts';
import {
  AutomationFailure,
  badArgument,
  cNvPrOf,
  containerAt,
  edit,
  numberIn,
  ownerSlideIndex,
  shapeAt,
  shapeHandle,
  targetOf,
  toEmu,
  type PlanContext,
  type ShapeLocation,
} from './targets.ts';

/** `ST_PositiveCoordinate` in points, for sizes and positions a write accepts. */
const MAX_POINTS = 2_147_483_647 / 12_700;

/** Map a box from slide space into a group's child space, outermost group first. */
function toChildSpace(
  groups: readonly Element[],
  box: { x: number; y: number; cx: number; cy: number }
): { x: number; y: number; cx: number; cy: number } {
  let { x, y, cx, cy } = box;
  for (const group of [...groups].reverse()) {
    const xfrm = child(child(group, 'p:grpSpPr'), 'a:xfrm');
    const n = (el: string, a: string): number => Number(attr(child(xfrm, el), a) ?? 0) || 0;
    const sx = n('a:ext', 'cx') ? n('a:chExt', 'cx') / n('a:ext', 'cx') : 1;
    const sy = n('a:ext', 'cy') ? n('a:chExt', 'cy') / n('a:ext', 'cy') : 1;
    x = n('a:chOff', 'x') + (x - n('a:off', 'x')) * sx;
    y = n('a:chOff', 'y') + (y - n('a:off', 'y')) * sy;
    cx *= sx;
    cy *= sy;
  }
  return { x, y, cx, cy };
}

/** Write a new box and rotation for a shape. Lengths in points, the rotation in degrees. */
function writeBox(
  ctx: TextPlanContext,
  at: ShapeLocation,
  box: AutomationBox,
  rotation: number | undefined
): void {
  const tree = shapeTreeOf(ctx.batch.pkg, at.part);
  const current = slideBox(ctx, at, tree);
  const value = (v: number | undefined, name: string, min: number, fallback: number): number =>
    v === undefined ? fallback : toEmu(numberIn(v, name, min, MAX_POINTS));
  const next = {
    x: value(box.left, 'left', -MAX_POINTS, current.x),
    y: value(box.top, 'top', -MAX_POINTS, current.y),
    cx: value(box.width, 'width', 0, current.cx),
    cy: value(box.height, 'height', 0, current.cy),
  };
  const file = toChildSpace(ancestorsOf(at, tree), next);
  const own = fileBox(ctx, at);
  const rot = rotation === undefined ? own.rot : numberIn(rotation, 'rotation', -1e6, 1e6);
  const turn = ((rot % 360) + 360) % 360;
  edit(() =>
    setTransform(ctx.batch, at.part, at.el.id, file.x, file.y, file.cx, file.cy, turn, {
      h: own.flipH,
      v: own.flipV,
    })
  );
}

const CONNECTORS = { straight: 'line', elbow: 'elbow', curve: 'curved' } as const;

/** The deck's slide size in EMU. */
function slideSize(ctx: TextPlanContext): { w: number; h: number } {
  const info = presentationInfo(ctx.batch.pkg);
  return { w: info.width, h: info.height };
}

export function addShape(
  ctx: TextPlanContext,
  container: AutomationRef,
  shape: AutomationNewShape,
  box: AutomationBox | undefined
): AutomationValue {
  if (typeof shape !== 'object' || shape === null) badArgument('shape');
  const target = containerAt(ctx, container);
  if (target.group)
    throw new AutomationFailure(
      'unsupported-content',
      'add a shape to a slide, a layout, or a master'
    );
  const { w, h } = slideSize(ctx);
  const part = target.part;
  const nodeId = edit(() => {
    switch (shape.kind) {
      case 'geometric':
        if (typeof shape.geometry !== 'string') badArgument('geometry');
        return insertShape(ctx.batch, part, shape.geometry, w, h);
      case 'line': {
        const kind = Object.prototype.hasOwnProperty.call(CONNECTORS, shape.connector)
          ? CONNECTORS[shape.connector]
          : badArgument('connector');
        return insertLine(ctx.batch, part, kind, w, h);
      }
      case 'textBox':
        if (typeof shape.text !== 'string') badArgument('text');
        return insertTextBox(ctx.batch, part, w, h);
      case 'table': {
        const rows = numberIn(shape.rows, 'rows', 1, 75);
        const columns = numberIn(shape.columns, 'columns', 1, 75);
        if (!Number.isInteger(rows) || !Number.isInteger(columns)) badArgument('rows');
        return insertTableFrame(ctx.batch, part, rows, columns, w, h);
      }
      default:
        return badArgument('shape');
    }
  });
  const el = ctx.batch.element(part, nodeId);
  const handle = shapeHandle(ctx, target.owner, el);
  const at = shapeAt(ctx, handle);
  if (box) writeBox(ctx, at, box, undefined);
  if (shape.kind === 'textBox' && shape.text) setText(ctx, handle, undefined, shape.text);
  if (shape.kind === 'table' && shape.values) writeCells(ctx, handle, shape.values);
  return { kind: 'handle', handle };
}

export function deleteShape(ctx: TextPlanContext, ref: AutomationRef): AutomationValue {
  const at = shapeAt(ctx, ref);
  edit(() => deleteShapes(ctx.batch, at.part, [at.el.id]));
  return { kind: 'applied' };
}

const DECORATIVE_EXT = '{C183D7F6-B498-43B3-948B-1728B52AA6E4}';
const DECORATIVE_NS = 'http://schemas.microsoft.com/office/drawing/2017/decorative';

/** Set or clear the decorative flag in `p:cNvPr/a:extLst` (the `a16`-era extension). */
function writeDecorative(ctx: TextPlanContext, at: ShapeLocation, on: boolean): void {
  const cNvPr = cNvPrOf(at.el)!;
  if (isDecorative(cNvPr) === on) return;
  const extLst = child(cNvPr, 'a:extLst');
  const ext = elements(extLst).find((e) => attr(e, 'uri') === DECORATIVE_EXT);
  if (ext) ctx.batch.push({ op: 'removeNode', part: at.part, node: ext.id });
  if (!on) return;
  const xml = `<a:ext uri="${DECORATIVE_EXT}"><adec:decorative xmlns:adec="${DECORATIVE_NS}" val="1"/></a:ext>`;
  if (extLst)
    ctx.batch.insert({
      op: 'insertNode',
      part: at.part,
      parent: extLst.id,
      index: ctx.batch.element(at.part, extLst.id).children.length,
      spec: specFromXml(xml),
    });
  else
    ctx.batch.insert({
      op: 'insertNode',
      part: at.part,
      parent: cNvPr.id,
      index: cNvPr.children.length,
      spec: specFromXml(`<a:extLst>${xml}</a:extLst>`),
    });
}

export function setShape(
  ctx: TextPlanContext,
  ref: AutomationRef,
  props: AutomationShapeWrite
): AutomationValue {
  if (typeof props !== 'object' || props === null) badArgument('props');
  const at = shapeAt(ctx, ref);
  const cNvPr = cNvPrOf(at.el);
  if (!cNvPr) throw new AutomationFailure('unsupported-content', 'a shape with no p:cNvPr');
  const boxed = ['left', 'top', 'width', 'height'].some(
    (k) => (props as Record<string, unknown>)[k] !== undefined
  );
  if (boxed || props.rotation !== undefined) {
    if (at.el.kind === 'graphicFrame' && props.rotation !== undefined)
      throw new AutomationFailure(
        'unsupported-content',
        'a graphic frame does not rotate',
        'rotation'
      );
    writeBox(ctx, at, props, props.rotation);
  }
  if (props.name !== undefined && props.name !== attr(cNvPr, 'name'))
    edit(() => renameShape(ctx.batch, at.part, at.el.id, props.name!));
  if (props.visible !== undefined) {
    if (typeof props.visible !== 'boolean') badArgument('visible');
    const hidden = attr(cNvPr, 'hidden') === '1' || attr(cNvPr, 'hidden') === 'true';
    if (hidden === props.visible)
      edit(() => setShapesVisible(ctx.batch, at.part, [at.el.id], props.visible!));
  }
  const text = (name: 'title' | 'descr', value: unknown, key: string): void => {
    if (value === undefined) return;
    if (typeof value !== 'string' || value.length > 4096 || !isValidXmlText(value))
      badArgument(key);
    ctx.batch.push({
      op: 'setAttribute',
      part: at.part,
      node: cNvPr.id,
      name: { ns: '', local: name },
      value: value === '' ? null : value,
    });
  };
  edit(() => {
    text('title', props.altTextTitle, 'altTextTitle');
    text('descr', props.altTextDescription, 'altTextDescription');
  });
  if (props.isDecorative !== undefined) {
    if (typeof props.isDecorative !== 'boolean') badArgument('isDecorative');
    edit(() => writeDecorative(ctx, shapeAt(ctx, ref), props.isDecorative!));
  }
  return { kind: 'applied' };
}

const ORDER_MOVES: Readonly<Record<AutomationZOrder, 'forward' | 'backward' | 'front' | 'back'>> = {
  bringForward: 'forward',
  sendBackward: 'backward',
  bringToFront: 'front',
  sendToBack: 'back',
};

export function setZOrder(
  ctx: TextPlanContext,
  ref: AutomationRef,
  position: AutomationZOrder
): AutomationValue {
  const at = shapeAt(ctx, ref);
  if (!Object.prototype.hasOwnProperty.call(ORDER_MOVES, position)) badArgument('position');
  const siblings = elements(at.parent).filter((e) => cNvPrOf(e));
  const index = siblings.indexOf(at.el);
  const atEnd =
    (position === 'bringForward' || position === 'bringToFront') && index === siblings.length - 1;
  const atStart = (position === 'sendBackward' || position === 'sendToBack') && index === 0;
  // Already in place: nothing to move, and not an error.
  if (atEnd || atStart) return { kind: 'applied' };
  edit(() => orderShapes(ctx.batch, at.part, [at.el.id], ORDER_MOVES[position]));
  return { kind: 'applied' };
}

/** Layout items of a slide, for the group edits' bounds. Text does not change a group's box. */
function slideItems(ctx: TextPlanContext, index: number, measurer: TextMeasurer | undefined) {
  const resolved = resolveSlide(ctx.batch.pkg, index);
  return resolved ? layoutSlide(resolved, measurer ?? createFixedMeasurer()).items : [];
}

export function group(ctx: TextPlanContext, refs: readonly AutomationRef[]): AutomationValue {
  if (!Array.isArray(refs) || refs.length > 10_000) badArgument('shapes');
  const shapes = refs.map((ref) => shapeAt(ctx, ref));
  // A group holds two or more different shapes: a repeated shape counts once.
  if (new Set(shapes.map((s) => s.el)).size !== shapes.length || shapes.length < 2)
    throw new AutomationFailure('invalid-argument', 'two or more different shapes', 'group-count');
  const first = shapes[0]!;
  if (shapes.some((s) => s.part !== first.part || s.parent !== first.parent))
    throw new AutomationFailure(
      'invalid-argument',
      'the shapes do not share a parent',
      'group-parent'
    );
  const index = ownerSlideIndex(ctx.batch.pkg, first.owner);
  if (index === undefined)
    throw new AutomationFailure('unsupported-content', 'group shapes on a slide', 'group-owner');
  const items = slideItems(ctx, index, ctx.measurer);
  const id = edit(() =>
    groupShapes(
      ctx.batch,
      first.part,
      shapes.map((s) => s.el.id),
      items
    )
  );
  return {
    kind: 'handle',
    handle: shapeHandle(ctx, first.owner, ctx.batch.element(first.part, id)),
  };
}

export function ungroup(ctx: TextPlanContext, ref: AutomationRef): AutomationValue {
  const at = shapeAt(ctx, ref);
  if (at.el.kind !== 'grpSp')
    throw new AutomationFailure('invalid-argument', 'that shape is not a group', 'group');
  const index = ownerSlideIndex(ctx.batch.pkg, at.owner);
  if (index === undefined)
    throw new AutomationFailure('unsupported-content', 'ungroup shapes on a slide', 'group-owner');
  const items = slideItems(ctx, index, ctx.measurer);
  const ids = edit(() => ungroupShapes(ctx.batch, at.part, [at.el.id], items));
  const handles: AutomationHandle[] = ids.map((id) =>
    shapeHandle(ctx, at.owner, ctx.batch.element(at.part, id))
  );
  return { kind: 'handles', handles };
}

export function parentOf(
  ctx: TextPlanContext,
  ref: AutomationRef,
  kind: 'slide' | 'slideLayout' | 'slideMaster',
  orNull?: boolean
): AutomationValue {
  const { owner } = targetOf(ctx, ref, 'shape');
  if (owner.kind === kind)
    return {
      kind: 'handle',
      handle:
        owner.kind === 'slide'
          ? ctx.handles.mint({ kind: 'slide', sldId: owner.sldId })
          : ctx.handles.mint({ kind: owner.kind, part: owner.part }),
    };
  if (orNull) return { kind: 'null' };
  throw new AutomationFailure('item-not-found', `the shape is not on a ${kind}`, kind);
}

/** The group a shape is in, or null (or `item-not-found`) for a shape at the top level. */
export function parentGroupOf(
  ctx: TextPlanContext,
  ref: AutomationRef,
  orNull?: boolean
): AutomationValue {
  const at = shapeAt(ctx, ref);
  if (at.parent.kind === 'grpSp')
    return { kind: 'handle', handle: shapeHandle(ctx, at.owner, at.parent) };
  if (orNull) return { kind: 'null' };
  throw new AutomationFailure('item-not-found', 'the shape is not in a group', 'parentGroup');
}

/**
 * Set or remove the click link of a whole shape: `a:hlinkClick` in `p:cNvPr`
 * (CT_NonVisualDrawingProps, ECMA-376 Part 1, 19.3.1.12: hlinkClick, hlinkHover, extLst).
 * The address goes through `sanitizeHref`, and the link is an external relationship.
 */
export function setShapeHyperlink(
  ctx: TextPlanContext,
  ref: AutomationRef,
  address: string | null,
  screenTip?: string
): AutomationValue {
  const at = shapeAt(ctx, ref);
  const cNvPr = cNvPrOf(at.el);
  if (!cNvPr) throw new AutomationFailure('unsupported-content', 'a shape with no p:cNvPr');
  let href: string | null = null;
  if (address !== null) {
    if (typeof address !== 'string' || address.length > 2048) badArgument('address');
    const projected = sanitizeHref(address);
    if (!projected.ok) badArgument('address');
    href = projected.href;
  }
  if (
    screenTip !== undefined &&
    (typeof screenTip !== 'string' || screenTip.length > 1024 || !isValidXmlText(screenTip))
  )
    badArgument('screenTip');
  edit(() => {
    const old = child(cNvPr, 'a:hlinkClick');
    const oldId = attr(old, 'r:id');
    if (old) ctx.batch.push({ op: 'removeNode', part: at.part, node: old.id });
    // The relationship the old link named goes once nothing else in the slide names it.
    if (oldId) removeUnusedHyperlink(ctx.batch, at.part, oldId);
    if (!href) return;
    const rId = addExternalRelationship(ctx.batch, at.part, REL.hyperlink, href);
    const tip = screenTip ? ` tooltip="${xmlAttr(screenTip, 'screenTip')}"` : '';
    ctx.batch.insert({
      op: 'insertNode',
      part: at.part,
      parent: cNvPr.id,
      index: 0,
      spec: specFromXml(`<a:hlinkClick r:id="${xmlAttr(rId, 'rId')}"${tip}/>`),
    });
  });
  return { kind: 'applied' };
}

/** The shape properties element (`p:spPr`) that holds a shape's fill and outline. */
function propsOf(at: ShapeLocation): Element {
  const props = at.el.kind === 'grpSp' ? child(at.el, 'p:grpSpPr') : child(at.el, 'p:spPr');
  if (!props)
    throw new AutomationFailure('unsupported-content', 'that shape has no fill or outline');
  return props;
}

/** EG_FillProperties. */
const FILLS = ['noFill', 'solidFill', 'gradFill', 'blipFill', 'pattFill', 'grpFill'];
/** `p:spPr` children before the fill (CT_ShapeProperties, ECMA-376 Part 1, 20.1.2.2.35). */
const BEFORE_FILL = ['xfrm', 'custGeom', 'prstGeom'];

/** Replace the fill of a shape properties element with `xml`. */
function replaceFill(ctx: TextPlanContext, part: string, props: Element, xml: string): void {
  for (const c of elements(ctx.batch.element(part, props.id)))
    if (FILLS.includes(c.local)) ctx.batch.push({ op: 'removeNode', part, node: c.id });
  let index = 0;
  ctx.batch.element(part, props.id).children.forEach((c, i) => {
    if (c.type === 'element' && BEFORE_FILL.includes(c.local)) index = i + 1;
  });
  ctx.batch.insert({ op: 'insertNode', part, parent: props.id, index, spec: specFromXml(xml) });
}

/** The longest base64 text of a picture fill: the media part cap, with room for line breaks. */
const MAX_BASE64 = Math.ceil(MAX_IMAGE_BYTES / 3) * 4 * 1.1;

/**
 * A new media part related to `part`, from base64 image bytes; returns the relationship id. The
 * bytes must decode and carry a PNG, JPEG, or GIF header.
 */
export function imageRelationship(ctx: PlanContext, part: string, base64: unknown): string {
  // Checked before the decode allocates: base64 carries 3 bytes in 4 characters.
  if (typeof base64 !== 'string' || base64.length === 0 || base64.length > MAX_BASE64)
    badArgument('base64');
  const bytes = fromBase64(base64);
  const mime = bytes ? imageType(bytes) : undefined;
  if (!bytes || !mime || !imageSize(bytes)) badArgument('base64');
  return addImagePart(ctx.batch, part, bytes, mime);
}

/** A stretched picture fill (`a:blipFill`, ECMA-376 Part 1, 20.1.8.14) over a new media part. */
function pictureFill(ctx: TextPlanContext, part: string, base64: unknown): string {
  const rId = imageRelationship(ctx, part, base64);
  return `<a:blipFill><a:blip r:embed="${rId}"/><a:stretch><a:fillRect/></a:stretch></a:blipFill>`;
}

export function setFill(
  ctx: TextPlanContext,
  ref: AutomationRef,
  fill: AutomationFillWrite
): AutomationValue {
  if (typeof fill !== 'object' || fill === null) badArgument('fill');
  const at = shapeAt(ctx, ref);
  if (at.el.kind === 'cxnSp')
    throw new AutomationFailure('unsupported-content', 'a line has no fill');
  const props = propsOf(at);
  edit(() => {
    if (fill.kind === 'clear') return replaceFill(ctx, at.part, props, '<a:noFill/>');
    if (fill.kind === 'solid')
      return replaceFill(
        ctx,
        at.part,
        props,
        `<a:solidFill>${colorSpec(fill.color, 'color', fill.transparency)}</a:solidFill>`
      );
    if (fill.kind === 'image')
      return replaceFill(ctx, at.part, props, pictureFill(ctx, at.part, fill.base64));
    if (fill.kind === 'transparency') {
      const current = readFill(ctx, ref);
      const color = current.kind === 'fill' ? current.fill.foregroundColor : null;
      if (!color)
        throw new AutomationFailure('unsupported-content', 'the fill is not solid', 'transparency');
      return replaceFill(
        ctx,
        at.part,
        props,
        `<a:solidFill>${colorSpec(color, 'color', fill.transparency)}</a:solidFill>`
      );
    }
    badArgument('fill');
  });
  return { kind: 'applied' };
}

/** `ST_PresetLineDashVal` (ECMA-376 Part 1, 20.1.10.48). */
export const DASHES = new Set([
  'solid',
  'dot',
  'dash',
  'lgDash',
  'dashDot',
  'lgDashDot',
  'lgDashDotDot',
  'sysDash',
  'sysDot',
  'sysDashDot',
  'sysDashDotDot',
]);
/** `ST_CompoundLine` (20.1.10.15). */
export const COMPOUNDS = new Set(['sng', 'dbl', 'thickThin', 'thinThick', 'tri']);
/** `a:ln` children (CT_LineProperties, 20.1.2.2.24), in schema order. */
const LN_ORDER = [
  'noFill',
  'solidFill',
  'gradFill',
  'pattFill',
  'prstDash',
  'custDash',
  'round',
  'bevel',
  'miter',
  'headEnd',
  'tailEnd',
  'extLst',
];
/** `p:spPr` children before `a:ln`. */
const BEFORE_LINE = [...BEFORE_FILL, ...FILLS];

/** Insert `xml` into `a:ln` at its schema position, replacing the same group. */
function lineChild(
  ctx: TextPlanContext,
  part: string,
  lnId: string,
  group: readonly string[],
  xml: string | null
): void {
  for (const c of elements(ctx.batch.element(part, lnId)))
    if (group.includes(c.local)) ctx.batch.push({ op: 'removeNode', part, node: c.id });
  if (!xml) return;
  const spec = specFromXml(xml);
  const rank = LN_ORDER.indexOf(spec.local);
  let index = 0;
  ctx.batch.element(part, lnId).children.forEach((c, i) => {
    if (c.type === 'element' && LN_ORDER.indexOf(c.local) < rank) index = i + 1;
  });
  ctx.batch.insert({ op: 'insertNode', part, parent: lnId, index, spec });
}

export function setLine(
  ctx: TextPlanContext,
  ref: AutomationRef,
  line: AutomationLineWrite
): AutomationValue {
  if (typeof line !== 'object' || line === null) badArgument('line');
  const at = shapeAt(ctx, ref);
  // CT_GroupShapeProperties (ECMA-376 Part 1, dml-main.xsd) has a fill but no `a:ln`.
  if (at.el.kind === 'grpSp')
    throw new AutomationFailure('unsupported-content', 'a group has no outline', 'group-line');
  const props = propsOf(at);
  const before = readLine(ctx, ref);
  const current = before.kind === 'line' ? before.line : null;
  edit(() => {
    let ln = child(ctx.batch.element(at.part, props.id), 'a:ln');
    if (!ln) {
      let index = 0;
      ctx.batch.element(at.part, props.id).children.forEach((c, i) => {
        if (c.type === 'element' && BEFORE_LINE.includes(c.local)) index = i + 1;
      });
      const id = ctx.batch.insert({
        op: 'insertNode',
        part: at.part,
        parent: props.id,
        index,
        spec: specFromXml('<a:ln/>'),
      });
      ln = ctx.batch.element(at.part, id);
    }
    const lnId = ln.id;
    const set = (name: string, value: string | null): void =>
      ctx.batch.push({
        op: 'setAttribute',
        part: at.part,
        node: lnId,
        name: { ns: '', local: name },
        value,
      });
    if (line.weight !== undefined)
      set('w', String(toEmu(numberIn(line.weight, 'weight', 0, 1584))));
    if (line.dashStyle !== undefined) {
      if (!DASHES.has(line.dashStyle)) badArgument('dashStyle');
      lineChild(
        ctx,
        at.part,
        lnId,
        ['prstDash', 'custDash'],
        `<a:prstDash val="${line.dashStyle}"/>`
      );
    }
    if (line.style !== undefined) {
      if (!COMPOUNDS.has(line.style)) badArgument('style');
      set('cmpd', line.style);
    }
    const fills = ['noFill', 'solidFill', 'gradFill', 'pattFill'];
    if (line.visible === false) lineChild(ctx, at.part, lnId, fills, '<a:noFill/>');
    else if (line.color !== undefined || line.transparency !== undefined || line.visible === true) {
      const color = line.color ?? current?.color ?? '#000000';
      const transparency = line.transparency ?? current?.transparency ?? undefined;
      lineChild(
        ctx,
        at.part,
        lnId,
        fills,
        `<a:solidFill>${colorSpec(color, 'color', transparency ?? undefined)}</a:solidFill>`
      );
    }
  });
  return { kind: 'applied' };
}
