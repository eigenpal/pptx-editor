// Shape commands: transform, insert (text box, preset shape, picture), duplicate, delete,
// z-order, and object alignment. Each compiles to ops on an OpBatch against one slide part.

import {
  placeholderIndex,
  removeFilledPlaceholder,
  type EmptyPlaceholder,
} from './placeholder-fill.ts';
import {
  attr,
  child,
  elements,
  escapeXmlAttribute,
  NS,
  parseXmlPart,
  REL,
  specOf,
  nodeIndex,
  xmlPartOf,
  type Element,
  type NodeSpec,
} from '../../store/index.ts';
import type { LineKind, ObjectAlign, OrderMove } from '../../contracts/neutral.ts';
import { toBase64 } from './base64.ts';
import { CommandRefusal, type OpBatch } from './op-batch.ts';
import { isInsertablePreset } from './shape-presets.ts';
import { removeShapes } from './shape-removal.ts';
import {
  addRelationship,
  coordinate,
  ensureContentType,
  freePartName,
  nextShapeId,
  shapeIds,
  parentOf,
  positiveCoordinate,
  renumberShapes,
  shapeChildren,
  spTreeOf,
} from './package-edits.ts';

const WRAP = `xmlns:a="${NS.a}" xmlns:p="${NS.p}" xmlns:r="${NS.r}"`;

/** Parse an XML fragment (with a, p, r prefixes bound) into a node spec. */
export function specFromXml(xml: string): NodeSpec {
  const r = parseXmlPart('/fragment.xml', `<w ${WRAP}>${xml}</w>`);
  const first = r.ok ? elements(r.part.root)[0] : undefined;
  if (!first) throw new CommandRefusal('errors.unsupported', 'template');
  return specOf(first);
}

const int = (v: number): string => String(Math.round(v));

/** The element holding a shape's `a:xfrm`: `p:spPr`, `p:grpSpPr`, or the frame's own `p:xfrm`. */
function transformHost(el: Element): Element | undefined {
  if (el.kind === 'grpSp') return child(el, 'p:grpSpPr');
  if (el.kind === 'graphicFrame') return el;
  return child(el, 'p:spPr');
}

/** Write an absolute transform (EMU, degrees) onto a shape, creating `a:xfrm` when inherited. */
export function setTransform(
  batch: OpBatch,
  part: string,
  nodeId: string,
  x: number,
  y: number,
  cx: number,
  cy: number,
  rot: number,
  flips?: { readonly h: boolean; readonly v: boolean }
): void {
  const el = batch.element(part, nodeId);
  const host = transformHost(el);
  if (!host) throw new CommandRefusal('disabled.unsupportedElement');
  const xfrm = el.kind === 'graphicFrame' ? child(el, 'p:xfrm') : child(host, 'a:xfrm');
  if (!Number.isFinite(rot)) throw new CommandRefusal('errors.unsupported', 'rot');
  const rotValue = ((Math.round(rot * 60_000) % 21_600_000) + 21_600_000) % 21_600_000;
  // Checked before any op: a value out of `ST_Coordinate` refuses the whole edit.
  const v = {
    x: coordinate(x, 'x'),
    y: coordinate(y, 'y'),
    cx: positiveCoordinate(Math.max(0, cx), 'cx'),
    cy: positiveCoordinate(Math.max(0, cy), 'cy'),
  };
  if (!xfrm) {
    if (el.kind === 'graphicFrame') throw new CommandRefusal('disabled.unsupportedElement');
    batch.insert({
      op: 'insertNode',
      part,
      parent: host.id,
      index: 0,
      spec: specFromXml(
        `<a:xfrm${rotValue ? ` rot="${rotValue}"` : ''}${flips?.h ? ' flipH="1"' : ''}${flips?.v ? ' flipV="1"' : ''}><a:off x="${v.x}" y="${v.y}"/><a:ext cx="${v.cx}" cy="${v.cy}"/></a:xfrm>`
      ),
    });
    return;
  }
  const off = child(xfrm, 'a:off');
  const ext = child(xfrm, 'a:ext');
  const set = (node: Element, name: string, value: string | null): void =>
    batch.push({ op: 'setAttribute', part, node: node.id, name: { ns: '', local: name }, value });
  // ECMA-376 Part 1, 20.1.7.6 xfrm (`CT_Transform2D`): `a:off` and `a:ext` are optional, in
  // that order. A missing one is added in its place, so the shape still moves and resizes.
  const add = (index: number, xml: string): void =>
    void batch.insert({ op: 'insertNode', part, parent: xfrm.id, index, spec: specFromXml(xml) });
  if (off) {
    set(off, 'x', v.x);
    set(off, 'y', v.y);
  } else add(0, `<a:off x="${v.x}" y="${v.y}"/>`);
  if (ext) {
    set(ext, 'cx', v.cx);
    set(ext, 'cy', v.cy);
  } else add(off ? xfrm.children.indexOf(off) + 1 : 1, `<a:ext cx="${v.cx}" cy="${v.cy}"/>`);
  if (rotValue !== (Number(attr(xfrm, 'rot')) || 0))
    set(xfrm, 'rot', rotValue ? String(rotValue) : null);
  if (flips) {
    set(xfrm, 'flipH', flips.h ? '1' : null);
    set(xfrm, 'flipV', flips.v ? '1' : null);
  }
}

/** The element that holds a shape's transform; exported for group commands. */
export { transformHost };

/** Append a shape spec to the slide's shape tree; returns the new node id. */
function appendShape(batch: OpBatch, part: string, xml: string): string {
  const tree = spTreeOf(batch, part);
  const ext = tree.children.findIndex((c) => c.type === 'element' && c.local === 'extLst');
  return batch.insert({
    op: 'insertNode',
    part,
    parent: tree.id,
    index: ext < 0 ? tree.children.length : ext,
    spec: specFromXml(xml),
  });
}

/** A new text box centered on the slide, with the Insert → Text Box defaults. */
export function insertTextBox(
  batch: OpBatch,
  part: string,
  slideW: number,
  slideH: number
): string {
  const id = nextShapeId(batch, spTreeOf(batch, part));
  const cx = Math.round(slideW * 0.3);
  const cy = 369_332;
  return appendShape(
    batch,
    part,
    `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="TextBox ${id - 1}"/><p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>` +
      `<p:spPr><a:xfrm><a:off x="${int((slideW - cx) / 2)}" y="${int((slideH - cy) / 2)}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>` +
      '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>' +
      '<p:txBody><a:bodyPr wrap="square" rtlCol="0"><a:spAutoFit/></a:bodyPr><a:lstStyle/><a:p><a:endParaRPr lang="en-US" dirty="0"/></a:p></p:txBody></p:sp>'
  );
}

/** A new preset shape with the default style references (`p:style`). */
export function insertShape(
  batch: OpBatch,
  part: string,
  preset: string,
  slideW: number,
  slideH: number
): string {
  if (!isInsertablePreset(preset)) throw new CommandRefusal('errors.unsupported', preset);
  const id = nextShapeId(batch, spTreeOf(batch, part));
  const cx = Math.round(slideW * 0.2);
  const cy = preset === 'line' ? 0 : Math.round(slideH * 0.25);
  const style =
    '<p:style><a:lnRef idx="2"><a:schemeClr val="accent1"><a:shade val="50000"/></a:schemeClr></a:lnRef>' +
    '<a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef>' +
    '<a:fontRef idx="minor"><a:schemeClr val="lt1"/></a:fontRef></p:style>';
  if (preset === 'line')
    return appendShape(
      batch,
      part,
      `<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="${id}" name="Straight Connector ${id - 1}"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>` +
        `<p:spPr><a:xfrm><a:off x="${int((slideW - cx) / 2)}" y="${int(slideH / 2)}"/><a:ext cx="${cx}" cy="0"/></a:xfrm><a:prstGeom prst="line"><a:avLst/></a:prstGeom><a:ln w="19050"/></p:spPr>` +
        '<p:style><a:lnRef idx="1"><a:schemeClr val="accent1"/></a:lnRef><a:fillRef idx="0"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="tx1"/></a:fontRef></p:style></p:cxnSp>'
    );
  return appendShape(
    batch,
    part,
    `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Shape ${id - 1}"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
      `<p:spPr><a:xfrm><a:off x="${int((slideW - cx) / 2)}" y="${int((slideH - cy) / 2)}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>` +
      `<a:prstGeom prst="${preset}"><a:avLst/></a:prstGeom></p:spPr>${style}` +
      '<p:txBody><a:bodyPr rtlCol="0" anchor="ctr"/><a:lstStyle/><a:p><a:pPr algn="ctr"/><a:endParaRPr lang="en-US" dirty="0"/></a:p></p:txBody></p:sp>'
  );
}

const LINE_STYLE =
  '<p:style><a:lnRef idx="1"><a:schemeClr val="accent1"/></a:lnRef><a:fillRef idx="0"><a:schemeClr val="accent1"/></a:fillRef><a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="tx1"/></a:fontRef></p:style>';

/** A new line or connector, as Insert → Shapes → Lines creates it. */
export function insertLine(
  batch: OpBatch,
  part: string,
  kind: LineKind,
  slideW: number,
  slideH: number
): string {
  const presets: Record<LineKind, [string, string]> = {
    line: ['line', 'Straight Connector'],
    arrow: ['straightConnector1', 'Straight Arrow Connector'],
    elbow: ['bentConnector3', 'Elbow Connector'],
    curved: ['curvedConnector3', 'Curved Connector'],
  };
  if (!Object.prototype.hasOwnProperty.call(presets, kind))
    throw new CommandRefusal('errors.unsupported', 'line');
  const [prst, name] = presets[kind];
  const id = nextShapeId(batch, spTreeOf(batch, part));
  const cx = Math.round(slideW * 0.2);
  const cy = kind === 'elbow' || kind === 'curved' ? Math.round(slideH * 0.2) : 0;
  // 1.5 pt, so a new line is clearly visible; the color still comes from the style.
  const ln =
    kind === 'arrow' ? '<a:ln w="19050"><a:tailEnd type="triangle"/></a:ln>' : '<a:ln w="19050"/>';
  return appendShape(
    batch,
    part,
    `<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="${id}" name="${name} ${id - 1}"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>` +
      `<p:spPr><a:xfrm><a:off x="${int((slideW - cx) / 2)}" y="${int((slideH - cy) / 2)}"/><a:ext cx="${cx}" cy="${cy}"/></a:xfrm>` +
      `<a:prstGeom prst="${prst}"><a:avLst/></a:prstGeom>${ln}</p:spPr>${LINE_STYLE}</p:cxnSp>`
  );
}

const IMAGE_EXT: Record<string, string> = {
  'image/png': 'png',
  'image/jpeg': 'jpeg',
  'image/gif': 'gif',
  'image/svg+xml': 'svg',
};

/** The largest image a command adds as a media part. */
export const MAX_IMAGE_BYTES = 50 * 1024 * 1024;

/** Add image bytes as a media part related to `part`; returns the relationship id. */
export function addImagePart(
  batch: OpBatch,
  part: string,
  bytes: Uint8Array,
  mime: string
): string {
  // An own key only: a mime such as `constructor` names no extension.
  const ext = Object.hasOwn(IMAGE_EXT, mime) ? IMAGE_EXT[mime] : undefined;
  if (!ext || bytes.length === 0 || bytes.length > MAX_IMAGE_BYTES)
    throw new CommandRefusal('errors.unsupported', mime);
  const media = freePartName(batch, '/ppt/media/image', `.${ext}`);
  batch.push({ op: 'addPart', part: media, base64: toBase64(bytes) });
  ensureContentType(batch, ext, mime, media);
  return addRelationship(batch, part, REL.image, media);
}

/**
 * Insert a picture from image bytes (`w` × `h` px). It fits inside 60% of the slide, centered.
 * With `into`, it takes the place of that empty placeholder, as a placeholder with the same
 * index: a picture placeholder is filled and the picture cropped to its shape; a content
 * placeholder holds the picture whole, centered, never larger than its own size.
 */
export function insertImage(
  batch: OpBatch,
  part: string,
  bytes: Uint8Array,
  mime: string,
  w: number,
  h: number,
  slideW: number,
  slideH: number,
  into?: EmptyPlaceholder
): string {
  const rId = addImagePart(batch, part, bytes, mime);
  const pw = Math.max(1, w * 9525);
  const ph = Math.max(1, h * 9525);
  const fills = into !== undefined && into.type !== 'obj';
  const box = into ?? { x: 0, y: 0, cx: slideW * 0.6, cy: slideH * 0.6 };
  const scale = fills ? 1 : Math.min(box.cx / pw, box.cy / ph, 1);
  const cx = fills ? box.cx : Math.round(pw * scale);
  const cy = fills ? box.cy : Math.round(ph * scale);
  const x = into ? box.x + (box.cx - cx) / 2 : (slideW - cx) / 2;
  const y = into ? box.y + (box.cy - cy) / 2 : (slideH - cy) / 2;
  // A filled picture placeholder crops the picture to the box's shape, from the middle.
  let crop = '';
  if (fills) {
    const over = pw / ph / (box.cx / box.cy);
    const cut = (k: number) => int(((1 - 1 / k) / 2) * 100_000);
    if (over > 1) crop = `<a:srcRect l="${cut(over)}" r="${cut(over)}"/>`;
    else if (over < 1) crop = `<a:srcRect t="${cut(1 / over)}" b="${cut(1 / over)}"/>`;
  }
  const tree = spTreeOf(batch, part);
  const id = nextShapeId(batch, tree);
  const name = into ? escapeXmlAttribute(into.name) : `Picture ${id - 1}`;
  const xml =
    `<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="${name}"/><p:cNvPicPr><a:picLocks${into ? ' noGrp="1"' : ''} noChangeAspect="1"/></p:cNvPicPr><p:nvPr>${into ? into.ph : ''}</p:nvPr></p:nvPicPr>` +
    `<p:blipFill><a:blip r:embed="${rId}"/>${crop}<a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
    `<p:spPr><a:xfrm><a:off x="${int(x)}" y="${int(y)}"/><a:ext cx="${int(cx)}" cy="${int(cy)}"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>`;
  const at = into ? placeholderIndex(batch, part, into) : -1;
  if (at < 0) return appendShape(batch, part, xml);
  // The picture takes the placeholder's place in the z-order.
  const pic = batch.insert({
    op: 'insertNode',
    part,
    parent: tree.id,
    index: at,
    spec: specFromXml(xml),
  });
  // A filled picture placeholder keeps its shape (a round photo frame stays round).
  const geometry = fills ? into!.geometry : undefined;
  if (geometry) {
    const spPr = child(batch.element(part, pic), 'p:spPr')!;
    const rect = child(spPr, 'a:prstGeom')!;
    const index = spPr.children.indexOf(rect);
    batch.push({ op: 'removeNode', part, node: rect.id });
    batch.insert({ op: 'insertNode', part, parent: spPr.id, index, spec: specOf(geometry) });
  }
  removeFilledPlaceholder(batch, part, into!, String(id));
  return pic;
}

/**
 * Duplicate shapes, offset by 0.25 in (Ctrl+D). Returns the new ids. A connector
 * duplicated with the shapes it glues to stays glued to their copies.
 */
export function duplicateShapes(batch: OpBatch, part: string, ids: readonly string[]): string[] {
  const sources = ids.map((id) => batch.element(part, id));
  const { specs } = renumberShapes(
    sources.map((el) => specOf(el)),
    shapeIds(batch, spTreeOf(batch, part))
  );
  // Each copy goes right above its original. The places come from the slide before any insert,
  // and the inserts run from the last place in each parent to the first, so every place stays
  // valid and no insert reads the slide: Duplicate of thousands of shapes stays linear.
  const xml = xmlPartOf(batch.pkg, part);
  const index = xml ? nodeIndex(xml) : undefined;
  const places = sources.map((el, i) => {
    const loc = index?.get(el.id);
    if (!loc || loc.parent === null) throw new CommandRefusal('errors.unsupported', el.id);
    return { i, parent: loc.parent, at: loc.index + 1 };
  });
  places.sort((x, y) => (x.parent === y.parent ? y.at - x.at : x.parent < y.parent ? -1 : 1));
  const out: string[] = new Array<string>(sources.length);
  for (const { i, parent, at } of places)
    out[i] = batch.insert({
      op: 'insertNode',
      part,
      parent,
      index: at,
      spec: shiftedCopy(sources[i]!, specs[i]!),
    });
  return out;
}

/** The offset a duplicate moves by: 0.25 in down and to the right. */
const DUPLICATE_SHIFT = 228_600;

/** `spec`, the copy of `el`, with its `a:off` moved by {@link DUPLICATE_SHIFT}. */
function shiftedCopy(el: Element, spec: NodeSpec): NodeSpec {
  const host = transformHost(el);
  const xfrm = el.kind === 'graphicFrame' ? child(el, 'p:xfrm') : child(host, 'a:xfrm');
  const off = child(xfrm, 'a:off');
  if (!host || !xfrm || !off) return spec;
  // The spec's children follow the element's children, so the same positions lead to `a:off`.
  const chain = host === el ? [el, xfrm, off] : [el, host, xfrm, off];
  const path = chain.slice(1).map((node, k) => chain[k]!.children.indexOf(node));
  const moved = (name: string) => ({
    ns: '',
    local: name,
    value: coordinate((Number(attr(off, name)) || 0) + DUPLICATE_SHIFT, name),
  });
  const shift = (s: NodeSpec): NodeSpec => ({
    ...s,
    attrs: [
      ...(s.attrs ?? []).filter((a) => !(a.ns === '' && (a.local === 'x' || a.local === 'y'))),
      moved('x'),
      moved('y'),
    ],
  });
  const walk = (s: NodeSpec, depth: number): NodeSpec => {
    if (depth === path.length) return shift(s);
    const kids = (s.children ?? []).slice();
    const kid = kids[path[depth]!];
    if (!kid || 'text' in kid) return s;
    kids[path[depth]!] = walk(kid, depth + 1);
    return { ...s, children: kids };
  };
  return walk(spec, 0);
}

/** Delete shapes, and the animation entries that name them (`removeShapes`). */
export function deleteShapes(batch: OpBatch, part: string, ids: readonly string[]): void {
  removeShapes(batch, part, ids);
}

/**
 * Change the z-order of the selected shapes among their siblings (the slide's top level or a
 * group). With several shapes selected, the selection keeps its relative
 * order, whatever the click order was. Front and Back move it as a block to the top or the
 * bottom. Forward and Backward move each selected shape one step past the next shape that is
 * not selected, so adjacent selected shapes move together.
 */
export function orderShapes(
  batch: OpBatch,
  part: string,
  ids: readonly string[],
  move: OrderMove
): void {
  const byParent = new Map<string, Set<string>>();
  for (const id of ids) {
    const parent = parentOf(batch, part, id).id;
    const set = byParent.get(parent) ?? new Set<string>();
    set.add(id);
    byParent.set(parent, set);
  }
  for (const [parentId, selected] of byParent) {
    const shapes = stacked(batch.element(part, parentId));
    const known = new Set(shapes);
    if ([...selected].some((id) => !known.has(id)))
      throw new CommandRefusal('disabled.unsupportedElement');
    const next = reorder(shapes, selected, move);
    if (move === 'front' || move === 'back') {
      moveBlock(batch, part, parentId, shapes, next, selected.size, move === 'front');
      continue;
    }
    // Put each shape in its final place from the top down; each move lands it right under the
    // shape already placed above it, so the sibling list ends in `next` order. The lists are
    // kept here as the moves change them: a read of the tree for each shape would cost the
    // length of the list each time, and a selection can hold thousands of shapes.
    const kids = batch.element(part, parentId).children.map((c) => c.id);
    const current = [...shapes];
    for (let i = next.length - 1; i >= 0; i -= 1) {
      if (current[i] === next[i]) continue;
      const above = next[i + 1];
      const index =
        above === undefined ? kids.indexOf(current[current.length - 1]!) + 1 : kids.indexOf(above);
      const from = kids.indexOf(next[i]!);
      // `moveNode` removes first, then inserts: an index after the old place shifts down by one.
      const to = from < index ? index - 1 : index;
      batch.push({ op: 'moveNode', part, node: next[i]!, parent: parentId, index: to });
      kids.splice(from, 1);
      kids.splice(to, 0, next[i]!);
      current.splice(current.indexOf(next[i]!), 1);
      current.splice(i, 0, next[i]!);
    }
  }
}

/**
 * The number of selected shapes from which Back moves the other shapes up instead: each move
 * down shifts the shapes that wait to move, and the store then updates their places again.
 */
const BACK_MOVES = 32;

/**
 * Front or Back, in one pass: `next` holds the `picked` selected shapes as a block at the top
 * (`front`) or the bottom. The selected shapes already in their final place at that end stay.
 */
function moveBlock(
  batch: OpBatch,
  part: string,
  parentId: string,
  shapes: readonly string[],
  next: readonly string[],
  picked: number,
  front: boolean
): void {
  const kids = batch.element(part, parentId).children.map((c) => c.id);
  const n = shapes.length;
  const at = (i: number) => (front ? n - 1 - i : i);
  let kept = 0;
  while (kept < n && shapes[at(kept)] === next[at(kept)]) kept += 1;
  if (kept >= picked) return;
  /** Move `ids` (bottom to top) up, to right under shape `under` (or over the last shape). */
  const raise = (ids: readonly string[], under: string | undefined) => {
    const top = under === undefined ? kids.indexOf(shapes[n - 1]!) + 1 : kids.indexOf(under);
    // Top down: each shape comes from below the places that the earlier moves shifted, so the
    // store reads its place without an update. `moveNode` removes first, then inserts.
    for (let k = 0; k < ids.length; k += 1) {
      const id = ids[ids.length - 1 - k]!;
      batch.push({ op: 'moveNode', part, node: id, parent: parentId, index: top - 1 - k });
    }
  };
  if (front) return raise(next.slice(n - picked, n - kept), shapes[n - kept]);
  if (picked - kept <= BACK_MOVES) {
    // Right over the kept run, or in the place of the bottom shape, which is not selected.
    const to = kept > 0 ? kids.indexOf(shapes[kept - 1]!) + 1 : kids.indexOf(shapes[0]!);
    for (const id of next.slice(kept, picked).reverse())
      batch.push({ op: 'moveNode', part, node: id, parent: parentId, index: to });
    return;
  }
  // Many selected shapes: raise the other shapes under the top selected one above it instead.
  const top = shapes.indexOf(next[picked - 1]!);
  const chosen = new Set(next.slice(0, picked));
  raise(
    shapes.slice(kept, top).filter((id) => !chosen.has(id)),
    shapes[top + 1]
  );
}

/** The ids of a tree's shapes, bottom to top; its `p:extLst` stays last and never moves. */
const stacked = (tree: Element): string[] =>
  shapeChildren(tree)
    .filter((e) => !(e.ns === NS.p && e.local === 'extLst'))
    .map((e) => e.id);

/** The sibling order after a z-order move of `selected` (ids bottom to top). */
export function reorder(
  shapes: readonly string[],
  selected: ReadonlySet<string>,
  move: OrderMove
): string[] {
  const picked = shapes.filter((id) => selected.has(id));
  const rest = shapes.filter((id) => !selected.has(id));
  if (move === 'front') return [...rest, ...picked];
  if (move === 'back') return [...picked, ...rest];
  const out = [...shapes];
  if (move === 'forward') {
    for (let i = out.length - 2; i >= 0; i -= 1)
      if (selected.has(out[i]) && !selected.has(out[i + 1]))
        [out[i], out[i + 1]] = [out[i + 1], out[i]];
  } else {
    for (let i = 1; i < out.length; i += 1)
      if (selected.has(out[i]) && !selected.has(out[i - 1]))
        [out[i], out[i - 1]] = [out[i - 1], out[i]];
  }
  return out;
}

/** The box around all `boxes`, by loop: a selection can hold every shape of a file. */
function unionOf(boxes: readonly { x: number; y: number; cx: number; cy: number }[]): {
  x: number;
  y: number;
  r: number;
  b: number;
} {
  let x = Infinity;
  let y = Infinity;
  let r = -Infinity;
  let b = -Infinity;
  for (const box of boxes) {
    x = Math.min(x, box.x);
    y = Math.min(y, box.y);
    r = Math.max(r, box.x + box.cx);
    b = Math.max(b, box.y + box.cy);
  }
  return { x, y, r, b };
}

/** Align boxes: one element aligns to the slide; several align to their common bounds. */
export function alignBoxes(
  boxes: readonly { id: string; x: number; y: number; cx: number; cy: number }[],
  align: ObjectAlign,
  slideW: number,
  slideH: number
): { id: string; x: number; y: number }[] {
  const bounds = boxes.length === 1 ? { x: 0, y: 0, r: slideW, b: slideH } : unionOf(boxes);
  return boxes.map((b) => {
    switch (align) {
      case 'left':
        return { id: b.id, x: bounds.x, y: b.y };
      case 'center':
        return { id: b.id, x: (bounds.x + bounds.r - b.cx) / 2, y: b.y };
      case 'right':
        return { id: b.id, x: bounds.r - b.cx, y: b.y };
      case 'top':
        return { id: b.id, x: b.x, y: bounds.y };
      case 'middle':
        return { id: b.id, x: b.x, y: (bounds.y + bounds.b - b.cy) / 2 };
      case 'bottom':
        return { id: b.id, x: b.x, y: bounds.b - b.cy };
    }
  });
}
