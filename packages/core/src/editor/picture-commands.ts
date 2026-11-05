// Picture commands: crop (`a:srcRect`), crop to shape, Change
// Picture, Reset Picture, corrections, color, and transparency (blip effects), picture styles, and
// alt text. Each compiles to ops on an OpBatch against one slide part, as shape-commands.ts does.
// Images come only from the bytes the command carries and become package parts; nothing is
// fetched.

import {
  attr,
  binaryPartOf,
  child,
  elements,
  imageSize,
  imageSizeEmu,
  intAttr,
  isValidXmlText,
  NS,
  relTarget,
  type Element,
} from '../store/index.ts';
import type {
  EditorCommand,
  ImageAdjustments,
  ImageCrop,
  ImageCropPercent,
  ImageProperties,
  ImageRecolor,
} from '../contracts/index.ts';
import type { LayoutItem } from '../layout/index.ts';
import type { CommandContext, CommandEffect } from './commands.ts';
import { selectedIds } from './commands.ts';
import { CommandRefusal, type OpBatch } from '../automation/commands/op-batch.ts';
import { boxOf, findItem, toFile, type Box } from '../automation/commands/placement.ts';
import {
  BLIP_EFFECTS,
  IMG_PROPS_EXT,
  PICTURE_STYLES,
  RECOLOR_EFFECTS,
  RECOLOR_PRESETS,
  SVG_BLIP_EXT,
  recolorOf,
  recolorXml,
} from './picture-presets.ts';
import { addImagePart, setTransform, specFromXml } from '../automation/commands/shape-commands.ts';
import { refitGroup } from '../automation/commands/group-commands.ts';
import { parentOf } from '../automation/commands/package-edits.ts';
import { isInsertablePreset, NO_INSIDE_PRESET } from './shape-gallery.ts';

/** The commands this module compiles. */
export type PictureCommand = Extract<
  EditorCommand,
  {
    type:
      | 'cropImage'
      | 'cropToShape'
      | 'replaceImage'
      | 'resetImage'
      | 'setImageAdjustments'
      | 'setImageStyle'
      | 'setImageProperties'
      | 'setAltText';
  }
>;

/** A crop as fractions of the image: `l`, `t`, `r`, `b` (`a:srcRect`). */
export interface CropRect {
  readonly l: number;
  readonly t: number;
  readonly r: number;
  readonly b: number;
}

/** The longest alt text the editor writes. */
export const MAX_ALT_TEXT = 4_000;
/**
 * The most space a crop adds outside the image, as a multiple of the image size: room for Fit
 * at aspect ratios up to 2000:1, and well inside `ST_Percentage`.
 */
export const MAX_PAD = 1_000;
/** The smallest visible part of the image a crop keeps. */
const MIN_VISIBLE = 0.001;

const ADEC = 'http://schemas.microsoft.com/office/drawing/2017/decorative';
/** The `a:ext/@uri` of "Mark as decorative" (`adec:decorative`). */
const DECORATIVE_EXT = '{C183D7F6-B498-43B3-948B-1728B52AA6E4}';

/** `p:spPr` children in schema order (ECMA-376 Part 1, 20.1.2.2.35 CT_ShapeProperties). */
const SPPR_ORDER = [
  'xfrm',
  'custGeom',
  'prstGeom',
  'noFill',
  'solidFill',
  'gradFill',
  'blipFill',
  'pattFill',
  'grpFill',
  'ln',
  'effectLst',
  'effectDag',
  'scene3d',
  'sp3d',
  'extLst',
];

/** A percentage attribute (ST_Percentage, 1000ths of a percent) as a fraction, read as the cascade reads it. */
const pct = (el: Element | undefined, name: string): number => (intAttr(el, name) ?? 0) / 100_000;

/** Refuse a command payload field that is not a plain object. */
function objectOf(value: unknown, name: string): Record<string, unknown> {
  if (typeof value !== 'object' || value === null || Array.isArray(value))
    throw new CommandRefusal('errors.unsupported', name);
  return value as Record<string, unknown>;
}

/** Refuse a field the command does not know, so nothing is dropped without a reason. */
function knownKeys(value: object, keys: readonly string[]): void {
  for (const key of Object.keys(value))
    if (!keys.includes(key)) throw new CommandRefusal('errors.unsupported', key);
}

const isA = (node: { type: string }, local: string): node is Element =>
  node.type === 'element' && (node as Element).ns === NS.a && (node as Element).local === local;

/** A crop in percent (the public shape) as fractions. */
const fromPercent = (c: ImageCropPercent): CropRect => ({
  l: c.left / 100,
  t: c.top / 100,
  r: c.right / 100,
  b: c.bottom / 100,
});

/** A crop as fractions, in percent (the public shape). */
export const toPercent = (c: CropRect): ImageCropPercent => ({
  left: c.l * 100,
  top: c.t * 100,
  right: c.r * 100,
  bottom: c.b * 100,
});

/** The crop a `p:blipFill` holds (`a:srcRect`), as fractions. */
export function cropOf(blipFill: Element | undefined): CropRect {
  const sr = child(blipFill, 'a:srcRect');
  return { l: pct(sr, 'l'), t: pct(sr, 't'), r: pct(sr, 'r'), b: pct(sr, 'b') };
}

/** The media elements of `p:nvPr`: the `a:EG_Media` group (ECMA-376 Part 1, `pml.xsd` CT_ApplicationNonVisualDrawingProps). */
const MEDIA = new Set(['videoFile', 'audioFile', 'audioCd', 'wavAudioFile', 'quickTimeFile']);

/** Whether a picture is a video or audio poster: its image cannot be changed. */
function isMedia(pic: Element): boolean {
  const nvPr = child(child(pic, 'p:nvPicPr'), 'p:nvPr');
  return elements(nvPr).some((m) => m.ns === NS.a && MEDIA.has(m.local));
}

function pictures(batch: OpBatch, part: string, ctx: CommandContext): Element[] {
  const ids = selectedIds(ctx.selection);
  if (ids.length === 0) throw new CommandRefusal('disabled.noSelection');
  const els = ids.map((id) => batch.element(part, id));
  if (els.some((el) => el.kind !== 'pic')) throw new CommandRefusal('disabled.noPicture');
  return els;
}

function onePicture(batch: OpBatch, part: string, ctx: CommandContext): Element {
  const els = pictures(batch, part, ctx);
  if (els.length !== 1) throw new CommandRefusal('disabled.onePicture');
  return els[0];
}

function partOf(el: Element, name: string): Element {
  const found = child(el, name);
  if (!found) throw new CommandRefusal('errors.malformed');
  return found;
}

/** The laid-out item of a picture, for its box in parent space. */
function itemOf(ctx: CommandContext, id: string): LayoutItem {
  const item = findItem(ctx.session.layout(ctx.slide)?.items ?? [], id);
  if (!item) throw new CommandRefusal('errors.unsupported', 'layout');
  return item;
}

const remove = (batch: OpBatch, part: string, el: Element): void =>
  batch.push({ op: 'removeNode', part, node: el.id });

/** Set an unprefixed attribute; null removes it. */
const setAttr = (
  batch: OpBatch,
  part: string,
  el: Element,
  name: string,
  value: string | null
): void => {
  if ((attr(el, name) ?? null) === value) return;
  batch.push({ op: 'setAttribute', part, node: el.id, name: { ns: '', local: name }, value });
};

/**
 * Replace the `p:spPr` children named in `names` with `xml` (nothing for null), at its schema
 * position.
 */
function setSpPrChild(
  batch: OpBatch,
  part: string,
  spPr: Element,
  names: readonly string[],
  xml: string | null
): void {
  for (const c of [...batch.element(part, spPr.id).children])
    if (c.type === 'element' && c.ns === NS.a && names.includes(c.local)) remove(batch, part, c);
  if (xml === null) return;
  const spec = specFromXml(xml);
  const rank = SPPR_ORDER.indexOf(spec.local);
  const parent = batch.element(part, spPr.id);
  let index = 0;
  parent.children.forEach((c, i) => {
    if (c.type === 'element' && c.ns === NS.a && SPPR_ORDER.indexOf(c.local) < rank) index = i + 1;
  });
  batch.insert({ op: 'insertNode', part, parent: spPr.id, index, spec });
}

/** Write `a:srcRect` after the blip; a crop of nothing removes it. */
function writeCrop(batch: OpBatch, part: string, blipFill: Element, crop: CropRect): void {
  const sr = child(blipFill, 'a:srcRect');
  const vals = (['l', 't', 'r', 'b'] as const).map((k) => {
    const v = Math.round(crop[k] * 100_000);
    return [k, v === 0 ? null : String(v)] as const;
  });
  if (vals.every(([, v]) => v === null)) {
    if (sr) remove(batch, part, sr);
    return;
  }
  if (sr) {
    for (const [k, v] of vals) setAttr(batch, part, sr, k, v);
    return;
  }
  const blipIndex = blipFill.children.findIndex((c) => isA(c, 'blip'));
  const xml = `<a:srcRect${vals.map(([k, v]) => (v === null ? '' : ` ${k}="${v}"`)).join('')}/>`;
  batch.insert({
    op: 'insertNode',
    part,
    parent: blipFill.id,
    index: blipIndex + 1,
    spec: specFromXml(xml),
  });
}

/**
 * The crop as the editor reads it: an axis that leaves nothing visible (two concurrent crops
 * merged edge by edge) counts as uncropped, so the frame is the image on that axis and a new
 * crop repairs the picture. Paint reads it the same way (`output/paint-fill.ts`, `pictureImage`).
 */
export function usableCrop(c: CropRect): CropRect {
  const w = 1 - c.l - c.r >= MIN_VISIBLE;
  const h = 1 - c.t - c.b >= MIN_VISIBLE;
  return w && h ? c : { l: w ? c.l : 0, r: w ? c.r : 0, t: h ? c.t : 0, b: h ? c.b : 0 };
}

/** Check a crop: finite, the visible part positive, and the space it adds bounded. */
function checkCrop(crop: CropRect): CropRect {
  const vals = [crop.l, crop.t, crop.r, crop.b];
  if (vals.some((v) => !Number.isFinite(v) || v < -MAX_PAD || v >= 1))
    throw new CommandRefusal('errors.unsupported', 'crop');
  if (1 - crop.l - crop.r < MIN_VISIBLE || 1 - crop.t - crop.b < MIN_VISIBLE)
    throw new CommandRefusal('errors.unsupported', 'crop');
  return crop;
}

/**
 * The frame after a crop change from `from` to `to`, with the image fixed on the slide, as a
 * crop handle moves it. The math runs in the frame's unrotated space; the frame turns
 * about its own center, so the new center is turned back with the rotation and the flips.
 */
export function croppedBox(
  box: Box,
  flips: { readonly h: boolean; readonly v: boolean },
  current: CropRect,
  to: CropRect
): Box {
  const from = usableCrop(current);
  const visW = 1 - from.l - from.r;
  const visH = 1 - from.t - from.b;
  const imgW = box.cx / visW;
  const imgH = box.cy / visH;
  const imgX = box.x - from.l * imgW;
  const imgY = box.y - from.t * imgH;
  const cx = imgW * (1 - to.l - to.r);
  const cy = imgH * (1 - to.t - to.b);
  const center = { x: box.x + box.cx / 2, y: box.y + box.cy / 2 };
  const dx = (imgX + to.l * imgW + cx / 2 - center.x) * (flips.h ? -1 : 1);
  const dy = (imgY + to.t * imgH + cy / 2 - center.y) * (flips.v ? -1 : 1);
  const r = (box.rot * Math.PI) / 180;
  const x = center.x + dx * Math.cos(r) - dy * Math.sin(r);
  const y = center.y + dx * Math.sin(r) + dy * Math.cos(r);
  return { x: x - cx / 2, y: y - cy / 2, cx, cy, rot: box.rot };
}

/**
 * The crop that makes an image of natural aspect `image` (width / height) cover (`fill`) or fit
 * in (`fit`) a frame of aspect `frame`, centered, as Crop > Fill and Crop > Fit do.
 */
export function fillOrFitCrop(kind: 'fill' | 'fit', image: number, frame: number): CropRect {
  // `fill` crops the longer side of the image; `fit` adds space to the shorter side of the frame.
  const wider = image > frame;
  const side =
    kind === 'fill'
      ? (1 - (wider ? frame / image : image / frame)) / 2
      : -((wider ? image / frame : frame / image) - 1) / 2;
  return wider === (kind === 'fill')
    ? { l: side, t: 0, r: side, b: 0 }
    : { l: 0, t: side, r: 0, b: side };
}

/** The crop that gives the visible image an aspect ratio, as large as the image allows. */
function aspectCrop(box: Box, from: CropRect, ratio: number): CropRect {
  const imgW = box.cx / (1 - from.l - from.r);
  const imgH = box.cy / (1 - from.t - from.b);
  const keepW = imgW / imgH > ratio ? (imgH * ratio) / imgW : 1;
  const keepH = imgW / imgH > ratio ? 1 : imgW / ratio / imgH;
  return { l: (1 - keepW) / 2, t: (1 - keepH) / 2, r: (1 - keepW) / 2, b: (1 - keepH) / 2 };
}

/** Write a new frame box (parent space) for a picture. */
function writeBox(batch: OpBatch, part: string, item: LayoutItem, box: Box): void {
  const f = toFile(item, box);
  if (!(f.cx >= 1 && f.cy >= 1)) throw new CommandRefusal('errors.coordinateRange');
  setTransform(batch, part, item.element.nodeId, f.x, f.y, f.cx, f.cy, f.rot);
}

const CROP_KEYS: Readonly<Record<ImageCrop['kind'], readonly string[]>> = {
  rect: ['kind', 'left', 'top', 'right', 'bottom', 'keepFrame'],
  fill: ['kind'],
  fit: ['kind'],
  aspect: ['kind', 'width', 'height'],
};

/** Crop the selected picture. True when the frame changed (a group around it then refits). */
function crop(batch: OpBatch, part: string, ctx: CommandContext, value: ImageCrop): boolean {
  // A crop from a host is plain JSON: an unknown kind or field is refused, never guessed.
  const raw = objectOf(value, 'crop');
  const kind = raw.kind as ImageCrop['kind'];
  if (typeof kind !== 'string' || !Object.hasOwn(CROP_KEYS, kind))
    throw new CommandRefusal('errors.unsupported', 'crop');
  knownKeys(raw, CROP_KEYS[kind]);
  for (const key of ['left', 'top', 'right', 'bottom', 'width', 'height'])
    if (key in raw && typeof raw[key] !== 'number')
      throw new CommandRefusal('errors.unsupported', key);
  if (raw.keepFrame !== undefined && typeof raw.keepFrame !== 'boolean')
    throw new CommandRefusal('errors.unsupported', 'keepFrame');
  const request = value;
  const pic = onePicture(batch, part, ctx);
  refuseLockedCrop(pic);
  const blipFill = partOf(pic, 'p:blipFill');
  const item = itemOf(ctx, pic.id);
  const box = boxOf(item);
  const from = usableCrop(cropOf(blipFill));
  const flips = { h: item.element.xfrm.flipH, v: item.element.xfrm.flipV };
  if (request.kind === 'rect') {
    const to = checkCrop(fromPercent(request));
    if (!request.keepFrame) {
      const next = croppedBox(box, flips, from, to);
      refuseLockedResize(pic, box, next);
      writeBox(batch, part, item, next);
    }
    writeCrop(batch, part, blipFill, to);
    return !request.keepFrame;
  }
  if (request.kind === 'aspect') {
    const ratio = request.width / request.height;
    if (!(ratio > 0.01 && ratio < 100)) throw new CommandRefusal('errors.unsupported', 'aspect');
    // A ratio far from the image's leaves almost nothing visible: refused as any such crop.
    const to = checkCrop(aspectCrop(box, from, ratio));
    const next = croppedBox(box, flips, from, to);
    refuseLockedResize(pic, box, next);
    writeBox(batch, part, item, next);
    writeCrop(batch, part, blipFill, to);
    return true;
  }
  const rId = attr(child(blipFill, 'a:blip'), 'r:embed');
  const natural = rId ? ctx.session.imageSize(part, rId) : undefined;
  if (!natural || !(natural.cx > 0 && natural.cy > 0) || !(box.cx > 0 && box.cy > 0))
    throw new CommandRefusal('disabled.imageSizeUnknown');
  writeCrop(
    batch,
    part,
    blipFill,
    checkCrop(fillOrFitCrop(request.kind, natural.cx / natural.cy, box.cx / box.cy))
  );
  return false;
}

const PROPERTY_KEYS = ['type', 'widthEmu', 'heightEmu', 'rotationDegrees', 'crop'] as const;
const CROP_EDGES = ['left', 'top', 'right', 'bottom'] as const;

/**
 * Format Picture: crop first (the image stays), then size and rotation, as one box write. Width
 * and height are the size on the slide, as the Size fields show it (`SelectedImageState`); the
 * top-left corner of the unrotated box stays. No field changes nothing.
 */
function setProperties(
  batch: OpBatch,
  part: string,
  ctx: CommandContext,
  p: ImageProperties
): void {
  knownKeys(p, PROPERTY_KEYS);
  if (p.crop !== undefined) {
    const edges = objectOf(p.crop, 'crop');
    knownKeys(edges, CROP_EDGES);
    for (const edge of CROP_EDGES)
      if (typeof edges[edge] !== 'number') throw new CommandRefusal('errors.unsupported', edge);
  }
  const pic = onePicture(batch, part, ctx);
  for (const v of [p.widthEmu, p.heightEmu])
    if (v !== undefined && !(Number.isFinite(v) && v > 0))
      throw new CommandRefusal('errors.coordinateRange');
  if (
    p.rotationDegrees !== undefined &&
    !(Number.isFinite(p.rotationDegrees) && Math.abs(p.rotationDegrees) <= 1e6)
  )
    throw new CommandRefusal('errors.unsupported', 'rotation');
  // As DOCX Editor's `setImageProperties`: a command with no field is refused.
  if (
    !p.crop &&
    p.widthEmu === undefined &&
    p.heightEmu === undefined &&
    p.rotationDegrees === undefined
  )
    throw new CommandRefusal(
      'errors.unsupported',
      'setImageProperties needs widthEmu, heightEmu, rotationDegrees, or crop'
    );
  const blipFill = partOf(pic, 'p:blipFill');
  const item = itemOf(ctx, pic.id);
  const original = boxOf(item);
  let box = original;
  // A locked picture (`a:picLocks`) refuses a field that changes what the lock holds; a value
  // equal to the current one changes nothing and passes. A crop that changes the frame's size
  // counts as a resize.
  const resized =
    (p.widthEmu !== undefined && p.widthEmu !== box.cx) ||
    (p.heightEmu !== undefined && p.heightEmu !== box.cy);
  if (resized && pictureLock(pic, 'noResize')) throw new CommandRefusal('disabled.sizeLocked');
  if (p.rotationDegrees !== undefined && p.rotationDegrees !== box.rot && pictureLock(pic, 'noRot'))
    throw new CommandRefusal('disabled.rotationLocked');
  if (p.crop) {
    refuseLockedCrop(pic);
    const to = checkCrop(fromPercent(p.crop));
    const flips = { h: item.element.xfrm.flipH, v: item.element.xfrm.flipV };
    box = croppedBox(box, flips, cropOf(blipFill), to);
    writeCrop(batch, part, blipFill, to);
  }
  const next = {
    ...box,
    cx: p.widthEmu ?? box.cx,
    cy: p.heightEmu ?? box.cy,
    rot: p.rotationDegrees ?? box.rot,
  };
  refuseLockedResize(pic, original, next);
  writeBox(batch, part, item, next);
}

/** Whether a picture can be cropped to `preset`: a preset with an inside (not a line). */
export function isCropShape(preset: string): boolean {
  return isInsertablePreset(preset) && !NO_INSIDE_PRESET.test(preset);
}

/** A picture lock flag of `a:picLocks` (ECMA-376 Part 1, 20.1.2.2.31 CT_PictureLocking). */
export function pictureLock(
  pic: Element,
  name: 'noCrop' | 'noChangeAspect' | 'noResize' | 'noRot' | 'noChangeShapeType'
): boolean {
  const locks = child(child(child(pic, 'p:nvPicPr'), 'p:cNvPicPr'), 'a:picLocks');
  const v = attr(locks, name);
  return v === '1' || v === 'true';
}

/** Whether a picture's `a:srcRect` crops or pads any edge. */
function isCropped(pic: Element): boolean {
  return Object.values(cropOf(child(pic, 'p:blipFill'))).some((v) => v !== 0);
}

function refuseLockedCrop(pic: Element): void {
  if (pictureLock(pic, 'noCrop')) throw new CommandRefusal('disabled.cropLocked');
  refuseCropFillMode(pic);
}

/**
 * Whether a picture's blip fill tiles (`a:tile`) or stretches into an inset `a:fillRect`: crop
 * mode and the crop math read the image as stretched over the whole frame, so they refuse it.
 * ECMA-376 Part 1, 20.1.8.14 CT_BlipFillProperties.
 */
export function hasCropFillMode(pic: Element): boolean {
  const blipFill = child(pic, 'p:blipFill');
  if (child(blipFill, 'a:tile')) return true;
  const rect = child(child(blipFill, 'a:stretch'), 'a:fillRect');
  return ['l', 't', 'r', 'b'].some((edge) => (Number(attr(rect, edge)) || 0) !== 0);
}

function refuseCropFillMode(pic: Element): void {
  if (hasCropFillMode(pic)) throw new CommandRefusal('disabled.cropFillMode');
}

/** A crop that changes the frame's size: refused for a picture locked against resizing. */
function refuseLockedResize(pic: Element, from: Box, to: Box): void {
  if ((from.cx !== to.cx || from.cy !== to.cy) && pictureLock(pic, 'noResize'))
    throw new CommandRefusal('disabled.sizeLocked');
}

function cropToShape(batch: OpBatch, part: string, ctx: CommandContext, preset: string): void {
  if (typeof preset !== 'string') throw new CommandRefusal('errors.unsupported', 'preset');
  if (!isCropShape(preset)) throw new CommandRefusal('errors.unsupported', preset);
  for (const pic of pictures(batch, part, ctx)) {
    const spPr = partOf(pic, 'p:spPr');
    if (attr(child(spPr, 'a:prstGeom'), 'prst') === preset) continue;
    if (pictureLock(pic, 'noChangeShapeType')) throw new CommandRefusal('disabled.shapeLocked');
    setSpPrChild(
      batch,
      part,
      spPr,
      ['prstGeom', 'custGeom'],
      `<a:prstGeom prst="${preset}"><a:avLst/></a:prstGeom>`
    );
  }
}

/** Remove the extensions of a blip that describe its old image (SVG, the picture layer). */
function dropImageExtensions(batch: OpBatch, part: string, blip: Element): void {
  const extLst = child(blip, 'a:extLst');
  for (const ext of elements(extLst)) {
    const uri = attr(ext, 'uri');
    if (uri === SVG_BLIP_EXT || uri === IMG_PROPS_EXT) remove(batch, part, ext);
  }
  if (extLst && elements(batch.element(part, extLst.id)).length === 0) remove(batch, part, extLst);
}

function replaceImage(
  batch: OpBatch,
  part: string,
  ctx: CommandContext,
  command: Extract<EditorCommand, { type: 'replaceImage' }>
): void {
  const pic = onePicture(batch, part, ctx);
  if (isMedia(pic)) throw new CommandRefusal('disabled.unsupportedElement');
  if (!(command.bytes instanceof Uint8Array))
    throw new CommandRefusal('errors.unsupported', 'bytes');
  // Both sides or none: one side alone would be ignored without a reason.
  for (const key of ['width', 'height'] as const)
    if (command[key] !== undefined && typeof command[key] !== 'number')
      throw new CommandRefusal('errors.unsupported', key);
  if ((command.width === undefined) !== (command.height === undefined))
    throw new CommandRefusal(
      'errors.unsupported',
      command.width === undefined ? 'width' : 'height'
    );
  // A deck keeps an SVG beside a raster blip (`asvg:svgBlip`). The editor cannot draw that
  // raster copy, so Change Picture takes raster images only.
  if (command.mime === 'image/svg+xml') throw new CommandRefusal('errors.unsupported', 'SVG');
  // The host's size when it gives one; else the stored pixels and resolution the header states.
  // A JPEG's EXIF orientation does not count.
  const header =
    command.width === undefined || command.height === undefined ? imageSize(command.bytes) : null;
  const natural = header
    ? imageSizeEmu(header)
    : { cx: command.width ?? NaN, cy: command.height ?? NaN };
  const aspect = natural.cx / natural.cy;
  if (!(natural.cx > 0 && natural.cy > 0 && Number.isFinite(aspect)))
    throw new CommandRefusal('errors.unsupported', 'image size');
  const blipFill = partOf(pic, 'p:blipFill');
  const item = itemOf(ctx, pic.id);
  // The new image fits in the old frame, about the same center, with its own aspect.
  const box = boxOf(item);
  const cx = aspect > box.cx / box.cy ? box.cx : box.cy * aspect;
  const cy = aspect > box.cx / box.cy ? box.cx / aspect : box.cy;
  // An image with the frame's aspect keeps the frame exactly: a size within 1 EMU of the old one
  // is only rounding (file sizes are whole EMU), not a resize.
  const same = Math.abs(cx - box.cx) <= 1 && Math.abs(cy - box.cy) <= 1;
  const next = same
    ? box
    : { x: box.x + (box.cx - cx) / 2, y: box.y + (box.cy - cy) / 2, cx, cy, rot: box.rot };
  // The locks hold here as in the crop commands: no new frame size, no removed crop.
  refuseLockedResize(pic, box, next);
  if (pictureLock(pic, 'noCrop') && isCropped(pic)) throw new CommandRefusal('disabled.cropLocked');
  const rId = addImagePart(batch, part, command.bytes, command.mime);
  const blip = child(blipFill, 'a:blip');
  if (blip) {
    batch.push({
      op: 'setAttribute',
      part,
      node: blip.id,
      name: { ns: NS.r, local: 'embed' },
      value: rId,
    });
    // A linked picture ("Insert and Link") would show the old file again on refresh.
    if (attr(blip, 'r:link') !== undefined)
      batch.push({
        op: 'setAttribute',
        part,
        node: blip.id,
        name: { ns: NS.r, local: 'link' },
        value: null,
      });
    dropImageExtensions(batch, part, blip);
  } else
    batch.insert({
      op: 'insertNode',
      part,
      parent: blipFill.id,
      index: 0,
      spec: specFromXml(`<a:blip r:embed="${rId}"/>`),
    });
  writeCrop(batch, part, batch.element(part, blipFill.id), { l: 0, t: 0, r: 0, b: 0 });
  // The old image and its relationship stay: in a shared document another participant may
  // still use them (a concurrent duplicate), and nothing may point at a removed part.
  writeBox(batch, part, item, next);
}

/**
 * Whether a blip holds 2010 picture effects (`a14:imgProps`: corrections, color, or an
 * artistic effect). Its `a14:imgLayer` keeps the original image while `r:embed` holds the
 * rendered one, so the editor cannot reset or re-correct the picture without losing the
 * original; it refuses instead.
 */
function hasOfficeEffects(blip: Element | undefined): boolean {
  return elements(child(blip, 'a:extLst')).some((ext) => attr(ext, 'uri') === IMG_PROPS_EXT);
}

function refuseImgPropsEffects(pics: readonly Element[]): void {
  if (pics.some((pic) => hasOfficeEffects(child(child(pic, 'p:blipFill'), 'a:blip'))))
    throw new CommandRefusal('disabled.pictureEffects');
}

function resetImage(batch: OpBatch, part: string, ctx: CommandContext, size: boolean): void {
  refuseImgPropsEffects(pictures(batch, part, ctx));
  if (size && pictures(batch, part, ctx).some((pic) => pictureLock(pic, 'noResize')))
    throw new CommandRefusal('disabled.sizeLocked');
  // Reset Picture & Size removes the crop: a picture locked against cropping keeps its crop.
  if (size && pictures(batch, part, ctx).some((p) => pictureLock(p, 'noCrop') && isCropped(p)))
    throw new CommandRefusal('disabled.cropLocked');
  for (const pic of pictures(batch, part, ctx)) {
    const blipFill = partOf(pic, 'p:blipFill');
    const blip = child(blipFill, 'a:blip');
    if (blip) {
      for (const e of elements(blip))
        if (e.ns === NS.a && BLIP_EFFECTS.has(e.local)) remove(batch, part, e);
    }
    const spPr = partOf(pic, 'p:spPr');
    for (const name of ['ln', 'effectLst', 'effectDag', 'scene3d', 'sp3d'])
      setSpPrChild(batch, part, spPr, [name], null);
    // Only the picture's own shape resets; a shape inherited from its placeholder stays.
    const own = child(spPr, 'a:prstGeom');
    if ((own && attr(own, 'prst') !== 'rect') || child(spPr, 'a:custGeom')) {
      if (pictureLock(pic, 'noChangeShapeType')) throw new CommandRefusal('disabled.shapeLocked');
      setSpPrChild(
        batch,
        part,
        spPr,
        ['prstGeom', 'custGeom'],
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
      );
    }
    const style = child(pic, 'p:style');
    if (style) remove(batch, part, style);
    if (!size) continue;
    writeCrop(batch, part, blipFill, { l: 0, t: 0, r: 0, b: 0 });
    const rId = attr(blip, 'r:embed');
    const natural = rId ? resetSize(ctx, part, rId) : undefined;
    if (!natural) throw new CommandRefusal('disabled.imageSizeUnknown');
    // The top-left corner stays where it is; the size is in slide units, as the box is.
    const item = itemOf(ctx, pic.id);
    const box = boxOf(item);
    writeBox(batch, part, item, { ...box, cx: natural.cx, cy: natural.cy });
  }
}

/**
 * The size Reset Picture & Size gives: the image's pixels at its resolution, 96 DPI when it
 * states none (`imageSizeEmu`). The stored pixels count, not a JPEG's EXIF orientation.
 * The session's `imageSize` uses 72 DPI, the tile rule.
 */
function resetSize(
  ctx: CommandContext,
  part: string,
  rId: string
): { cx: number; cy: number } | undefined {
  const pkg = ctx.session.store.package;
  const target = relTarget(pkg, part, rId);
  const bytes = target ? binaryPartOf(pkg, target) : undefined;
  const size = bytes ? imageSize(bytes) : null;
  return size ? imageSizeEmu(size) : undefined;
}

/** The corrections and transparency a blip holds. */
export function adjustmentsOf(blip: Element | undefined): {
  brightness: number;
  contrast: number;
  transparency: number;
  recolor: ImageRecolor | 'other';
} {
  const lum = child(blip, 'a:lum');
  const alpha = child(blip, 'a:alphaModFix');
  return {
    brightness: pct(lum, 'bright'),
    contrast: pct(lum, 'contrast'),
    transparency: alpha ? 1 - (intAttr(alpha, 'amt') === undefined ? 1 : pct(alpha, 'amt')) : 0,
    recolor: recolorOf(blip),
  };
}

function inRange(v: number | undefined, lo: number, hi: number, name: string): void {
  if (v !== undefined && !(Number.isFinite(v) && v >= lo && v <= hi))
    throw new CommandRefusal('errors.unsupported', name);
}

function setAdjustments(
  batch: OpBatch,
  part: string,
  ctx: CommandContext,
  adjust: ImageAdjustments
): void {
  const fields = objectOf(adjust, 'adjustments');
  knownKeys(fields, ['brightness', 'contrast', 'transparency', 'recolor']);
  // As `setImageProperties`: a command with no field is refused.
  if (Object.values(fields).every((v) => v === undefined))
    throw new CommandRefusal('errors.unsupported', 'adjustments');
  inRange(adjust.brightness, -1, 1, 'brightness');
  inRange(adjust.contrast, -1, 1, 'contrast');
  inRange(adjust.transparency, 0, 1, 'transparency');
  if (adjust.recolor !== undefined && !RECOLOR_PRESETS.includes(adjust.recolor))
    throw new CommandRefusal('errors.unsupported', 'recolor');
  // Only corrections conflict with the `a14:imgProps` layer. Recolor and transparency are blip
  // effects on the image that `r:embed` holds.
  if (adjust.brightness !== undefined || adjust.contrast !== undefined)
    refuseImgPropsEffects(pictures(batch, part, ctx));
  for (const pic of pictures(batch, part, ctx)) {
    const blipFill = partOf(pic, 'p:blipFill');
    const blip = partOf(blipFill, 'a:blip');
    const now = adjustmentsOf(blip);
    const put = (names: ReadonlySet<string>, xml: string | null, at: 'start' | 'end'): void => {
      for (const e of elements(batch.element(part, blip.id)))
        if (e.ns === NS.a && names.has(e.local)) remove(batch, part, e);
      if (xml === null) return;
      const kids = batch.element(part, blip.id).children;
      const ext = kids.findIndex((c) => isA(c, 'extLst'));
      const lastRecolor = kids.reduce(
        (n, c, i) =>
          c.type === 'element' && c.ns === NS.a && RECOLOR_EFFECTS.has(c.local) ? i + 1 : n,
        0
      );
      const index = at === 'start' ? lastRecolor : ext < 0 ? kids.length : ext;
      batch.insert({ op: 'insertNode', part, parent: blip.id, index, spec: specFromXml(xml) });
    };
    // Recolor first, then brightness and contrast, then transparency.
    if (adjust.recolor !== undefined) put(RECOLOR_EFFECTS, recolorXml(adjust.recolor), 'start');
    if (adjust.brightness !== undefined || adjust.contrast !== undefined) {
      const bright = Math.round((adjust.brightness ?? now.brightness) * 100_000);
      const contrast = Math.round((adjust.contrast ?? now.contrast) * 100_000);
      const xml =
        bright || contrast
          ? `<a:lum${bright ? ` bright="${bright}"` : ''}${contrast ? ` contrast="${contrast}"` : ''}/>`
          : null;
      put(new Set(['lum']), xml, 'start');
    }
    if (adjust.transparency !== undefined) {
      const amt = Math.round((1 - adjust.transparency) * 100_000);
      put(new Set(['alphaModFix']), amt < 100_000 ? `<a:alphaModFix amt="${amt}"/>` : null, 'end');
    }
  }
}

function setStyle(batch: OpBatch, part: string, ctx: CommandContext, id: string): void {
  if (typeof id !== 'string' || !Object.hasOwn(PICTURE_STYLES, id))
    throw new CommandRefusal('errors.unsupported', id);
  const style = PICTURE_STYLES[id as keyof typeof PICTURE_STYLES];
  for (const pic of pictures(batch, part, ctx)) {
    const spPr = partOf(pic, 'p:spPr');
    const geom = child(spPr, 'a:prstGeom');
    const prst = /prst="([^"]+)"/.exec(style.geometry)?.[1];
    if (
      pictureLock(pic, 'noChangeShapeType') &&
      (child(spPr, 'a:custGeom') || (attr(geom, 'prst') ?? 'rect') !== prst)
    )
      throw new CommandRefusal('disabled.shapeLocked');
    setSpPrChild(batch, part, spPr, ['prstGeom', 'custGeom'], style.geometry);
    setSpPrChild(batch, part, spPr, ['ln'], style.ln);
    setSpPrChild(batch, part, spPr, ['effectLst', 'effectDag'], style.effectLst);
    setSpPrChild(batch, part, spPr, ['scene3d', 'sp3d'], null);
    // A style reference would bring back a theme line or effect the style leaves out.
    const ref = child(pic, 'p:style');
    if (ref) remove(batch, part, ref);
  }
}

/** The `p:cNvPr` of a shape, picture, group, connector, or frame. */
export function nonVisualOf(el: Element): Element | undefined {
  const nv = elements(el).find((c) => c.ns === NS.p && c.local.startsWith('nv'));
  return child(nv, 'p:cNvPr');
}

/** The `adec:decorative` extension of a `p:cNvPr`, when present. */
export function decorativeExt(cNvPr: Element | undefined): Element | undefined {
  return elements(child(cNvPr, 'a:extLst')).find((e) => attr(e, 'uri') === DECORATIVE_EXT);
}

/** Whether a `p:cNvPr` is marked as decorative. */
export function isDecorative(cNvPr: Element | undefined): boolean {
  const flag = elements(decorativeExt(cNvPr)).find(
    (e) => e.ns === ADEC && e.local === 'decorative'
  );
  const v = attr(flag, 'val');
  return flag !== undefined && (v === '1' || v === 'true');
}

function setAltText(
  batch: OpBatch,
  part: string,
  ctx: CommandContext,
  command: Extract<EditorCommand, { type: 'setAltText' }>
): void {
  const ids = selectedIds(ctx.selection);
  if (ids.length === 0) throw new CommandRefusal('disabled.noSelection');
  if (ids.length !== 1) throw new CommandRefusal('disabled.selectOne');
  const el = batch.element(part, ids[0]);
  const cNvPr = nonVisualOf(el);
  if (!cNvPr) throw new CommandRefusal('disabled.unsupportedElement');
  knownKeys(command, ['type', 'description', 'title', 'decorative']);
  if (
    command.description === undefined &&
    command.title === undefined &&
    command.decorative === undefined
  )
    throw new CommandRefusal('errors.unsupported', 'setAltText');
  for (const text of [command.description, command.title])
    if (text !== undefined && typeof text !== 'string')
      throw new CommandRefusal('errors.invalidText');
  if (command.decorative !== undefined && typeof command.decorative !== 'boolean')
    throw new CommandRefusal('errors.unsupported', 'decorative');
  // Absent fields keep what the object has, as in the other picture commands: a title the
  // command does not give stays as the file has it, untrimmed.
  const decorative = command.decorative ?? isDecorative(cNvPr);
  const given = command.description?.trim();
  // A decorative object has no description: a new one must clear the mark explicitly.
  if (decorative && given && command.decorative !== false)
    throw new CommandRefusal('disabled.decorative');
  // A title-only edit leaves the stored description alone, even on a decorative object.
  const description =
    command.description === undefined && command.decorative === undefined
      ? undefined
      : decorative
        ? ''
        : given;
  const title = command.title?.trim();
  for (const text of [description ?? '', title ?? '']) {
    if (text.length > MAX_ALT_TEXT) throw new CommandRefusal('errors.altTextLength');
    // Checked here for a clear reason; the store refuses such text too.
    if (!isValidXmlText(text)) throw new CommandRefusal('errors.invalidText');
  }
  // An empty value removes the attribute; the serializer escapes the rest (`escapeXmlAttributeChecked`).
  if (description !== undefined) setAttr(batch, part, cNvPr, 'descr', description || null);
  if (title !== undefined) setAttr(batch, part, cNvPr, 'title', title || null);
  // An unchanged mark stays as the file has it, an `adec:decorative val="0"` included.
  if (decorative === isDecorative(cNvPr)) return;
  const ext = decorativeExt(cNvPr);
  if (ext) remove(batch, part, ext);
  const extLst = child(batch.element(part, cNvPr.id), 'a:extLst');
  if (!decorative) {
    if (extLst && elements(batch.element(part, extLst.id)).length === 0)
      remove(batch, part, extLst);
    return;
  }
  const xml = `<a:ext uri="${DECORATIVE_EXT}"><adec:decorative xmlns:adec="${ADEC}" val="1"/></a:ext>`;
  if (extLst) {
    batch.insert({
      op: 'insertNode',
      part,
      parent: extLst.id,
      index: batch.element(part, extLst.id).children.length,
      spec: specFromXml(xml),
    });
    return;
  }
  // `a:extLst` is the last child of CT_NonVisualDrawingProps.
  batch.insert({
    op: 'insertNode',
    part,
    parent: cNvPr.id,
    index: batch.element(part, cNvPr.id).children.length,
    spec: specFromXml(`<a:extLst>${xml}</a:extLst>`),
  });
}

/**
 * A frame change of a group child refits the group's bounds, as a move does (`setTransforms`
 * in commands.ts): the parent comes from the tree, not from the selection.
 */
function refitParents(batch: OpBatch, part: string, ctx: CommandContext): void {
  const items = ctx.session.layout(ctx.slide)?.items ?? [];
  const groups = new Set(
    selectedIds(ctx.selection)
      .map((id) => parentOf(batch, part, id))
      .filter((p) => p.kind === 'grpSp')
  );
  for (const g of groups) refitGroup(batch, part, g.id, items);
}

/** The fields of each picture command; any other field is refused. */
/** Whether `command` is one of the picture commands that `compilePictureCommand` takes. */
export function isPictureCommand(command: EditorCommand): command is PictureCommand {
  return Object.hasOwn(COMMAND_KEYS, command.type);
}

const COMMAND_KEYS: Readonly<Record<PictureCommand['type'], readonly string[]>> = {
  cropImage: ['type', 'crop'],
  cropToShape: ['type', 'preset'],
  replaceImage: ['type', 'bytes', 'mime', 'width', 'height'],
  resetImage: ['type', 'size'],
  setImageAdjustments: ['type', 'adjustments'],
  setImageStyle: ['type', 'style'],
  setImageProperties: PROPERTY_KEYS,
  setAltText: ['type', 'description', 'title', 'decorative'],
};

/** Compile a picture command against the current slide. Throws CommandRefusal. */
export function compilePictureCommand(
  ctx: CommandContext,
  batch: OpBatch,
  part: string,
  command: PictureCommand
): CommandEffect {
  // Alt text has its own command (DOCX Editor's `setImageProperties` takes it): say so.
  if (command.type === 'setImageProperties')
    for (const key of ['title', 'description'])
      if (key in command) throw new CommandRefusal('errors.unsupported', `${key}: use setAltText`);
  // Ops are plain JSON: a field the command does not know (a DOCX Editor `drawingNodeId`, a
  // typo) is refused, never ignored.
  knownKeys(command, COMMAND_KEYS[command.type]);
  if (
    command.type === 'resetImage' &&
    command.size !== undefined &&
    typeof command.size !== 'boolean'
  )
    throw new CommandRefusal('errors.unsupported', 'size');
  switch (command.type) {
    case 'cropImage':
      // Fill, Fit, and a crop that keeps the frame change only `a:srcRect`: no group refits.
      if (crop(batch, part, ctx, command.crop)) refitParents(batch, part, ctx);
      return { label: 'cropImage' };
    case 'cropToShape':
      cropToShape(batch, part, ctx, command.preset);
      return { label: 'cropToShape' };
    case 'replaceImage':
      replaceImage(batch, part, ctx, command);
      refitParents(batch, part, ctx);
      return { label: 'replaceImage' };
    case 'resetImage':
      resetImage(batch, part, ctx, command.size ?? false);
      if (command.size) refitParents(batch, part, ctx);
      return { label: 'resetImage' };
    case 'setImageAdjustments':
      setAdjustments(batch, part, ctx, command.adjustments);
      return { label: 'pictureAdjustments' };
    case 'setImageStyle':
      setStyle(batch, part, ctx, command.style);
      return { label: 'pictureStyle' };
    case 'setImageProperties':
      setProperties(batch, part, ctx, command);
      refitParents(batch, part, ctx);
      return { label: 'imageProperties' };
    case 'setAltText':
      setAltText(batch, part, ctx, command);
      return { label: 'altText' };
  }
}
