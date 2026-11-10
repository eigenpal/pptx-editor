import { describe, expect, test } from 'bun:test';
import { PNG } from 'pngjs';
import { buildPptx, shape, solid } from '../../../../../scripts/fixtures/builder.ts';
import { readPptx, serializeXmlPart, xmlPartOf } from '../../store/index.ts';
import { resolveSlide } from '../../cascade/index.ts';
import {
  CROP_SHAPE_GALLERY,
  canvasMenuKind,
  chromeControlForSlot,
  commandForSlotValue,
  createPptxEditor,
  toolbarCommandState,
  type PptxEditorInstance,
} from '../index.ts';
import { croppedBox, fillOrFitCrop } from '../picture-commands.ts';
import {
  altTextDraftCommand,
  picturePropertiesDraftCommand,
  picturePropertiesDraftFrom,
  validateAltTextDraft,
  validatePicturePropertiesDraft,
} from '../picture-fields.ts';
import type { EditorCommand } from '../../contracts/index.ts';
import { cropAt, startCropDrag } from '../surface-crop.ts';
import { itemById } from '../surface.ts';

const IN = 914_400;
const R_IMAGE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';

/** A `w`×`h` opaque PNG. */
function png(w: number, h: number): Uint8Array {
  const img = new PNG({ width: w, height: h });
  for (let i = 0; i < w * h * 4; i += 4) img.data.set([200, 80, 40, 255], i);
  return new Uint8Array(PNG.sync.write(img));
}

const pic = (id: number, o: { blip?: string; crop?: string; spPr?: string; rot?: number } = {}) =>
  `<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="Picture ${id}"/><p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr><p:nvPr/></p:nvPicPr>` +
  `<p:blipFill><a:blip r:embed="rIdImg">${o.blip ?? ''}</a:blip>${o.crop ?? ''}<a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
  `<p:spPr><a:xfrm${o.rot ? ` rot="${o.rot * 60_000}"` : ''}><a:off x="${IN}" y="${IN}"/><a:ext cx="${4 * IN}" cy="${3 * IN}"/></a:xfrm>` +
  `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>${o.spPr ?? ''}</p:spPr></p:pic>`;

const deck = (shapes: string): Uint8Array =>
  buildPptx({
    slides: [{ layout: 'blank', rels: [['rIdImg', R_IMAGE, '../media/image1.png']], shapes }],
    media: [{ name: 'image1.png', bytes: png(128, 96), contentType: 'image/png' }],
  });

async function open(shapes = pic(2)): Promise<PptxEditorInstance> {
  const e = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await e.load(deck(shapes))).ok).toBe(true);
  return e;
}

const nodeOf = (e: PptxEditorInstance, i = 0): string =>
  e.session!.layout(0)!.items[i].element.nodeId;
function select(e: PptxEditorInstance, i = 0): string {
  const id = nodeOf(e, i);
  e.setSelection({ kind: 'shapes', nodeIds: [id] });
  return id;
}
const slideXml = (e: PptxEditorInstance): string =>
  serializeXmlPart(xmlPartOf(e.session!.store.package, e.session!.info.slides[0].partName)!);
const box = (e: PptxEditorInstance) => {
  const g = e.snapshot().geometry!;
  return { x: g.x, y: g.y, cx: g.cx, cy: g.cy };
};

async function openSelected(shapes?: string): Promise<PptxEditorInstance> {
  const e = await open(shapes);
  e.selectAll();
  return e;
}

describe('crop', () => {
  test('a crop rectangle keeps the image in place and shrinks the frame', async () => {
    const e = await openSelected();
    const r = e.exec({
      type: 'cropImage',
      crop: { kind: 'rect', left: 25, top: 0, right: 0, bottom: 50 },
    });
    expect(r.ok).toBe(true);
    expect(box(e)).toEqual({ x: 2 * IN, y: IN, cx: 3 * IN, cy: 1.5 * IN });
    expect(slideXml(e)).toContain('<a:srcRect l="25000" b="50000"/>');
    expect(e.snapshot().image?.crop).toEqual({ left: 25, top: 0, right: 0, bottom: 50 });
    // A second crop goes back out: the frame grows over the same image.
    e.exec({ type: 'cropImage', crop: { kind: 'rect', left: 0, top: 0, right: 0, bottom: 0 } });
    expect(box(e)).toEqual({ x: IN, y: IN, cx: 4 * IN, cy: 3 * IN });
    expect(slideXml(e)).not.toContain('srcRect');
    e.exec({ type: 'undo' });
    expect(box(e)).toEqual({ x: 2 * IN, y: IN, cx: 3 * IN, cy: 1.5 * IN });
  });

  test('a rotated frame moves its center along its own axes', () => {
    // Cropping a quarter off the left of a frame turned 90° moves its center down by an eighth.
    const b = croppedBox(
      { x: 0, y: 0, cx: 4 * IN, cy: 2 * IN, rot: 90 },
      { h: false, v: false },
      { l: 0, t: 0, r: 0, b: 0 },
      { l: 0.25, t: 0, r: 0, b: 0 }
    );
    expect(b.cx).toBeCloseTo(3 * IN);
    expect(b.x + b.cx / 2).toBeCloseTo(2 * IN);
    expect(b.y + b.cy / 2).toBeCloseTo(IN + 0.5 * IN);
    // A flip mirrors the move.
    const f = croppedBox(
      { x: 0, y: 0, cx: 4 * IN, cy: 2 * IN, rot: 0 },
      { h: true, v: false },
      { l: 0, t: 0, r: 0, b: 0 },
      { l: 0.25, t: 0, r: 0, b: 0 }
    );
    expect(f.x).toBeCloseTo(0);
  });

  test('keepFrame moves the image inside the frame', async () => {
    const e = await openSelected(pic(2, { crop: '<a:srcRect l="10000" r="10000"/>' }));
    const before = box(e);
    e.exec({
      type: 'cropImage',
      crop: { kind: 'rect', left: 20, top: 0, right: 0, bottom: 0, keepFrame: true },
    });
    expect(box(e)).toEqual(before);
    expect(slideXml(e)).toContain('<a:srcRect l="20000"/>');
  });

  test('fill, fit, and aspect ratio', async () => {
    // The image is 4:3; a 2:1 frame filled crops the top and bottom, fitted pads the sides.
    expect(fillOrFitCrop('fill', 4 / 3, 2).t).toBeCloseTo(1 / 6);
    expect(fillOrFitCrop('fit', 4 / 3, 2).l).toBeCloseTo(-0.25);
    expect(fillOrFitCrop('fill', 2, 1).l).toBeCloseTo(0.25);
    expect(fillOrFitCrop('fit', 2, 1).t).toBeCloseTo(-0.5);
    const e = await openSelected();
    e.exec({
      type: 'setTransform',
      nodeId: nodeOf(e),
      x: IN,
      y: IN,
      cx: 4 * IN,
      cy: 2 * IN,
      rot: 0,
    });
    expect(e.exec({ type: 'cropImage', crop: { kind: 'fill' } }).ok).toBe(true);
    expect(box(e).cy).toBe(2 * IN);
    expect(slideXml(e)).toContain('<a:srcRect t="16667" b="16667"/>');
    e.exec({ type: 'cropImage', crop: { kind: 'fit' } });
    expect(slideXml(e)).toContain('<a:srcRect l="-25000" r="-25000"/>');
    const sq = await openSelected();
    sq.exec({ type: 'cropImage', crop: { kind: 'aspect', width: 1, height: 1 } });
    expect(box(sq)).toEqual({ x: 1.5 * IN, y: IN, cx: 3 * IN, cy: 3 * IN });
    expect(slideXml(sq)).toContain('<a:srcRect l="12500" r="12500"/>');
  });

  test('a crop that leaves nothing visible is refused', async () => {
    const e = await openSelected();
    const bad = {
      type: 'cropImage',
      crop: { kind: 'rect', left: 60, top: 0, right: 50, bottom: 0 },
    } as const;
    expect(e.can(bad).ok).toBe(false);
    const r = e.exec(bad);
    expect(r.ok).toBe(false);
    expect(slideXml(e)).not.toContain('srcRect');
    const nan = {
      type: 'cropImage',
      crop: { kind: 'rect', left: NaN, top: 0, right: 0, bottom: 0 },
    } as const;
    expect(e.exec(nan).ok).toBe(false);
  });

  test('picture commands refuse a shape that is not a picture', async () => {
    const e = await openSelected(
      shape({ id: 2, x: IN, y: IN, cx: IN, cy: IN, fill: solid('FF0000') })
    );
    const r = e.can({ type: 'cropToShape', preset: 'ellipse' });
    expect(r).toEqual({ ok: false, reason: 'disabled.noPicture' });
    expect(e.can({ type: 'toggleCrop' })).toEqual({ ok: false, reason: 'disabled.noPicture' });
    expect(e.snapshot().image).toBeNull();
  });

  test('a crop drag on the surface: a handle crops, the image moves', async () => {
    const e = await openSelected();
    const item = itemById(e.session!.layout(0), nodeOf(e))!;
    const left = startCropDrag(item, 'w', { x: IN, y: 2 * IN }, 'g1')!;
    expect(cropAt(left, { x: 2 * IN, y: 2 * IN })).toEqual({
      kind: 'rect',
      left: 25,
      top: 0,
      right: 0,
      bottom: 0,
    });
    // An edge stops short of the opposite edge.
    const far = cropAt(left, { x: 9 * IN, y: 2 * IN });
    expect(far.kind === 'rect' && far.left).toBeCloseTo(99);
    const move = startCropDrag(item, 'move', { x: 2 * IN, y: 2 * IN }, 'g2')!;
    expect(cropAt(move, { x: 3 * IN, y: 2 * IN })).toEqual({
      kind: 'rect',
      left: -25,
      top: 0,
      right: 25,
      bottom: 0,
      keepFrame: true,
    });
    // A move far past the image stops at the limit of the crop range: the command takes it.
    for (const x of [400 * IN, -4000 * IN]) {
      const stop = cropAt(move, { x, y: 2 * IN });
      expect(stop.kind === 'rect' && stop.left < 100 && stop.right < 100).toBe(true);
      expect(stop.kind === 'rect' && stop.left > -100_000 && stop.right > -100_000).toBe(true);
      expect(e.can({ type: 'cropImage', crop: stop }).ok).toBe(true);
    }
    // A handle dragged far outside stops at the padding limit.
    const out = cropAt(left, { x: -9000 * IN, y: 2 * IN });
    expect(out.kind === 'rect' && out.left).toBe(-100_000);
    expect(e.can({ type: 'cropImage', crop: out }).ok).toBe(true);
  });
});

describe('Format Picture', () => {
  test('size, rotation, and crop apply as one undo step', async () => {
    const e = await openSelected();
    const r = e.exec({
      type: 'setImageProperties',
      crop: { left: 25, top: 0, right: 0, bottom: 0 },
      heightEmu: 2 * IN,
      rotationDegrees: 30,
    });
    expect(r.ok).toBe(true);
    // The crop moves the left edge (the image stays); the height then changes, top-left kept.
    expect(box(e)).toEqual({ x: 2 * IN, y: IN, cx: 3 * IN, cy: 2 * IN });
    expect(e.snapshot().geometry?.rot).toBe(30);
    expect(slideXml(e)).toContain('<a:srcRect l="25000"/>');
    e.exec({ type: 'undo' });
    expect(box(e)).toEqual({ x: IN, y: IN, cx: 4 * IN, cy: 3 * IN });
    expect(slideXml(e)).not.toContain('srcRect');
  });

  test('the size is the size on the slide inside a scaled group', async () => {
    const group =
      `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="10" name="Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
      `<p:grpSpPr><a:xfrm><a:off x="${IN}" y="${IN}"/><a:ext cx="${2 * IN}" cy="${1.5 * IN}"/>` +
      `<a:chOff x="0" y="0"/><a:chExt cx="${4 * IN}" cy="${3 * IN}"/></a:xfrm></p:grpSpPr>` +
      pic(2).replace(`<a:off x="${IN}" y="${IN}"/>`, '<a:off x="0" y="0"/>') +
      '</p:grpSp>';
    const e = await open(group);
    const items = e.session!.layout(0)!.items;
    const grp = items[0]!;
    const child = grp.kind === 'group' ? grp.children[0]!.element.nodeId : '';
    e.setSelection({ kind: 'shapes', nodeIds: [child], group: grp.element.nodeId });
    // The group halves its child space: the file says 4 in, the slide shows 2 in.
    expect(e.snapshot().geometry?.cx).toBe(4 * IN);
    expect(e.snapshot().image?.widthEmu).toBe(2 * IN);
    expect(e.exec({ type: 'setImageProperties', widthEmu: IN }).ok).toBe(true);
    expect(e.snapshot().image?.widthEmu).toBe(IN);
    expect(slideXml(e)).toContain(`<a:ext cx="${2 * IN}" cy="${3 * IN}"/>`);
  });

  test('the dialog checks only the sides the user changed', () => {
    // A picture wider than the largest side (56 in) still takes a new rotation.
    const none = { left: 0, top: 0, right: 0, bottom: 0 };
    const initial = picturePropertiesDraftFrom({
      widthEmu: 60 * IN,
      heightEmu: IN,
      rotationDegrees: 0,
      crop: none,
    });
    expect(validatePicturePropertiesDraft({ ...initial, rotation: 45 }, initial)).toEqual({});
    expect(validatePicturePropertiesDraft({ ...initial, width: 57 }, initial)).toEqual({
      width: 'imageProperties.sizeRange',
    });
    expect(validatePicturePropertiesDraft(initial).width).toBe('imageProperties.sizeRange');
    // A Fit crop pads far more than the dialog offers; an untouched edge does not block OK.
    const fit = picturePropertiesDraftFrom({
      widthEmu: IN,
      heightEmu: IN,
      rotationDegrees: 0,
      crop: { ...none, left: -4950, right: -4950 },
    });
    expect(validatePicturePropertiesDraft({ ...fit, rotation: 10 }, fit)).toEqual({});
    expect(validatePicturePropertiesDraft({ ...fit, cropLeft: -1500 }, fit).cropLeft).toBe(
      'imageProperties.cropPadRange'
    );
  });

  test('a command with no field is refused, as DOCX Editor refuses it', async () => {
    const e = await openSelected();
    expect(e.exec({ type: 'setImageProperties' })).toMatchObject({ ok: false });
    expect(e.snapshot().canUndo).toBe(false);
  });

  test('alt text: an absent title keeps the title; an empty one removes it', async () => {
    const e = await openSelected();
    e.exec({ type: 'setAltText', description: 'A', title: 'T' });
    e.exec({ type: 'setAltText', description: 'B' });
    expect(e.snapshot().geometry?.altText).toEqual({
      description: 'B',
      title: 'T',
      decorative: false,
    });
    e.exec({ type: 'setAltText', description: 'B', title: '' });
    expect(slideXml(e)).not.toContain('title=');
  });

  test('alt text: a title the command does not give stays as the file has it', async () => {
    const e = await openSelected(
      pic(2).replace('name="Picture 2"', 'name="Picture 2" title=" T "')
    );
    e.exec({ type: 'setAltText', description: 'A' });
    expect(slideXml(e)).toContain('title=" T "');
    // A value that is not text is refused, not thrown.
    expect(e.can({ type: 'setAltText', description: 3 as unknown as string })).toEqual({
      ok: false,
      reason: 'errors.invalidText',
    });
  });

  test('alt text: a description for a decorative object must clear the mark', async () => {
    const e = await openSelected();
    e.exec({ type: 'setAltText', decorative: true });
    expect(e.can({ type: 'setAltText', description: 'A chart' })).toEqual({
      ok: false,
      reason: 'disabled.decorative',
    });
    expect(e.exec({ type: 'setAltText', description: 'A chart', decorative: false }).ok).toBe(true);
    expect(e.snapshot().geometry?.altText).toEqual({
      description: 'A chart',
      title: '',
      decorative: false,
    });
  });

  test('the Alt Text dialog sends only the fields that changed', () => {
    const initial = { description: 'A', title: ' T ', decorative: false };
    expect(altTextDraftCommand(initial, initial)).toBeNull();
    expect(altTextDraftCommand({ ...initial, description: 'B' }, initial)).toEqual({
      type: 'setAltText',
      description: 'B',
    });
    expect(altTextDraftCommand({ ...initial, decorative: true }, initial)).toEqual({
      type: 'setAltText',
      decorative: true,
    });
    // A decorative object keeps no description, so its length does not block OK.
    const long = { ...initial, description: 'x'.repeat(5000), decorative: true };
    expect(validateAltTextDraft(long)).toEqual({});
  });

  test('bad values are refused before anything changes', async () => {
    const e = await openSelected();
    for (const properties of [
      { widthEmu: 0 },
      { heightEmu: -1 },
      { rotationDegrees: NaN },
      { widthEmu: 1e15 },
    ])
      expect(e.can({ type: 'setImageProperties', ...properties }).ok).toBe(false);
    expect(
      e.can({ type: 'setImageProperties', crop: { left: 70, top: 0, right: 40, bottom: 0 } }).ok
    ).toBe(false);
  });

  test('a payload the command does not know is refused with a reason, never guessed', async () => {
    const e = await openSelected();
    const bad = [
      // DOCX Editor's alt text fields belong to setAltText here.
      { type: 'setImageProperties', description: 'x' },
      { type: 'setImageProperties', properties: { widthEmu: IN } },
      { type: 'setImageAdjustments', adjustments: { saturation: 0.5 } },
      { type: 'setImageAdjustments' },
      { type: 'cropImage', crop: { kind: 'bogus' } },
      { type: 'cropImage', crop: { left: 10, top: 0, right: 0, bottom: 0 } },
      { type: 'cropImage', crop: { kind: 'fill', left: 10 } },
      { type: 'cropImage' },
      { type: 'replaceImage', mime: 'image/png' },
      { type: 'setAltText', decorative: 'yes' },
      // A field at the top level of any picture command, and values of the wrong type.
      { type: 'replaceImage', bytes: png(2, 2), mime: 'image/png', drawingNodeId: 'other' },
      { type: 'cropImage', crop: { kind: 'fill' }, extra: 1 },
      { type: 'setImageStyle', style: 'softEdgeOval', target: 'x' },
      { type: 'resetImage', size: 'yes' },
      { type: 'setImageProperties', crop: { left: '10', top: 0, right: 0, bottom: 0 } },
      {
        type: 'cropImage',
        crop: { kind: 'rect', left: 10, top: 0, right: 0, bottom: 0, keepFrame: 1 },
      },
      { type: 'replaceImage', bytes: png(2, 2), mime: 'image/png', width: 100 },
      { type: 'replaceImage', bytes: png(2, 2), mime: 'image/png', width: '100', height: '50' },
      { type: 'setImageStyle', style: ['simpleFrameWhite'] },
      { type: 'setImageAdjustments', adjustments: { brightness: undefined } },
      { type: 'toggleCrop', foo: 1 },
    ] as unknown as EditorCommand[];
    for (const command of bad) {
      expect(e.can(command).ok).toBe(false);
      expect(() => e.exec(command)).not.toThrow();
      expect(e.exec(command).ok).toBe(false);
    }
    expect(e.snapshot().canUndo).toBe(false);
  });
});

describe('picture locks', () => {
  test('a picture locked against cropping refuses crops and crop mode; its aspect lock shows', async () => {
    const locked = pic(2).replace(
      '<a:picLocks noChangeAspect="1"/>',
      '<a:picLocks noChangeAspect="1" noCrop="1"/>'
    );
    const e = await openSelected(locked);
    expect(e.snapshot().image).toMatchObject({ lockAspect: true, canCrop: false });
    for (const command of [
      { type: 'toggleCrop' },
      { type: 'cropImage', crop: { kind: 'fill' } },
      { type: 'setImageProperties', crop: { left: 10, top: 0, right: 0, bottom: 0 } },
    ] as const)
      expect(e.can(command)).toEqual({ ok: false, reason: 'disabled.cropLocked' });
    // Size and shape still change.
    expect(e.can({ type: 'setImageProperties', widthEmu: IN }).ok).toBe(true);
    expect(e.can({ type: 'cropToShape', preset: 'ellipse' }).ok).toBe(true);
    const free = await openSelected(pic(2).replace('<a:picLocks noChangeAspect="1"/>', ''));
    expect(free.snapshot().image).toMatchObject({ lockAspect: false, canCrop: true });
    expect(picturePropertiesDraftFrom(free.snapshot().image).lockAspect).toBe(false);
  });

  test('Reset Picture & Size keeps a locked crop: refused before any op', async () => {
    const locked = pic(2, { crop: '<a:srcRect l="10000"/>' }).replace(
      '<a:picLocks noChangeAspect="1"/>',
      '<a:picLocks noChangeAspect="1" noCrop="1"/>'
    );
    const e = await openSelected(locked);
    const before = e.snapshot().revision;
    expect(e.exec({ type: 'resetImage', size: true })).toMatchObject({
      ok: false,
      reason: 'disabled.cropLocked',
    });
    expect(e.snapshot().revision).toBe(before);
    // Without a crop, nothing the lock holds changes: it runs; Reset Picture alone runs too.
    const plain = await openSelected(locked.replace('<a:srcRect l="10000"/>', ''));
    expect(plain.can({ type: 'resetImage', size: true }).ok).toBe(true);
    expect(e.can({ type: 'resetImage' }).ok).toBe(true);
  });

  test('a crop that changes the frame size is refused for a picture locked against resizing', async () => {
    const locked = pic(2).replace(
      '<a:picLocks noChangeAspect="1"/>',
      '<a:picLocks noChangeAspect="1" noResize="1"/>'
    );
    const e = await openSelected(locked);
    const edge = { left: 10, top: 0, right: 0, bottom: 0 };
    for (const command of [
      { type: 'cropImage', crop: { kind: 'rect', ...edge } },
      { type: 'cropImage', crop: { kind: 'aspect', width: 1, height: 1 } },
      { type: 'setImageProperties', crop: edge },
    ] as const)
      expect(e.can(command)).toEqual({ ok: false, reason: 'disabled.sizeLocked' });
    // Fill, Fit, and a move of the image in its frame keep the frame: they pass.
    for (const crop of [
      { kind: 'fill' },
      { kind: 'fit' },
      { kind: 'rect', ...edge, keepFrame: true },
    ] as const)
      expect(e.can({ type: 'cropImage', crop }).ok).toBe(true);
  });

  test('a tiled or inset picture fill refuses crop mode and crops', async () => {
    const tiled = pic(2).replace(
      '<a:stretch><a:fillRect/></a:stretch>',
      '<a:tile tx="0" ty="0" sx="100000" sy="100000" flip="none" algn="tl"/>'
    );
    const inset = pic(2).replace('<a:fillRect/>', '<a:fillRect l="10000" r="10000"/>');
    for (const shapes of [tiled, inset]) {
      const e = await openSelected(shapes);
      for (const command of [
        { type: 'toggleCrop' },
        { type: 'cropImage', crop: { kind: 'fill' } },
        { type: 'cropImage', crop: { kind: 'rect', left: 10, top: 0, right: 0, bottom: 0 } },
        { type: 'setImageProperties', crop: { left: 10, top: 0, right: 0, bottom: 0 } },
      ] as const)
        expect(e.can(command)).toEqual({ ok: false, reason: 'disabled.cropFillMode' });
      // The toolbar shows the reason on Crop.
      expect(toolbarCommandState(e, 'image.crop').disabledReason).toBe('disabled.cropFillMode');
      expect(e.can({ type: 'setImageProperties', widthEmu: IN }).ok).toBe(true);
    }
  });

  test('empty commands are refused alike; the crop gallery has no lines', async () => {
    const e = await openSelected();
    expect(e.can({ type: 'setAltText' }).ok).toBe(false);
    expect(e.can({ type: 'setImageAdjustments', adjustments: {} }).ok).toBe(false);
    const presets = CROP_SHAPE_GALLERY.flatMap((c) => c.presets);
    expect(presets).toContain('ellipse');
    for (const line of ['line', 'straightConnector1', 'bentConnector3', 'curvedConnector3'])
      expect(presets).not.toContain(line);
  });
});

describe('crop mode', () => {
  test('the Crop button turns crop mode on and off; another selection ends it', async () => {
    const e = await open(pic(2) + shape({ id: 3, x: 6 * IN, y: IN, cx: IN, cy: IN }));
    expect(e.can({ type: 'toggleCrop' }).ok).toBe(false);
    e.selectAll();
    expect(e.can({ type: 'toggleCrop' })).toEqual({ ok: false, reason: 'disabled.noPicture' });
    select(e, 0);
    expect(e.snapshot().image?.cropping).toBe(false);
    expect(e.exec({ type: 'toggleCrop' }).ok).toBe(true);
    expect(e.snapshot().image?.cropping).toBe(true);
    expect(e.snapshot().canUndo).toBe(false);
    e.exec({ type: 'toggleCrop' });
    expect(e.snapshot().image?.cropping).toBe(false);
    e.exec({ type: 'toggleCrop' });
    select(e, 1);
    select(e, 0);
    expect(e.snapshot().image?.cropping).toBe(false);
    // Viewing ends crop mode.
    e.exec({ type: 'toggleCrop' });
    e.setMode('view');
    expect(e.snapshot().image?.cropping).toBe(false);
    e.setMode('edit');
    expect(e.snapshot().image?.cropping).toBe(false);
  });

  test('change picture refuses a type name that only an object prototype has', async () => {
    const e = await openSelected();
    for (const mime of ['constructor', 'toString', '__proto__'])
      expect(e.can({ type: 'replaceImage', bytes: png(2, 2), mime, width: 2, height: 2 }).ok).toBe(
        false
      );
    expect(() =>
      e.exec({ type: 'replaceImage', bytes: png(2, 2), mime: 'constructor', width: 2, height: 2 })
    ).not.toThrow();
  });
});

describe('picture chrome', () => {
  test('image.altText takes the description as its value, as DOCX Editor', async () => {
    const e = await open(pic(2).replace('name="Picture 2"', 'name="Picture 2" descr="A chart"'));
    select(e, 0);
    expect(toolbarCommandState(e, 'image.altText').value).toBe('A chart');
    const command = commandForSlotValue('image.altText', 'Sales');
    expect(command).toEqual({ type: 'setAltText', description: 'Sales' });
    expect(e.exec(command!).ok).toBe(true);
    expect(toolbarCommandState(e, 'image.altText').value).toBe('Sales');
    expect(commandForSlotValue('image.altText', 3)).toBeNull();
  });

  test('Format picture on a shape asks for a picture', async () => {
    const e = await open(pic(2) + shape({ id: 3, x: 6 * IN, y: IN, cx: IN, cy: IN }));
    expect(toolbarCommandState(e, 'image.properties').disabledReason).toBe('disabled.noSelection');
    select(e, 1);
    expect(toolbarCommandState(e, 'image.properties').disabledReason).toBe('disabled.noPicture');
    select(e, 0);
    expect(toolbarCommandState(e, 'image.properties').disabledReason).toBeNull();
  });

  test('Corrections offers the 25 Brightness/Contrast presets', () => {
    const options = chromeControlForSlot('image.corrections')!.options!;
    expect(options.length).toBe(25);
    expect(new Set(options.map((o) => o.value)).size).toBe(25);
    expect(options.map((o) => o.value)).toContain('-40,40');
  });

  test('an image slot takes a string or a number only', () => {
    const tricky = { toString: () => 'grayscale' };
    for (const slot of ['image.color', 'image.styles', 'image.cropFit', 'image.reset'] as const) {
      expect(commandForSlotValue(slot, tricky)).toBeNull();
      expect(commandForSlotValue(slot, null)).toBeNull();
      expect(commandForSlotValue(slot, ['fill'])).toBeNull();
    }
    expect(commandForSlotValue('image.transparency', 50)).toEqual({
      type: 'setImageAdjustments',
      adjustments: { transparency: 0.5 },
    });
    expect(commandForSlotValue('image.color', 'grayscale')).toEqual({
      type: 'setImageAdjustments',
      adjustments: { recolor: 'grayscale' },
    });
  });

  test('pictures get the picture context menu; a shape among them gets the shape menu', async () => {
    const e = await open(pic(2) + shape({ id: 3, x: 6 * IN, y: IN, cx: IN, cy: IN }));
    expect(canvasMenuKind(e.snapshot())).toBe('slide');
    select(e, 0);
    expect(canvasMenuKind(e.snapshot())).toBe('picture');
    e.selectAll();
    expect(canvasMenuKind(e.snapshot())).toBe('shapes');
  });
});

describe('crop to shape, change, and reset', () => {
  test('crop to shape writes the preset after the transform', async () => {
    const e = await openSelected();
    expect(e.exec({ type: 'cropToShape', preset: 'ellipse' }).ok).toBe(true);
    expect(slideXml(e)).toContain('</a:xfrm><a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>');
    expect(e.snapshot().image?.shape).toBe('ellipse');
    expect(e.can({ type: 'cropToShape', preset: 'line' }).ok).toBe(false);
    expect(e.can({ type: 'cropToShape', preset: 'notAShape' }).ok).toBe(false);
    const saved = readPptx(e.save()!);
    if (!saved.ok) throw new Error('save failed');
    const el = resolveSlide(saved.pkg, 0)!.elements[0];
    expect(el.kind === 'picture' && el.geometry.kind === 'preset' && el.geometry.preset).toBe(
      'ellipse'
    );
  });

  test('change picture adds a part, fits the frame, and keeps the adjustments', async () => {
    const e = await openSelected(
      pic(2, { blip: '<a:alphaModFix amt="50000"/>', crop: '<a:srcRect l="10000"/>' })
    );
    const before = box(e);
    // With no size from the host, the engine reads the size the image states.
    const r = e.exec({ type: 'replaceImage', bytes: png(10, 10), mime: 'image/png' });
    expect(r.ok).toBe(true);
    const xml = slideXml(e);
    expect(xml).not.toContain('rIdImg');
    expect(xml).toContain('<a:alphaModFix amt="50000"/>');
    expect(xml).not.toContain('srcRect');
    // A square image in a 3.6 × 3 in frame: 3 in square, about the old center.
    const b = box(e);
    expect(b.cy).toBe(before.cy);
    expect(b.cx).toBe(before.cy);
    expect(b.x + b.cx / 2).toBeCloseTo(before.x + before.cx / 2, -1);
    const saved = readPptx(e.save()!);
    expect(saved.ok).toBe(true);
    if (!saved.ok) return;
    // Save omits the unused image. The live store keeps it for undo and collaboration.
    expect([...saved.pkg.parts.keys()].filter((p) => p.startsWith('/ppt/media/'))).toEqual([
      '/ppt/media/image2.png',
    ]);
    expect(
      e.can({ type: 'replaceImage', bytes: png(1, 1), mime: 'text/html', width: 1, height: 1 }).ok
    ).toBe(false);
    // SVG needs a raster copy beside it, which the editor cannot draw: refused.
    expect(
      e.can({ type: 'replaceImage', bytes: png(1, 1), mime: 'image/svg+xml', width: 1, height: 1 })
    ).toEqual({ ok: false, reason: 'errors.unsupported' });
  });

  test('change picture keeps an image another picture still uses', async () => {
    const e = await open(pic(2) + pic(3).replace(`x="${IN}" y="${IN}"`, `x="${6 * IN}" y="${IN}"`));
    e.setSelection({ kind: 'shapes', nodeIds: [nodeOf(e, 0)] });
    e.exec({ type: 'replaceImage', bytes: png(10, 10), mime: 'image/png', width: 10, height: 10 });
    expect(slideXml(e)).toContain('r:embed="rIdImg"');
    const saved = readPptx(e.save()!);
    if (!saved.ok) throw new Error('save failed');
    expect([...saved.pkg.parts.keys()].filter((p) => p.startsWith('/ppt/media/')).length).toBe(2);
    e.exec({ type: 'undo' });
    expect(slideXml(e).match(/r:embed="rIdImg"/g)?.length).toBe(2);
  });

  test('reset picture removes adjustments and styles; reset size removes the crop', async () => {
    const e = await openSelected(
      pic(2, {
        blip: '<a:grayscl/><a:lum bright="20000"/><a:alphaModFix amt="50000"/>',
        crop: '<a:srcRect l="10000"/>',
        spPr: '<a:ln w="12700"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln><a:effectLst><a:softEdge rad="112500"/></a:effectLst>',
      })
    );
    e.exec({ type: 'cropToShape', preset: 'ellipse' });
    expect(e.exec({ type: 'resetImage' }).ok).toBe(true);
    let xml = slideXml(e);
    for (const gone of ['grayscl', 'a:lum', 'alphaModFix', 'a:ln', 'effectLst', 'ellipse'])
      expect(xml).not.toContain(gone);
    expect(xml).toContain('<a:srcRect l="10000"/>');
    expect(xml).toContain('prst="rect"');
    expect(e.exec({ type: 'resetImage', size: true }).ok).toBe(true);
    xml = slideXml(e);
    expect(xml).not.toContain('srcRect');
    // 128 × 96 px at 96 dpi.
    expect(box(e)).toEqual({ x: IN, y: IN, cx: 128 * 9525, cy: 96 * 9525 });
  });
});

describe('`a14:imgProps` picture effects and groups', () => {
  // An `a14:imgProps` extension keeps the original image in `a14:imgLayer`; `r:embed` holds
  // the image rendered with the effects.
  const IMG_PROPS =
    '<a:extLst><a:ext uri="{BEBA8EAE-BF5A-486C-A8C5-ECC9F3942E4B}">' +
    '<a14:imgProps xmlns:a14="http://schemas.microsoft.com/office/drawing/2010/main"/></a:ext></a:extLst>';

  test('reset and corrections are refused on a picture with `a14:imgProps` effects', async () => {
    const e = await openSelected(pic(2, { blip: IMG_PROPS }));
    for (const command of [
      { type: 'resetImage' },
      { type: 'setImageAdjustments', adjustments: { brightness: 0.2 } },
    ] as const)
      expect(e.can(command)).toEqual({ ok: false, reason: 'disabled.pictureEffects' });
    // Recolor and transparency act on the image that `r:embed` holds.
    expect(
      e.exec({
        type: 'setImageAdjustments',
        adjustments: { recolor: 'grayscale', transparency: 0.5 },
      }).ok
    ).toBe(true);
    // A crop does not touch the blip, so it still works, and the extension stays.
    expect(
      e.exec({ type: 'cropImage', crop: { kind: 'rect', left: 10, top: 0, right: 0, bottom: 0 } })
        .ok
    ).toBe(true);
    expect(slideXml(e)).toContain('{BEBA8EAE-BF5A-486C-A8C5-ECC9F3942E4B}');
  });

  test('a crop inside an entered group refits the group, as a move does', async () => {
    const group =
      `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="10" name="Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
      `<p:grpSpPr><a:xfrm><a:off x="${IN}" y="${IN}"/><a:ext cx="${4 * IN}" cy="${3 * IN}"/>` +
      `<a:chOff x="${IN}" y="${IN}"/><a:chExt cx="${4 * IN}" cy="${3 * IN}"/></a:xfrm></p:grpSpPr>` +
      pic(2) +
      '</p:grpSp>';
    const e = await open(group);
    const grp = e.session!.layout(0)!.items[0]!;
    const child = grp.kind === 'group' ? grp.children[0]!.element.nodeId : '';
    e.setSelection({ kind: 'shapes', nodeIds: [child], group: grp.element.nodeId });
    e.exec({ type: 'cropImage', crop: { kind: 'rect', left: 25, top: 0, right: 0, bottom: 0 } });
    // The group's bounds follow the cropped picture: one inch narrower on the left.
    const g = e.session!.layout(0)!.items[0]!.element.xfrm;
    expect([g.x, g.cx]).toEqual([2 * IN, 3 * IN]);
    e.exec({ type: 'undo' });
    const back = e.session!.layout(0)!.items[0]!.element.xfrm;
    expect([back.x, back.cx]).toEqual([IN, 4 * IN]);
    // The parent comes from the tree: a host selection without `group` refits too.
    e.setSelection({ kind: 'shapes', nodeIds: [child] });
    e.exec({ type: 'cropImage', crop: { kind: 'rect', left: 25, top: 0, right: 0, bottom: 0 } });
    const again = e.session!.layout(0)!.items[0]!.element.xfrm;
    expect([again.x, again.cx]).toEqual([2 * IN, 3 * IN]);
  });

  test('a crop that leaves nothing visible (a concurrent merge) counts as no crop', async () => {
    // Two participants cropped 60 % off opposite edges at once; the merge keeps both edges.
    const e = await openSelected(pic(2, { crop: '<a:srcRect l="60000" r="60000"/>' }));
    expect(
      e.exec({ type: 'cropImage', crop: { kind: 'rect', left: 10, top: 0, right: 10, bottom: 0 } })
        .ok
    ).toBe(true);
    // The frame was the whole image: 10 % comes off each side of it.
    expect(box(e)).toEqual({ x: 1.4 * IN, y: IN, cx: 3.2 * IN, cy: 3 * IN });
  });

  test('Format Picture reads a merged crop as paint does, so another edge can change', async () => {
    const e = await openSelected(pic(2, { crop: '<a:srcRect l="60000" r="60000"/>' }));
    const picture = e.snapshot().image!;
    // The axis that shows nothing reads as uncropped, as paint and crop mode draw it.
    expect([picture.crop.left, picture.crop.right]).toEqual([0, 0]);
    const initial = picturePropertiesDraftFrom(picture);
    const command = picturePropertiesDraftCommand(
      { ...initial, cropTop: 10 },
      initial,
      picture.crop
    );
    expect(e.exec(command!).ok).toBe(true);
    expect(e.snapshot().image?.crop).toEqual({ left: 0, top: 10, right: 0, bottom: 0 });
  });
});

describe('adjustments and styles', () => {
  test('corrections, recolor, and transparency write blip effects in order', async () => {
    const e = await openSelected();
    const set = (adjustments: object) =>
      e.exec({ type: 'setImageAdjustments', adjustments } as never);
    expect(set({ brightness: 0.2, contrast: -0.4 }).ok).toBe(true);
    expect(set({ transparency: 0.5 }).ok).toBe(true);
    expect(set({ recolor: 'grayscale' }).ok).toBe(true);
    expect(slideXml(e)).toContain(
      '<a:blip r:embed="rIdImg"><a:grayscl/><a:lum bright="20000" contrast="-40000"/><a:alphaModFix amt="50000"/></a:blip>'
    );
    const p = e.snapshot().image!;
    expect([p.brightness, p.contrast, p.transparency, p.recolor]).toEqual([
      0.2,
      -0.4,
      0.5,
      'grayscale',
    ]);
    set({ recolor: 'accent2Dark' });
    expect(e.snapshot().image?.recolor).toBe('accent2Dark');
    set({ brightness: 0 });
    expect(slideXml(e)).toContain('<a:lum contrast="-40000"/>');
    set({ contrast: 0, transparency: 0, recolor: 'none' });
    expect(slideXml(e)).toContain('<a:blip r:embed="rIdImg"/>');
    expect(set({ transparency: 2 }).ok).toBe(false);
    expect(set({ recolor: 'purple' }).ok).toBe(false);
    // The resolved picture carries the effects to paint.
    set({ recolor: 'sepia', transparency: 0.25 });
    const el = e.session!.layout(0)!.items[0].element;
    const effects = el.kind === 'picture' && el.fill.type === 'picture' ? el.fill.effects : [];
    expect(effects.map((x) => x.kind)).toEqual(['duotone', 'alphaModFix']);
  });

  test('a duotone the gallery does not offer shows as other', async () => {
    const e = await openSelected(
      pic(2, { blip: '<a:duotone><a:prstClr val="black"/><a:srgbClr val="FFC000"/></a:duotone>' })
    );
    expect(e.snapshot().image?.recolor).toBe('other');
  });

  test('picture styles set the border, effects, and shape', async () => {
    const e = await openSelected();
    expect(e.exec({ type: 'setImageStyle', style: 'simpleFrameWhite' }).ok).toBe(true);
    const xml = slideXml(e);
    expect(xml).toContain('<a:ln w="88900" cap="sq">');
    expect(xml).toMatch(/<\/a:ln><a:effectLst><a:outerShdw/);
    expect(e.snapshot().image?.style).toBe('simpleFrameWhite');
    e.exec({ type: 'setImageStyle', style: 'softEdgeOval' });
    expect(slideXml(e)).not.toContain('a:ln ');
    expect(e.snapshot().image?.shape).toBe('ellipse');
    expect(e.snapshot().image?.style).toBe('softEdgeOval');
    // A border edit makes the style no longer match.
    e.exec({ type: 'setLineColor', color: '#FF0000' });
    expect(e.snapshot().image?.style).toBeUndefined();
    expect(e.can({ type: 'setImageStyle', style: 'toString' as never }).ok).toBe(false);
  });

  test('a style with no border removes the style reference that would bring a theme line', async () => {
    const STYLE =
      '<p:style><a:lnRef idx="2"><a:schemeClr val="accent1"/></a:lnRef><a:fillRef idx="0"><a:schemeClr val="accent1"/></a:fillRef>' +
      '<a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="tx1"/></a:fontRef></p:style>';
    const e = await openSelected(
      pic(2, { spPr: '<a:ln><a:noFill/></a:ln>' }).replace(
        '</p:spPr></p:pic>',
        `</p:spPr>${STYLE}</p:pic>`
      )
    );
    e.exec({ type: 'setImageStyle', style: 'softEdgeRectangle' });
    expect(slideXml(e)).not.toContain('p:style');
    const item = e.session!.layout(0)!.items[0]!;
    const line = item.kind === 'picture' ? item.element.line : undefined;
    expect(!line || line.fill.type === 'none').toBe(true);
  });
});

describe('alt text', () => {
  test('description and title are written, escaped on save, and read back', async () => {
    const e = await openSelected();
    const text = 'A "red" <box> & more';
    expect(e.exec({ type: 'setAltText', description: text, title: 'Box' }).ok).toBe(true);
    expect(e.snapshot().geometry?.altText).toEqual({
      description: text,
      title: 'Box',
      decorative: false,
    });
    const saved = e.save()!;
    const zipXml = new TextDecoder().decode(saved);
    expect(zipXml).not.toContain('<box>');
    const r = readPptx(saved);
    expect(r.ok).toBe(true);
    const reopened = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    await reopened.load(saved);
    reopened.selectAll();
    expect(reopened.snapshot().geometry?.altText?.description).toBe(text);
    expect(slideXml(reopened)).toContain('descr="A &quot;red&quot; &lt;box&gt; &amp; more"');
  });

  test('decorative removes the description and adds the extension; clearing removes both', async () => {
    const e = await openSelected();
    e.exec({ type: 'setAltText', description: 'Logo' });
    // A description together with the mark is refused: a decorative object has none.
    expect(e.can({ type: 'setAltText', description: 'Logo', decorative: true }).ok).toBe(false);
    e.exec({ type: 'setAltText', decorative: true });
    let xml = slideXml(e);
    expect(xml).not.toContain('descr=');
    expect(xml).toContain('uri="{C183D7F6-B498-43B3-948B-1728B52AA6E4}"');
    expect(xml).toMatch(/<adec:decorative xmlns:adec="[^"]+" val="1"\/>/);
    expect(e.snapshot().geometry?.altText?.decorative).toBe(true);
    // An absent `decorative` keeps the mark; `false` removes it.
    e.exec({ type: 'setAltText', description: '' });
    expect(e.snapshot().geometry?.altText?.decorative).toBe(true);
    e.exec({ type: 'setAltText', description: '', decorative: false });
    xml = slideXml(e);
    expect(xml).not.toContain('decorative');
    expect(xml).not.toContain('a:extLst');
    expect(e.snapshot().geometry?.altText).toEqual({
      description: '',
      title: '',
      decorative: false,
    });
  });

  test('alt text works on shapes; bad text and several shapes are refused', async () => {
    const e = await open(shape({ id: 2, x: IN, y: IN, cx: IN, cy: IN }) + pic(3));
    e.selectAll();
    expect(e.can({ type: 'setAltText', description: 'x' })).toEqual({
      ok: false,
      reason: 'disabled.selectOne',
    });
    select(e, 0);
    expect(e.exec({ type: 'setAltText', description: 'A square' }).ok).toBe(true);
    expect(slideXml(e)).toContain('descr="A square"');
    expect(e.exec({ type: 'setAltText', description: 'bad \u0001 char' })).toMatchObject({
      ok: false,
      reason: 'errors.invalidText',
    });
    expect(e.exec({ type: 'setAltText', description: 'x'.repeat(4001) })).toMatchObject({
      ok: false,
      reason: 'errors.altTextLength',
    });
  });

  test('picture edits are refused while viewing', async () => {
    const e = await openSelected();
    e.setMode('view');
    expect(e.can({ type: 'setAltText', description: 'x' })).toEqual({
      ok: false,
      reason: 'disabled.viewing',
    });
    expect(e.can({ type: 'toggleCrop' }).ok).toBe(false);
  });
});
