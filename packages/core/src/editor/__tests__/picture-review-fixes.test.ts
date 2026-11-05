import { describe, expect, test } from 'bun:test';
import { PNG } from 'pngjs';
import { buildPptx } from '../../../../../scripts/fixtures/builder.ts';
import { serializeXmlPart, xmlPartOf } from '../../store/index.ts';
import { imageSlotState } from '../chrome-picture.ts';
import {
  createPptxEditor,
  editorCommandFor,
  type ChromeActionHost,
  type PptxEditorInstance,
} from '../index.ts';

// Picture command rules: `a:picLocks` (noResize, noRot, noChangeShapeType), the Alt Text value
// path without a host, Crop > aspect bounds, group refits, the decorative mark, media posters,
// recolor order, and refusal details.

const IN = 914_400;
const R_IMAGE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';

function png(w: number, h: number): Uint8Array {
  const img = new PNG({ width: w, height: h });
  img.data.fill(200);
  return new Uint8Array(PNG.sync.write(img));
}

const pic = (
  o: { locks?: string; nvPr?: string; blip?: string; cNvPr?: string; cx?: number } = {}
) =>
  `<p:pic><p:nvPicPr><p:cNvPr id="2" name="Picture 2">${o.cNvPr ?? ''}</p:cNvPr>` +
  `<p:cNvPicPr><a:picLocks ${o.locks ?? ''}/></p:cNvPicPr><p:nvPr>${o.nvPr ?? ''}</p:nvPr></p:nvPicPr>` +
  `<p:blipFill><a:blip r:embed="rIdImg">${o.blip ?? ''}</a:blip><a:stretch><a:fillRect/></a:stretch></p:blipFill>` +
  `<p:spPr><a:xfrm><a:off x="${IN}" y="${IN}"/><a:ext cx="${o.cx ?? 4 * IN}" cy="${3 * IN}"/></a:xfrm>` +
  '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>';

async function open(shapes: string, select = true): Promise<PptxEditorInstance> {
  const e = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  const bytes = buildPptx({
    slides: [{ layout: 'blank', rels: [['rIdImg', R_IMAGE, '../media/image1.png']], shapes }],
    media: [{ name: 'image1.png', bytes: png(128, 96), contentType: 'image/png' }],
  });
  expect((await e.load(bytes)).ok).toBe(true);
  if (select) e.selectAll();
  return e;
}

const slideXml = (e: PptxEditorInstance): string =>
  serializeXmlPart(xmlPartOf(e.session!.store.package, e.session!.info.slides[0].partName)!);

describe('picture locks', () => {
  test('noResize refuses a new size and Reset Picture & Size; the state says canResize', async () => {
    const e = await open(pic({ locks: 'noResize="1"' }));
    expect(e.snapshot().image?.canResize).toBe(false);
    expect(e.can({ type: 'setImageProperties', widthEmu: 2 * IN })).toEqual({
      ok: false,
      reason: 'disabled.sizeLocked',
    });
    expect(e.can({ type: 'resetImage', size: true })).toEqual({
      ok: false,
      reason: 'disabled.sizeLocked',
    });
    // The current size is no change, and Reset Picture without size still runs.
    expect(e.can({ type: 'setImageProperties', widthEmu: 4 * IN }).ok).toBe(true);
    expect(e.can({ type: 'resetImage' }).ok).toBe(true);
    expect(e.can({ type: 'setImageProperties', rotationDegrees: 30 }).ok).toBe(true);
    const free = await open(pic());
    expect(free.snapshot().image?.canResize).toBe(true);
  });

  test('noRot refuses a new rotation; Format Picture still opens', async () => {
    const e = await open(pic({ locks: 'noRot="1"' }));
    expect(e.can({ type: 'setImageProperties', rotationDegrees: 30 })).toEqual({
      ok: false,
      reason: 'disabled.rotationLocked',
    });
    expect(e.can({ type: 'setImageProperties', widthEmu: 2 * IN }).ok).toBe(true);
    expect(e.snapshot().image && e.can({ type: 'setImageProperties', rotationDegrees: 0 }).ok).toBe(
      true
    );
  });

  test('noChangeShapeType refuses Crop to Shape', async () => {
    const e = await open(pic({ locks: 'noChangeShapeType="1"' }));
    expect(e.can({ type: 'cropToShape', preset: 'ellipse' })).toEqual({
      ok: false,
      reason: 'disabled.shapeLocked',
    });
    // The shape it already has is no change.
    expect(e.can({ type: 'cropToShape', preset: 'rect' }).ok).toBe(true);
  });

  test('noChangeShapeType also refuses a picture style with another shape and Reset Picture', async () => {
    const e = await open(pic({ locks: 'noChangeShapeType="1"' }));
    expect(e.can({ type: 'setImageStyle', style: 'softEdgeOval' })).toEqual({
      ok: false,
      reason: 'disabled.shapeLocked',
    });
    // A rectangle style keeps the shape.
    expect(e.can({ type: 'setImageStyle', style: 'simpleFrameWhite' }).ok).toBe(true);
    const oval = await open(
      pic({ locks: 'noChangeShapeType="1"' }).replace('prst="rect"', 'prst="ellipse"')
    );
    expect(oval.can({ type: 'resetImage' })).toEqual({ ok: false, reason: 'disabled.shapeLocked' });
  });

  test('Change Picture honors noResize and noCrop', async () => {
    const wide = png(300, 50);
    const sized = await open(pic({ locks: 'noResize="1"' }));
    expect(sized.can({ type: 'replaceImage', bytes: wide, mime: 'image/png' })).toEqual({
      ok: false,
      reason: 'disabled.sizeLocked',
    });
    // An image with the frame's aspect keeps the frame and passes.
    expect(sized.can({ type: 'replaceImage', bytes: png(160, 120), mime: 'image/png' }).ok).toBe(
      true
    );
    // A size within 1 EMU of the frame is rounding: the frame stays exactly.
    const odd = await open(
      pic({ locks: 'noResize="1"', cx: 16335375 }).replace(`cy="${3 * IN}"`, 'cy="7000875"')
    );
    expect(odd.can({ type: 'replaceImage', bytes: png(343, 147), mime: 'image/png' }).ok).toBe(
      true
    );
    // Change Picture stays on for a locked frame that is not square.
    expect(imageSlotState(sized, 'image.change')?.enabled).toBe(true);
    const cropped = await open(
      pic({ locks: 'noCrop="1"' }).replace('</a:blip>', '</a:blip><a:srcRect l="10000"/>')
    );
    expect(cropped.can({ type: 'replaceImage', bytes: wide, mime: 'image/png' })).toEqual({
      ok: false,
      reason: 'disabled.cropLocked',
    });
  });
});

describe('the Alt Text value path', () => {
  test('a string runs setAltText without a host and while the popup is off', async () => {
    const e = await open(pic());
    const noRoot = editorCommandFor(e, 'image.altText', null, 'disabled.noRoot');
    expect(noRoot.state.isEnabled).toBe(true);
    expect(noRoot.execute('A chart')).toBe(true);
    expect(e.snapshot().geometry?.altText?.description).toBe('A chart');
    // Without a value it opens the dialog, which needs the host: nothing happens.
    expect(noRoot.execute()).toBe(false);
    const opened: string[] = [];
    const host: ChromeActionHost = {
      openDialog: (key) => void opened.push(key),
      startSlideshow: () => {},
      openFind: () => {},
      openLink: () => {},
      pickImage: () => {},
      unavailable: (slot) => (slot === 'image.altText' ? 'dialog.unavailable' : null),
    };
    const off = editorCommandFor(e, 'image.altText', host, 'disabled.noRoot');
    expect(off.execute('Sales by region')).toBe(true);
    expect(e.snapshot().geometry?.altText?.description).toBe('Sales by region');
    expect(off.execute()).toBe(false);
    expect(opened).toEqual([]);
  });
});

describe('crop rules', () => {
  test('an aspect crop that leaves almost nothing visible is refused', async () => {
    // A 40×3 in picture cropped to 1:99 would keep less than a thousandth of its width.
    const e = await open(pic({ cx: 40 * IN }));
    expect(e.can({ type: 'cropImage', crop: { kind: 'aspect', width: 1, height: 99 } })).toEqual({
      ok: false,
      reason: 'errors.unsupported',
    });
    expect(e.can({ type: 'cropImage', crop: { kind: 'aspect', width: 1, height: 1 } }).ok).toBe(
      true
    );
  });

  test('Fill and a crop that keeps the frame leave the group bounds alone', async () => {
    // The group's bounds differ from its child's on purpose: a refit would change them.
    const group =
      `<p:grpSp><p:nvGrpSpPr><p:cNvPr id="10" name="Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>` +
      `<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${6 * IN}" cy="${5 * IN}"/>` +
      `<a:chOff x="0" y="0"/><a:chExt cx="${6 * IN}" cy="${5 * IN}"/></a:xfrm></p:grpSpPr>` +
      pic() +
      '</p:grpSp>';
    const e = await open(group, false);
    const grp = e.session!.layout(0)!.items[0]!;
    const child = grp.kind === 'group' ? grp.children[0]!.element.nodeId : '';
    e.setSelection({ kind: 'shapes', nodeIds: [child], group: grp.element.nodeId });
    const bounds = () => {
      const g = e.session!.layout(0)!.items[0]!.element.xfrm;
      return [g.x, g.y, g.cx, g.cy];
    };
    expect(e.exec({ type: 'cropImage', crop: { kind: 'fill' } }).ok).toBe(true);
    expect(
      e.exec({
        type: 'cropImage',
        crop: { kind: 'rect', left: 10, top: 0, right: 0, bottom: 0, keepFrame: true },
      }).ok
    ).toBe(true);
    expect(bounds()).toEqual([0, 0, 6 * IN, 5 * IN]);
    // A crop that moves the frame refits.
    e.exec({ type: 'cropImage', crop: { kind: 'rect', left: 25, top: 0, right: 0, bottom: 0 } });
    expect(bounds()).not.toEqual([0, 0, 6 * IN, 5 * IN]);
  });
});

describe('alt text and refusal details', () => {
  const NOT_DECORATIVE =
    '<a:extLst><a:ext uri="{C183D7F6-B498-43B3-948B-1728B52AA6E4}">' +
    '<adec:decorative xmlns:adec="http://schemas.microsoft.com/office/drawing/2017/decorative" val="0"/>' +
    '</a:ext></a:extLst>';

  test('a title edit keeps an explicit "not decorative" mark', async () => {
    const e = await open(pic({ cNvPr: NOT_DECORATIVE }));
    expect(e.exec({ type: 'setAltText', title: 'Sales' }).ok).toBe(true);
    expect(e.exec({ type: 'setAltText', description: 'A chart', decorative: false }).ok).toBe(true);
    expect(slideXml(e)).toContain('val="0"');
    expect(e.snapshot().geometry?.altText).toEqual({
      description: 'A chart',
      title: 'Sales',
      decorative: false,
    });
  });

  test('setImageProperties points alt text to setAltText; an empty command names its fields', async () => {
    const e = await open(pic());
    const r = e.exec({ type: 'setImageProperties', title: 'x' } as never);
    expect(r).toMatchObject({
      ok: false,
      reason: 'errors.unsupported',
      detail: 'title: use setAltText',
    });
    const empty = e.exec({ type: 'setImageProperties' });
    expect(empty.ok === false && empty.detail).toContain('widthEmu');
  });
});

describe('media and recolor', () => {
  test('every media poster refuses Change Picture', async () => {
    const posters = [
      '<a:audioCd><a:st track="1"/><a:end track="2"/></a:audioCd>',
      '<a:wavAudioFile r:embed="rIdImg"/>',
      '<a:quickTimeFile r:link="rIdImg"/>',
      '<a:videoFile r:link="rIdImg"/>',
      '<a:audioFile r:link="rIdImg"/>',
    ];
    for (const nvPr of posters) {
      const e = await open(pic({ nvPr }));
      expect(
        e.can({ type: 'replaceImage', bytes: png(4, 3), mime: 'image/png', width: 4, height: 3 })
      ).toEqual({ ok: false, reason: 'disabled.unsupportedElement' });
    }
  });

  test('a foreign element named like a recolor effect does not move the corrections', async () => {
    const e = await open(pic({ blip: '<x:grayscl xmlns:x="urn:example"/>' }));
    expect(e.exec({ type: 'setImageAdjustments', adjustments: { brightness: 0.2 } }).ok).toBe(true);
    const xml = slideXml(e);
    expect(xml.indexOf('<a:lum')).toBeLessThan(xml.indexOf('grayscl'));
  });
});
