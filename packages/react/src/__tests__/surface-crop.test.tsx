import { afterEach, describe, expect, test } from 'bun:test';
import { PNG } from 'pngjs';
import { createPptxEditor, type PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import {
  child,
  nodeIndex,
  xmlPartOf,
  type Element as XmlElement,
} from '@pptx-editor.dev/core/store';
import { buildPptx } from '../../../../scripts/fixtures/builder.ts';

// Crop mode on the slide surface in a DOM: a crop handle drag crops with the image fixed, a
// drag on the image moves it in the frame, each drag is one undo step, and Escape ends it.

const IN = 914_400;
const PX = 9_525;
const R_IMAGE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';

function png(): Uint8Array {
  const img = new PNG({ width: 4, height: 3 });
  img.data.fill(200);
  return new Uint8Array(PNG.sync.write(img));
}

async function setup(
  cNvPicPr = '<p:cNvPicPr/>'
): Promise<{ editor: PptxEditorInstance; host: HTMLDivElement }> {
  const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false, zoom: 1 });
  const r = await editor.load(
    buildPptx({
      slides: [
        {
          layout: 'blank',
          rels: [['rIdImg', R_IMAGE, '../media/image1.png']],
          shapes:
            `<p:pic><p:nvPicPr><p:cNvPr id="2" name="Picture 1"/>${cNvPicPr}<p:nvPr/></p:nvPicPr>` +
            '<p:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>' +
            `<p:spPr><a:xfrm><a:off x="${IN}" y="${IN}"/><a:ext cx="${4 * IN}" cy="${3 * IN}"/></a:xfrm>` +
            '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>',
        },
      ],
      media: [{ name: 'image1.png', bytes: png(), contentType: 'image/png' }],
    })
  );
  expect(r.ok).toBe(true);
  const host = document.createElement('div');
  document.body.append(host);
  editor.attach(host);
  // happy-dom lays nothing out: one CSS pixel is one screen pixel at zoom 1.
  const stage = host.querySelector('.pptx-surface__stage') as HTMLElement;
  stage.getBoundingClientRect = () =>
    ({
      left: 0,
      top: 0,
      right: 1280,
      bottom: 720,
      width: 1280,
      height: 720,
      x: 0,
      y: 0,
    }) as DOMRect;
  const id = editor.session!.layout(0)!.items[0]!.element.nodeId;
  editor.setSelection({ kind: 'shapes', nodeIds: [id] });
  return { editor, host };
}

const at = (x: number, y: number) => ({ clientX: x / PX, clientY: y / PX, bubbles: true });

function drag(from: Element, x0: number, y0: number, x1: number, y1: number): void {
  from.dispatchEvent(new PointerEvent('pointerdown', { ...at(x0, y0), button: 0 }));
  document.dispatchEvent(new PointerEvent('pointermove', at((x0 + x1) / 2, (y0 + y1) / 2)));
  document.dispatchEvent(new PointerEvent('pointermove', at(x1, y1)));
  document.dispatchEvent(new PointerEvent('pointerup', at(x1, y1)));
}

afterEach(() => {
  document.body.replaceChildren();
});

describe('crop mode on the surface', () => {
  test('shows the crop handles in place of the box handles', async () => {
    const { editor, host } = await setup();
    expect(host.querySelectorAll('[data-crop-handle]').length).toBe(0);
    editor.exec({ type: 'toggleCrop' });
    expect(host.querySelectorAll('.pptx-surface__crop-handle[data-crop-handle]').length).toBe(8);
    // The halo under each handle is the wider hit target.
    expect(host.querySelectorAll('.pptx-surface__crop-halo[data-crop-handle]').length).toBe(8);
    expect(host.querySelector('[data-handle="se"]') === null).toBe(true);
    expect(host.querySelector('.pptx-surface__crop-image') !== null).toBe(true);
    // The stored pixels, as paint draws them: the EXIF orientation tag is ignored.
    expect(host.querySelector('.pptx-surface__crop-image')!.getAttribute('style')).toBe(
      'image-orientation: none'
    );
    editor.destroy();
  });

  test('a handle drag crops with the image fixed, as one undo step', async () => {
    const { editor, host } = await setup();
    editor.exec({ type: 'toggleCrop' });
    const left = host.querySelector('[data-crop-handle="w"]')!;
    drag(left, IN, 2.5 * IN, 2 * IN, 2.5 * IN);
    expect(editor.snapshot().image?.crop.left).toBeCloseTo(25, 1);
    const g = editor.snapshot().geometry!;
    expect(g.x).toBeCloseTo(2 * IN, -2);
    expect(g.cx).toBeCloseTo(3 * IN, -2);
    // The handles follow the new frame, and crop mode stays on.
    expect(editor.snapshot().image?.cropping).toBe(true);
    editor.exec({ type: 'undo' });
    expect(editor.snapshot().image?.crop.left).toBe(0);
    expect(editor.snapshot().geometry?.x).toBe(IN);
    editor.destroy();
  });

  test('a drag on the image moves it in the frame; Escape ends crop mode', async () => {
    const { editor, host } = await setup();
    editor.exec({ type: 'toggleCrop' });
    const stage = host.querySelector('.pptx-surface__stage')!;
    drag(stage, 3 * IN, 2.5 * IN, 4 * IN, 2.5 * IN);
    const p = editor.snapshot().image!;
    expect([p.crop.left, p.crop.right]).toEqual([-25, 25]);
    expect(editor.snapshot().geometry?.x).toBe(IN);
    const root = host.querySelector('.pptx-surface')!;
    root.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
    expect(editor.snapshot().image?.cropping).toBe(false);
    expect(host.querySelectorAll('[data-crop-handle]').length).toBe(0);
    editor.destroy();
  });

  test('a drag step the engine refuses says why (lastRejection)', async () => {
    const { editor, host } = await setup('<p:cNvPicPr><a:picLocks noCrop="0"/></p:cNvPicPr>');
    editor.exec({ type: 'toggleCrop' });
    const left = host.querySelector('[data-crop-handle="w"]')!;
    left.dispatchEvent(new PointerEvent('pointerdown', { ...at(IN, 2.5 * IN), button: 0 }));
    // Another participant locks the picture against cropping during the drag.
    const part = editor.session!.info.slides[0]!.partName;
    const picId = editor.session!.layout(0)!.items[0]!.element.nodeId;
    const pic = nodeIndex(xmlPartOf(editor.session!.store.package, part)!).get(picId)!
      .node as XmlElement;
    const locks = child(child(child(pic, 'p:nvPicPr'), 'p:cNvPicPr'), 'a:picLocks')!;
    editor.session!.store.transact([
      { op: 'setAttribute', part, node: locks.id, name: { ns: '', local: 'noCrop' }, value: '1' },
    ]);
    document.dispatchEvent(new PointerEvent('pointermove', at(2 * IN, 2.5 * IN)));
    document.dispatchEvent(new PointerEvent('pointerup', at(2 * IN, 2.5 * IN)));
    expect(editor.snapshot().lastRejection?.reason).toBe('disabled.cropLocked');
    expect(editor.snapshot().image?.crop.left).toBe(0);
    editor.destroy();
  });

  test('a refused transform drag says why and leaves the document unchanged', async () => {
    const { editor, host } = await setup();
    const before = editor.snapshot().revision;
    const handle = host.querySelector('[data-handle="se"]')!;
    // A size beyond `ST_PositiveCoordinate`: every step of the drag is refused.
    handle.dispatchEvent(new PointerEvent('pointerdown', { ...at(5 * IN, 4 * IN), button: 0 }));
    for (const x of [3e13, 4e13])
      document.dispatchEvent(new PointerEvent('pointermove', at(x, 4 * IN)));
    document.dispatchEvent(new PointerEvent('pointerup', at(4e13, 4 * IN)));
    expect(editor.snapshot().lastRejection?.reason).toBe('errors.coordinateRange');
    expect(editor.snapshot().revision).toBe(before);
    expect(editor.snapshot().geometry?.cx).toBe(4 * IN);
    editor.destroy();
  });
});
