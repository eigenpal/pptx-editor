import { afterEach, describe, expect, test } from 'bun:test';
import { PNG } from 'pngjs';
import { buildPptx } from '../../../../../scripts/fixtures/builder.ts';
import {
  captureImageTarget,
  createPptxEditor,
  insertImageFile,
  type PptxEditorInstance,
} from '../index.ts';

// Change Picture from a file (`insertImageFile`): the browser decode reads the stored pixels
// (not the EXIF orientation), and a selection that moves while the file loads refuses the change.

const IN = 914_400;
const R_IMAGE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image';
const g = globalThis as { createImageBitmap?: unknown };
const realBitmap = g.createImageBitmap;
afterEach(() => {
  g.createImageBitmap = realBitmap;
});

function png(w: number, h: number): Uint8Array {
  const img = new PNG({ width: w, height: h });
  img.data.fill(120);
  return new Uint8Array(PNG.sync.write(img));
}

const pic = (id: number, y: number): string =>
  `<p:pic><p:nvPicPr><p:cNvPr id="${id}" name="Picture ${id}"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>` +
  '<p:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>' +
  `<p:spPr><a:xfrm><a:off x="${IN}" y="${y}"/><a:ext cx="${4 * IN}" cy="${3 * IN}"/></a:xfrm>` +
  '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>';

async function open(): Promise<PptxEditorInstance> {
  const e = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  const bytes = buildPptx({
    slides: [
      {
        layout: 'blank',
        rels: [['rIdImg', R_IMAGE, '../media/image1.png']],
        shapes: pic(2, IN) + pic(3, 4 * IN),
      },
    ],
    media: [{ name: 'image1.png', bytes: png(4, 3), contentType: 'image/png' }],
  });
  expect((await e.load(bytes)).ok).toBe(true);
  return e;
}

const nodeAt = (e: PptxEditorInstance, i: number): string =>
  e.session!.layout(0)!.items[i]!.element.nodeId;

describe('Change Picture from a file', () => {
  test('a picker keeps its launch target before the file starts loading', async () => {
    const e = await open();
    e.setSelection({ kind: 'shapes', nodeIds: [nodeAt(e, 0)] });
    const current = captureImageTarget(e);
    e.setSelection({ kind: 'shapes', nodeIds: [nodeAt(e, 1)] });
    const before = e.save();
    const file = new File([png(30, 40).slice().buffer as ArrayBuffer], 'tall.png', {
      type: 'image/png',
    });
    await insertImageFile(e, file, 'replaceImage', current);
    expect(e.save()).toEqual(before);
    expect(e.snapshot().lastRejection?.reason).toBe('disabled.selectionChanged');
    e.destroy();
  });

  test('a replaced document invalidates a pending picker despite matching node identities', async () => {
    const e = await open();
    const source = e.save()!;
    e.setSelection({ kind: 'shapes', nodeIds: [nodeAt(e, 0)] });
    const current = captureImageTarget(e);
    expect((await e.load(source)).ok).toBe(true);
    e.setSelection({ kind: 'shapes', nodeIds: [nodeAt(e, 0)] });
    const before = e.save();
    const file = new File([png(30, 40).slice().buffer as ArrayBuffer], 'tall.png', {
      type: 'image/png',
    });
    await insertImageFile(e, file, 'replaceImage', current);
    expect(e.save()).toEqual(before);
    expect(e.snapshot().lastRejection?.reason).toBe('disabled.selectionChanged');
    e.destroy();
  });

  test('a file read failure reports a refusal without an unhandled rejection', async () => {
    const e = await open();
    e.setSelection({ kind: 'shapes', nodeIds: [nodeAt(e, 0)] });
    const before = e.save();
    const file = {
      arrayBuffer: async () => {
        throw new Error('read failed');
      },
    } as unknown as File;
    await insertImageFile(e, file, 'replaceImage');
    expect(e.save()).toEqual(before);
    expect(e.snapshot().lastRejection?.reason).toBe('errors.unsupported');
    e.destroy();
  });
  test('decodes the stored pixels, not the EXIF orientation', async () => {
    const e = await open();
    e.setSelection({ kind: 'shapes', nodeIds: [nodeAt(e, 0)] });
    const options: unknown[] = [];
    g.createImageBitmap = async (_file: unknown, opts?: unknown) => {
      options.push(opts);
      return { width: 30, height: 40, close() {} };
    };
    const file = new File([png(30, 40).slice().buffer as ArrayBuffer], 'tall.png', {
      type: 'image/png',
    });
    await insertImageFile(e, file, 'replaceImage');
    expect(options).toEqual([{ imageOrientation: 'none' }]);
    // A 3:4 image fits the 4×3 in frame: 2.25×3 in.
    const geo = e.snapshot().geometry!;
    expect([geo.cx, geo.cy]).toEqual([2.25 * IN, 3 * IN]);
  });

  test('a selection that moves while the file loads refuses the change', async () => {
    const e = await open();
    e.setSelection({ kind: 'shapes', nodeIds: [nodeAt(e, 0)] });
    const revision = e.snapshot().revision;
    const file = new File([png(30, 40).slice().buffer as ArrayBuffer], 'tall.png', {
      type: 'image/png',
    });
    const pending = insertImageFile(e, file, 'replaceImage');
    // The user selects the other picture before the file is read.
    e.setSelection({ kind: 'shapes', nodeIds: [nodeAt(e, 1)] });
    await pending;
    expect(e.snapshot().revision).toBe(revision);
    expect(e.snapshot().lastRejection?.reason).toBe('disabled.selectionChanged');
  });
});
