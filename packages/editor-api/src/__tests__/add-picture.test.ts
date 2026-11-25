/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/editor-api/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// `ShapeCollection.addPicture`, an SDK extension. It adds a picture (`p:pic`, `CT_Picture`,
// ECMA-376 Part 1, §19.3.1.37) whose `p:blipFill` (§19.3.1.4) embeds a new Image part
// (§15.2.14) through `a:blip/@r:embed` (`CT_Blip`, §20.1.8.13).

import { describe, expect, test } from 'bun:test';
import type { AutomationHost, AutomationOperation } from '@pptx-editor.dev/core/automation';
import { createBrowserAutomationHost, createPptxEditor } from '@pptx-editor.dev/core/editor';
import {
  attr,
  child,
  elements,
  presentationInfo,
  readPptx,
  relationshipsOf,
  xmlPartOf,
  type Element,
  type PptxPackage,
} from '@pptx-editor.dev/core/store';
import { buildPptx, para, run, shape } from '../../../../scripts/fixtures/builder.ts';
import { PptxEditor as BrowserEditor } from '../browser.ts';
import {
  isPptxEditorError,
  type PptxEditorRuntime,
  type RequestContext,
  type ShapeAddOptions,
} from '../index.ts';
import { createRuntime } from '../runtime/runtime.ts';
import { open, refusedAtSync, reopen } from './support.ts';

const EXT_URI = '{8F3A3B5E-1D2C-4E5F-9A0B-17F2C5D3E4A1}';

/** One blank 960 × 540 pt slide with a text box, an unknown slide extension, and a layout. */
function deck(): Uint8Array {
  return buildPptx({
    slides: [
      {
        layout: 'blank',
        ext: `<p:extLst><p:ext uri="${EXT_URI}"><x:keep xmlns:x="urn:fixture:keep" v="1"/></p:ext></p:extLst>`,
        shapes: shape({
          id: 2,
          name: 'Title',
          x: 914_400,
          y: 914_400,
          cx: 914_400,
          cy: 914_400,
          text: para(run('Report')),
        }),
      },
    ],
  });
}

const toBase64 = (b: Uint8Array): string => Buffer.from(b).toString('base64');

/** The PNG signature and an `IHDR` chunk of `width` × `height` pixels. */
function png(width: number, height: number): string {
  const b = new Uint8Array(33);
  b.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0, 0, 0, 13, 0x49, 0x48, 0x44, 0x52]);
  const view = new DataView(b.buffer);
  view.setUint32(16, width);
  view.setUint32(20, height);
  b.set([8, 6, 0, 0, 0], 24);
  return toBase64(b);
}

/** A GIF header of `width` × `height` pixels. */
function gif(width: number, height: number): string {
  const b = new Uint8Array(13);
  b.set([0x47, 0x49, 0x46, 0x38, 0x39, 0x61]);
  const view = new DataView(b.buffer);
  view.setUint16(6, width, true);
  view.setUint16(8, height, true);
  return toBase64(b);
}

/** A JPEG start and a baseline frame header (`SOF0`) of `width` × `height` pixels. */
function jpeg(width: number, height: number): string {
  const b = new Uint8Array(13);
  b.set([0xff, 0xd8, 0xff, 0xc0, 0, 11, 8]);
  const view = new DataView(b.buffer);
  view.setUint16(7, height);
  view.setUint16(9, width);
  b.set([1, 1], 11);
  return toBase64(b);
}

const code = (error: unknown): string | undefined =>
  isPptxEditorError(error) ? error.code : undefined;

function packageOf(bytes: Uint8Array): PptxPackage {
  const read = readPptx(bytes);
  if (!read.ok) throw new Error('the saved presentation does not read back');
  return read.pkg;
}

const SLIDE = '/ppt/slides/slide1.xml';

/** The `p:pic` elements of the first slide. */
function picturesOf(pkg: PptxPackage): Element[] {
  const tree = child(child(xmlPartOf(pkg, SLIDE)!.root, 'p:cSld'), 'p:spTree');
  return elements(tree).filter((e) => e.local === 'pic');
}

interface PictureRead {
  type: string;
  name: string;
  left: number;
  top: number;
  width: number;
  height: number;
  altTextDescription: string;
}

/** Load the shapes of the first slide. */
async function shapesOf(context: RequestContext): Promise<PictureRead[]> {
  const shapes = context.presentation.slides
    .getItemAt(0)
    .shapes.load(
      'items/type,items/name,items/left,items/top,items/width,items/height,items/altTextDescription'
    );
  await context.sync();
  return shapes.items.map((s) => ({
    type: s.type,
    name: s.name,
    left: s.left,
    top: s.top,
    width: s.width,
    height: s.height,
    altTextDescription: s.altTextDescription,
  }));
}

/** Add one picture to the first slide and read it back. */
async function addOne(
  runtime: PptxEditorRuntime,
  base64: string,
  options?: ShapeAddOptions
): Promise<PictureRead> {
  return runtime.run(async (context) => {
    const picture = context.presentation.slides.getItemAt(0).shapes.addPicture(base64, options);
    picture.load('type,name,left,top,width,height,altTextDescription');
    await context.sync();
    return {
      type: picture.type,
      name: picture.name,
      left: picture.left,
      top: picture.top,
      width: picture.width,
      height: picture.height,
      altTextDescription: picture.altTextDescription,
    };
  });
}

describe('ShapeCollection.addPicture', () => {
  test('adds a picture shape that reads back, saves, and reopens', async () => {
    const runtime = await open(deck());
    // 400 × 200 px at 96 pixels per inch is 300 × 150 pt; it fits 60% of the slide as it is.
    const added = await addOne(runtime, png(400, 200));
    expect(added).toEqual({
      type: 'Image',
      name: 'Picture 2',
      left: 330,
      top: 195,
      width: 300,
      height: 150,
      altTextDescription: '',
    });
    const pkg = packageOf(await runtime.save());
    const [pic] = picturesOf(pkg);
    expect(pic).toBeDefined();
    // No alt text by default (`p:cNvPr/@descr`, §19.3.1.12), and the aspect ratio is locked
    // (`a:picLocks/@noChangeAspect`, §20.1.2.2.31).
    const cNvPr = child(child(pic!, 'p:nvPicPr'), 'p:cNvPr')!;
    expect(attr(cNvPr, 'descr')).toBeUndefined();
    const locks = child(child(child(pic!, 'p:nvPicPr'), 'p:cNvPicPr'), 'a:picLocks')!;
    expect(attr(locks, 'noChangeAspect')).toBe('1');
    // The blip embeds a new PNG Image part through a slide relationship.
    const embed = attr(child(child(pic!, 'p:blipFill'), 'a:blip')!, 'r:embed');
    const rel = relationshipsOf(pkg, SLIDE).find((r) => r.id === embed);
    expect(rel?.type).toBe(
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image'
    );
    expect(rel?.target).toMatch(/media\/image\d+\.png$/);
    const again = await reopen(runtime);
    const read = await again.run(shapesOf);
    expect(read.map((s) => s.type)).toEqual(['GeometricShape', 'Image']);
    expect(read[1]).toMatchObject({ left: 330, top: 195, width: 300, height: 150 });
    runtime.dispose();
    again.dispose();
  });

  test('keeps unknown content through open, add, save, and reopen', async () => {
    const runtime = await open(deck());
    await addOne(runtime, png(10, 10));
    const pkg = packageOf(await runtime.save());
    const root = xmlPartOf(pkg, SLIDE)!.root;
    const ext = elements(child(root, 'p:extLst')).find((e) => attr(e, 'uri') === EXT_URI);
    expect(ext).toBeDefined();
    expect(presentationInfo(pkg).slides).toHaveLength(1);
    runtime.dispose();
  });

  test('a large picture keeps its aspect ratio and fits 60% of the slide', async () => {
    const runtime = await open(deck());
    // 4000 × 1000 px: the width limits it to 60% of 960 pt.
    const added = await addOne(runtime, png(4000, 1000));
    expect(added.width).toBeCloseTo(576, 1);
    expect(added.height).toBeCloseTo(144, 1);
    expect(added.left).toBeCloseTo(192, 1);
    expect(added.top).toBeCloseTo(198, 1);
    runtime.dispose();
  });

  test('one dimension keeps the aspect ratio, and a size with no position centers it', async () => {
    const runtime = await open(deck());
    expect(await addOne(runtime, png(400, 200), { width: 200 })).toMatchObject({
      left: 380,
      top: 220,
      width: 200,
      height: 100,
    });
    expect(await addOne(runtime, png(400, 200), { height: 50, left: 10 })).toMatchObject({
      left: 10,
      top: 245,
      width: 100,
      height: 50,
    });
    // A full box is used as given, and a position alone keeps the default size.
    expect(
      await addOne(runtime, png(400, 200), { left: 1, top: 2, width: 30, height: 40 })
    ).toMatchObject({ left: 1, top: 2, width: 30, height: 40 });
    expect(await addOne(runtime, png(400, 200), { left: 5, top: 6 })).toMatchObject({
      left: 5,
      top: 6,
      width: 300,
      height: 150,
    });
    runtime.dispose();
  });

  test('takes a GIF, and alt text set in the same batch saves as p:cNvPr/@descr', async () => {
    const runtime = await open(deck());
    await runtime.run(async (context) => {
      const picture = context.presentation.slides.getItemAt(0).shapes.addPicture(gif(64, 32));
      picture.altTextDescription = 'Company logo';
      await context.sync();
    });
    const pkg = packageOf(await runtime.save());
    const [pic] = picturesOf(pkg);
    expect(attr(child(child(pic!, 'p:nvPicPr'), 'p:cNvPr')!, 'descr')).toBe('Company logo');
    const embed = attr(child(child(pic!, 'p:blipFill'), 'a:blip')!, 'r:embed');
    expect(relationshipsOf(pkg, SLIDE).find((r) => r.id === embed)?.target).toMatch(/\.gif$/);
    runtime.dispose();
  });

  test('takes a JPEG', async () => {
    const runtime = await open(deck());
    expect(await addOne(runtime, jpeg(200, 400))).toMatchObject({ width: 150, height: 300 });
    const pkg = packageOf(await runtime.save());
    const embed = attr(child(child(picturesOf(pkg)[0]!, 'p:blipFill'), 'a:blip')!, 'r:embed');
    expect(relationshipsOf(pkg, SLIDE).find((r) => r.id === embed)?.target).toMatch(/\.jpeg$/);
    runtime.dispose();
  });

  test('bytes over 50 MiB refuse with InvalidArgument before and after the decode', async () => {
    const runtime = await open(deck());
    const before = await runtime.save();
    // A PNG header and padding past 50 MiB: the base64 fits its length cap, the bytes do not.
    const big = new Uint8Array(50 * 1024 * 1024 + 1);
    big.set(Buffer.from(png(10, 10), 'base64'));
    // Base64 text past the length cap refuses before it is decoded.
    const long = 'A'.repeat(Math.ceil((50 * 1024 * 1024) / 3) * 4 * 1.1 + 4);
    for (const value of [toBase64(big), long]) {
      const error = await runtime.run((context) =>
        refusedAtSync(context, () =>
          context.presentation.slides.getItemAt(0).shapes.addPicture(value)
        )
      );
      expect(code(error)).toBe('InvalidArgument');
    }
    expect(await runtime.save()).toEqual(before);
    runtime.dispose();
  });

  test('a computed side out of range refuses with picture-size, not with a value never given', async () => {
    const runtime = await open(deck());
    // 1 × 60000 px at a width of 100 pt makes a height of 6 000 000 pt.
    const error = await runtime.run((context) =>
      refusedAtSync(context, () =>
        context.presentation.slides.getItemAt(0).shapes.addPicture(png(1, 60_000), { width: 100 })
      )
    );
    expect(code(error)).toBe('InvalidArgument');
    const hint = isPptxEditorError(error) ? error.debugInfo.innerError?.message : undefined;
    expect(hint).toContain('Give a width and a height');
    // A width out of range names the width, though the left is computed from it.
    const wide = await runtime.run((context) =>
      refusedAtSync(context, () =>
        context.presentation.slides.getItemAt(0).shapes.addPicture(png(10, 10), { width: 1e9 })
      )
    );
    const named = isPptxEditorError(wide) ? wide.debugInfo.innerError?.message : undefined;
    expect(named).toBe('Use a number of points from 0 to 169 093.');
    runtime.dispose();
  });

  test('adds to a slide layout', async () => {
    const runtime = await open(deck());
    const type = await runtime.run(async (context) => {
      const layout = context.presentation.slides.getItemAt(0).layout;
      const picture = layout.shapes.addPicture(png(20, 20));
      picture.load('type');
      await context.sync();
      return picture.type;
    });
    expect(type).toBe('Image');
    runtime.dispose();
  });

  test('bytes that are no PNG, JPEG, or GIF refuse with InvalidArgument and write nothing', async () => {
    const runtime = await open(deck());
    const before = await runtime.save();
    const svg = toBase64(new TextEncoder().encode('<svg xmlns="http://www.w3.org/2000/svg"/>'));
    const bmp = toBase64(new Uint8Array([0x42, 0x4d, ...new Array<number>(52).fill(0)]));
    // A PNG signature with no readable size.
    const truncated = toBase64(new Uint8Array([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a]));
    for (const value of [svg, bmp, truncated, '', 'not base64 !', 42 as unknown as string]) {
      const error = await runtime.run((context) =>
        refusedAtSync(context, () =>
          context.presentation.slides.getItemAt(0).shapes.addPicture(value)
        )
      );
      expect(code(error)).toBe('InvalidArgument');
    }
    expect(await runtime.save()).toEqual(before);
    runtime.dispose();
  });

  test('a bad box refuses with InvalidArgument and leaves no image part', async () => {
    const runtime = await open(deck());
    const before = await runtime.save();
    const error = await runtime.run((context) =>
      refusedAtSync(context, () =>
        context.presentation.slides.getItemAt(0).shapes.addPicture(png(10, 10), { width: -5 })
      )
    );
    expect(code(error)).toBe('InvalidArgument');
    const after = packageOf(await runtime.save());
    expect(picturesOf(after)).toHaveLength(0);
    expect([...after.parts.keys()].some((p) => p.startsWith('/ppt/media/'))).toBe(false);
    expect(await runtime.save()).toEqual(before);
    runtime.dispose();
  });

  test('the browser runtime adds the same picture, as one undo step', async () => {
    const live = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    expect((await live.load(deck())).ok).toBe(true);
    const runtime = BrowserEditor.createBrowser(live);
    const added = await runtime.run(async (context) => {
      const picture = context.presentation.slides.getItemAt(0).shapes.addPicture(png(400, 200));
      picture.altTextDescription = 'Chart';
      await context.sync();
      return shapesOf(context);
    });
    const server = await open(deck());
    expect(await addOne(server, png(400, 200))).toMatchObject({
      type: added[1]!.type,
      left: added[1]!.left,
      top: added[1]!.top,
      width: added[1]!.width,
      height: added[1]!.height,
    });
    expect(added[1]).toMatchObject({ type: 'Image', altTextDescription: 'Chart' });
    expect(picturesOf(packageOf(live.save()!))).toHaveLength(1);
    expect(live.exec({ type: 'undo' }).ok).toBe(true);
    expect(picturesOf(packageOf(live.save()!))).toHaveLength(0);
    expect(live.snapshot().canUndo).toBe(false);
    server.dispose();
  });

  test('a call sends the lookups and one addShape operation', async () => {
    // The collaboration test (`packages/pro/src/__tests__/yjs-add-picture.test.ts`) sends these
    // operations to two linked editors; this test keeps them equal to what `addPicture` sends.
    const live = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    expect((await live.load(deck())).ok).toBe(true);
    const inner = createBrowserAutomationHost(live);
    const batches: AutomationOperation[][] = [];
    const host: AutomationHost = {
      ...inner,
      execute: (request) => {
        batches.push([...request.operations]);
        return inner.execute(request);
      },
    };
    const base64 = png(8, 8);
    const runtime = createRuntime({ host, save: false });
    for (const options of [undefined, { left: 10, top: 20 }])
      await runtime.run(async (context) => {
        context.presentation.slides.getItemAt(0).shapes.addPicture(base64, options);
        await context.sync();
      });
    const sent = (box: object): AutomationOperation[] => [
      { op: 'getPresentation' },
      { op: 'getSlide', key: { index: 0 } },
      { op: 'addShape', container: { result: 1 }, shape: { kind: 'picture', base64 }, box },
    ];
    expect(batches).toEqual([sent({}), sent({ left: 10, top: 20 })]);
  });

  test('the example of the pictures page runs as written', async () => {
    // docs/site/content/editor-api/pictures.mdx, "Add a picture shape".
    async function addLogo(runtime: PptxEditorRuntime, logoBase64: string) {
      return runtime.run(async (context) => {
        const slide = context.presentation.slides.getItemAt(0);
        const logo = slide.shapes.addPicture(logoBase64, { left: 24, top: 24, width: 120 });
        logo.altTextDescription = 'Company logo';
        logo.load('id');
        await context.sync();
        return logo.id;
      });
    }
    const runtime = await open(deck());
    const id = await addLogo(runtime, png(300, 100));
    const logo = await runtime.run(async (context) => {
      const shape = context.presentation.slides.getItemAt(0).shapes.getItem(id);
      shape.load('type,left,top,width,height,altTextDescription');
      await context.sync();
      return shape.toJSON();
    });
    expect(logo).toMatchObject({
      type: 'Image',
      left: 24,
      top: 24,
      width: 120,
      height: 40,
      altTextDescription: 'Company logo',
    });
    runtime.dispose();
  });
});
