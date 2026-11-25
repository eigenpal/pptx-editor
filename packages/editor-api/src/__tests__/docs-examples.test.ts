/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/editor-api/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// The examples of docs/site/content/editor-api, run as written against a synthetic deck.

import { readFileSync } from 'node:fs';
import { expect, test } from 'bun:test';
import { createCanvas, GlobalFonts, loadImage } from '@napi-rs/canvas';
import { buildPptx } from '../../../../scripts/fixtures/builder.ts';
import { PptxEditor, ShapeZOrder } from '../index.ts';
import { deck, IN, reopen } from './support.ts';

test('the overview example edits and saves', async () => {
  const runtime = await PptxEditor.createServer(deck(), { substitutes: false });
  try {
    await runtime.run(async (context) => {
      const slide = context.presentation.slides.getItemAt(0);
      const box = slide.shapes.addTextBox('Quarterly results', { left: 40, top: 40, width: 400 });
      box.textFrame.textRange.font.bold = true;
      const table = slide.shapes
        .addTable(2, 2, {
          values: [
            ['Region', 'Total'],
            ['North', '42'],
          ],
        })
        .getTable();
      await context.sync();

      const cell = table.getCellOrNullObject(1, 1);
      cell.horizontalAlignment = 'Right';
      cell.fill.setSolidColor('#FFF2CC');
      await context.sync();
    });
    const again = await reopen(runtime);
    await again.run(async (context) => {
      const shapes = context.presentation.slides.getItemAt(0).shapes.load('items/name,items/type');
      await context.sync();
      expect(shapes.items.map((s) => s.type)).toContain('Table');
    });
  } finally {
    runtime.dispose();
  }
});

test('the compatibility example renames shapes, and enums take both spellings', async () => {
  const runtime = await PptxEditor.createServer(deck(), { substitutes: false });
  await runtime.run(async (context) => {
    const shapes = context.presentation.slides.getItemAt(0).shapes;
    shapes.load('items/name');
    await context.sync();
    for (const shape of shapes.items) shape.name = shape.name.toUpperCase();
    shapes.items[0]!.setZOrder(ShapeZOrder.bringToFront);
    shapes.items[1]!.setZOrder('BringToFront');
    await context.sync();
    shapes.load('items/name');
    await context.sync();
    expect(shapes.items.map((s) => s.name)).toEqual(['BOX', 'OTHER']);
  });
  runtime.dispose();
});

test('the common mistakes examples behave as the docs say', async () => {
  const runtime = await PptxEditor.createServer(deck(), { substitutes: false });
  try {
    await runtime.run(async (context) => {
      const slide = context.presentation.slides.getItemAt(0);
      // A read with no load.
      const box = slide.shapes.getItem('2');
      expect(() => box.name).toThrow();
      box.load('name');
      await context.sync();
      expect(box.name).toBe('Box');
      // A sync inside a loop: one sync for all the writes.
      const shapes = slide.shapes.load('items/name');
      await context.sync();
      for (const shape of shapes.items) shape.name = shape.name.trim();
      await context.sync();
      // A retry after a refusal: the same proxy works.
      let refused = false;
      try {
        box.fill.setSolidColor('notacolor');
        await context.sync();
      } catch {
        refused = true;
        box.fill.setSolidColor('#1F4E79');
        await context.sync();
      }
      expect(refused).toBe(true);
      const fill = box.fill.load('foregroundColor');
      await context.sync();
      expect(fill.foregroundColor).toBe('#1F4E79');
      // Bindings are in memory only.
      context.presentation.bindings.add(box, 'Shape', 'kpi');
      await context.sync();
    });
    const again = await reopen(runtime);
    await again.run(async (context) => {
      const all = context.presentation.bindings.load('items');
      await context.sync();
      expect(all.items).toEqual([]);
    });
  } finally {
    runtime.dispose();
  }
});

/**
 * The `rasterize` example of docs/site/content/editor-api/runtime.mdx, as written. Keep the two
 * equal: the next test compares them.
 */
async function openWithImages(bytes: Uint8Array) {
  return PptxEditor.createServer(bytes, {
    rasterize: async ({ svg, width, height, fonts }) => {
      // Register the faces for this picture only, and remove them before returning.
      const keys = fonts.map((font) => GlobalFonts.register(Buffer.from(font.bytes), font.family));
      try {
        const canvas = createCanvas(width, height);
        canvas.getContext('2d').drawImage(await loadImage(Buffer.from(svg)), 0, 0, width, height);
        return new Uint8Array(canvas.toBuffer('image/png'));
      } finally {
        for (const key of keys) if (key) GlobalFonts.remove(key);
      }
    },
  });
}

test('the rasterize example of the test equals the example of runtime.mdx', () => {
  // The lines from `return PptxEditor.createServer(` to the end of its call, without indentation.
  const body = (source: string): string[] => {
    const lines = source.split('\n').map((line) => line.trim());
    const start = lines.indexOf('return PptxEditor.createServer(bytes, {');
    const end = lines.indexOf('});', start);
    expect(start).toBeGreaterThan(-1);
    expect(end).toBeGreaterThan(start);
    return lines.slice(start, end + 1);
  };
  const guide = readFileSync(
    new URL('../../../../docs/site/content/editor-api/runtime.mdx', import.meta.url),
    'utf8'
  );
  const own = readFileSync(new URL(import.meta.url), 'utf8');
  expect(body(own)).toEqual(body(guide));
  const canvasImport = "import { createCanvas, GlobalFonts, loadImage } from '@napi-rs/canvas';";
  // A whole line: the literal above is part of a longer line, so it does not match.
  expect(own.split('\n')).toContain(canvasImport);
  expect(guide.split('\n')).toContain(canvasImport);
  // The minimum release that the guide states is the floor of the optional peer range.
  const manifest = JSON.parse(
    readFileSync(new URL('../../package.json', import.meta.url), 'utf8')
  ) as { peerDependencies: Record<string, string> };
  const floor = manifest.peerDependencies['@napi-rs/canvas']!.replace(/^\^/, '');
  expect(guide).toContain(`Install \`@napi-rs/canvas\` ${floor} or a later 1.x release`);
  expect(manifest.peerDependencies['@napi-rs/canvas']).toBe(`^${floor}`);
});

test('the runtime example draws the embedded pictures of a slide', async () => {
  // A red PNG in a picture (`p:pic`, ECMA-376 Part 1, 19.3.1.37; its `a:blip`, 20.1.8.13) two
  // inches by one inch at one inch from the top left corner. The engine paints it as a data URL
  // in an SVG `<image>`. `loadImage` of `@napi-rs/canvas` draws it from 1.0.12 only: a probe
  // that draws a red data URL `<image>` over a blue SVG background leaves the pixel blue in
  // 0.1.100 and 1.0.0 to 1.0.11, and draws it red in 1.0.12 and later. The lockfile installs
  // a later release, so this test checks the engine output, not the floor itself.
  const red = createCanvas(8, 8);
  const pen = red.getContext('2d');
  pen.fillStyle = '#FF0000';
  pen.fillRect(0, 0, 8, 8);
  const bytes = buildPptx({
    slides: [
      {
        layout: 'blank',
        rels: [
          [
            'rIdImg',
            'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
            '../media/image1.png',
          ],
        ],
        shapes:
          '<p:pic><p:nvPicPr><p:cNvPr id="2" name="Red"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>' +
          '<p:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>' +
          `<p:spPr><a:xfrm><a:off x="${IN}" y="${IN}"/><a:ext cx="${2 * IN}" cy="${IN}"/></a:xfrm>` +
          '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>',
      },
    ],
    media: [
      {
        name: 'image1.png',
        bytes: new Uint8Array(red.toBuffer('image/png')),
        contentType: 'image/png',
      },
    ],
  });
  const runtime = await openWithImages(bytes);
  try {
    const png = await runtime.run(async (context) => {
      const image = context.presentation.slides.getItemAt(0).getImageAsBase64({ width: 480 });
      await context.sync();
      return image.value;
    });
    const picture = await loadImage(Buffer.from(png, 'base64'));
    expect([picture.width, picture.height]).toEqual([480, 270]);
    const canvas = createCanvas(480, 270);
    const ctx = canvas.getContext('2d');
    ctx.drawImage(picture, 0, 0);
    // 480 pixels across a 13.33-inch slide is 36 pixels per inch: the picture covers x 36-108
    // and y 36-72. Its center is red, and the slide background outside it is white.
    expect([...ctx.getImageData(72, 54, 1, 1).data]).toEqual([255, 0, 0, 255]);
    expect([...ctx.getImageData(200, 200, 1, 1).data]).toEqual([255, 255, 255, 255]);
  } finally {
    runtime.dispose();
  }
});
