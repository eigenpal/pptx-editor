/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/editor-api/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// The object model end to end on a server runtime, with a save and reopen after the edits.

import { describe, expect, test } from 'bun:test';
import { presetNames } from '@pptx-editor.dev/core/layout';
import { geometryOf } from '../model/geometry.ts';
import { isPptxEditorError, type RequestContext } from '../index.ts';
import { readZip } from '@pptx-editor.dev/core/store';
import { open, rejection, reopen } from './support.ts';

const code = (error: unknown): string | undefined =>
  isPptxEditorError(error) ? error.code : String(error);

const box = (context: RequestContext) =>
  context.presentation.slides.getItemAt(0).shapes.getItem('2');

describe('presentation', () => {
  test('reads its id, title, and page setup, and writes properties', async () => {
    const runtime = await open();
    await runtime.run(async (context) => {
      const p = context.presentation;
      p.properties.title = 'Plan';
      p.properties.author = 'Ops';
      p.properties.revisionNumber = 7;
      p.pageSetup.slideWidth = 720;
      await context.sync();
      p.load('id,title');
      p.pageSetup.load();
      const props = p.properties.load('title,author,revisionNumber');
      await context.sync();
      expect(p.title).toBe('Plan');
      expect(p.id.length).toBeGreaterThan(0);
      expect(p.pageSetup.slideWidth).toBe(720);
      expect([props.author, props.revisionNumber]).toEqual(['Ops', 7]);
    });
    const again = await reopen(runtime);
    await again.run(async (context) => {
      const props = context.presentation.properties.load('title');
      await context.sync();
      expect(props.title).toBe('Plan');
    });
  });

  test('tags on the presentation, a slide, and a shape', async () => {
    const runtime = await open();
    await runtime.run(async (context) => {
      context.presentation.tags.add('deck', 'q3');
      context.presentation.slides.getItemAt(0).tags.add('status', 'draft');
      box(context).tags.add('role', 'title');
      await context.sync();
    });
    const again = await reopen(runtime);
    await again.run(async (context) => {
      const deckTags = context.presentation.tags.load('items');
      const role = box(context).tags.getItem('ROLE').load('value');
      const missing = box(context).tags.getItemOrNullObject('nope').load('value');
      const count = context.presentation.slides.getItemAt(0).tags.getCount();
      await context.sync();
      expect(deckTags.items.map((t) => [t.key, t.value])).toEqual([['DECK', 'q3']]);
      expect(role.value).toBe('title');
      expect(missing.isNullObject).toBe(true);
      expect(count.value).toBe(1);
      role.value = 'body';
      box(context).tags.delete('role');
      await context.sync().catch(() => undefined);
    });
    const error = await rejection(
      again.run(async (context) => {
        box(context).tags.delete('never');
        await context.sync();
      })
    );
    expect(code(error)).toBe('ItemNotFound');
  });
});

describe('slides', () => {
  test('add by layout id, move, apply a layout, and delete', async () => {
    const runtime = await open();
    await runtime.run(async (context) => {
      const slides = context.presentation.slides;
      const master = context.presentation.slideMasters.getItemAt(0);
      const layouts = master.layouts.load('items/id,items/name,items/type');
      master.load('id,name');
      await context.sync();
      expect(layouts.items.length).toBeGreaterThan(1);
      const title = layouts.items[1]!;
      slides.add({ layoutId: title.id, slideMasterId: master.id });
      await context.sync();
      const all = slides.load('items/id');
      await context.sync();
      expect(all.items).toHaveLength(3);
      const added = all.items[2]!;
      const layout = added.layout.load('id');
      await context.sync();
      expect(layout.id).toBe(title.id);
      added.moveTo(0);
      added.applyLayout(layouts.items[0]!);
      await context.sync();
      added.load('index');
      added.layout.load('type');
      await context.sync();
      expect(added.index).toBe(0);
      slides.getItemAt(2).delete();
      await context.sync();
      expect(slides.getCount()).toBeDefined();
      const count = slides.getCount();
      await context.sync();
      expect(count.value).toBe(2);
    });
  });

  test('a slide and its layout and master in one sync', async () => {
    const runtime = await open();
    await runtime.run(async (context) => {
      const slide = context.presentation.slides.getItemAt(0);
      const layout = slide.layout.load('name,type');
      const master = slide.slideMaster.load('name');
      const layoutShapes = slide.layout.shapes.getCount();
      await context.sync();
      expect(layout.type).toBe('Blank');
      expect(typeof master.name).toBe('string');
      expect(layoutShapes.value).toBeGreaterThanOrEqual(0);
    });
  });

  test('export a slide and a set of slides as presentations', async () => {
    const runtime = await open();
    const [one, both] = await runtime.run(async (context) => {
      const a = context.presentation.slides.getItemAt(1).exportAsBase64();
      const slides = context.presentation.slides.load('items');
      await context.sync();
      const b = context.presentation.slides.exportAsBase64Presentation(slides.items);
      await context.sync();
      return [a.value, b.value];
    });
    for (const [base64, count] of [
      [one, 1],
      [both, 2],
    ] as const) {
      const bytes = Uint8Array.from(atob(base64), (c) => c.charCodeAt(0));
      const exported = await open(bytes);
      await exported.run(async (context) => {
        const n = context.presentation.slides.getCount();
        await context.sync();
        expect(n.value).toBe(count);
      });
    }
  });
});

describe('shapes', () => {
  test('read and write geometry, names, alternative text, and visibility', async () => {
    const runtime = await open();
    await runtime.run(async (context) => {
      const shape = box(context);
      shape.left = 100;
      shape.width = 250;
      shape.rotation = 30;
      shape.altTextTitle = 'Revenue';
      shape.altTextDescription = 'A blue box';
      shape.visible = false;
      await context.sync();
      shape.load();
      await context.sync();
      expect(shape.toJSON()).toMatchObject({
        id: '2',
        name: 'Box',
        type: 'GeometricShape',
        left: 100,
        top: 72,
        width: 250,
        height: 72,
        rotation: 30,
        visible: false,
        zOrderPosition: 0,
        level: 0,
        altTextTitle: 'Revenue',
        altTextDescription: 'A blue box',
        isDecorative: false,
        creationId: null,
      });
    });
  });

  test('add geometric shapes and lines, and order them', async () => {
    const runtime = await open();
    await runtime.run(async (context) => {
      const shapes = context.presentation.slides.getItemAt(1).shapes;
      const star = shapes.addGeometricShape('Star5', { left: 10, top: 10, width: 50, height: 50 });
      const rect = shapes.addGeometricShape('Rectangle');
      const line = shapes.addLine('Elbow', { left: 0, top: 0, width: 100, height: 20 });
      star.setZOrder('BringToFront');
      rect.setZOrder('SendToBack');
      star.load('zOrderPosition,type');
      line.load('type');
      await context.sync();
      expect(star.zOrderPosition).toBe(2);
      expect(star.type).toBe('GeometricShape');
      expect(line.type).toBe('Line');
    });
    const error = await rejection(
      runtime.run(async (context) => {
        context.presentation.slides.getItemAt(1).shapes.addGeometricShape('Blob' as never);
        await context.sync();
      })
    );
    expect(code(error)).toBe('InvalidArgument');
  });

  test('every geometric shape type names an ECMA-376 preset', () => {
    const presets = new Set(presetNames());
    const values = [
      'Rectangle',
      'RoundRectangle',
      'WedgeRRectCallout',
      'DiagonalStripe',
      'LineInverse',
      'FlowChartOffpageConnector',
      'ChartPlus',
    ];
    for (const value of values) expect(presets.has(geometryOf(value, 'test'))).toBe(true);
  });

  test('fill and outline', async () => {
    const runtime = await open();
    await runtime.run(async (context) => {
      const shape = box(context);
      shape.fill.setSolidColor('#FF0000');
      shape.fill.transparency = 0.25;
      shape.lineFormat.color = '#00FF00';
      shape.lineFormat.weight = 3;
      shape.lineFormat.dashStyle = 'LongDash';
      shape.lineFormat.style = 'ThickThin';
      await context.sync();
      const fill = shape.fill.load();
      const line = shape.lineFormat.load();
      await context.sync();
      expect(fill.toJSON()).toEqual({
        foregroundColor: '#FF0000',
        transparency: 0.25,
        type: 'Solid',
      });
      expect(line.toJSON()).toMatchObject({
        color: '#00FF00',
        weight: 3,
        dashStyle: 'LongDash',
        style: 'ThickThin',
        visible: true,
      });
      shape.fill.clear();
      shape.lineFormat.visible = false;
      await context.sync();
      shape.fill.load('type');
      shape.lineFormat.load('visible');
      await context.sync();
      expect(shape.fill.type).toBe('NoFill');
      expect(shape.lineFormat.visible).toBe(false);
      // A 1 by 1 PNG.
      shape.fill.setImage(
        'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mNk+M9QDwADhgGAWjR9awAAAABJRU5ErkJggg=='
      );
      await context.sync();
      shape.fill.load('type');
      await context.sync();
      expect(shape.fill.type).toBe('PictureAndTexture');
    });
    const error = await rejection(
      runtime.run(async (context) => {
        box(context).fill.setImage('bm90IGFuIGltYWdl');
        await context.sync();
      })
    );
    expect(code(error)).toBe('InvalidArgument');
  });

  test('group by id, read the members, and ungroup', async () => {
    const runtime = await open();
    await runtime.run(async (context) => {
      const shapes = context.presentation.slides.getItemAt(0).shapes;
      const group = shapes.addGroup(['2', '3']);
      group.load('type,id');
      const members = group.group.shapes.load('items/name');
      await context.sync();
      expect(group.type).toBe('Group');
      expect(members.items.map((s) => s.name)).toEqual(['Box', 'Other']);
      const parent = members.items[0]!.parentGroup.load('id');
      const level = members.items[0]!.load('level');
      await context.sync();
      expect(parent.id).toBe(group.id);
      expect(level.level).toBe(1);
      group.group.ungroup();
      await context.sync();
      const top = shapes.load('items/name');
      await context.sync();
      expect(top.items.map((s) => s.name)).toEqual(['Box', 'Other']);
    });
  });

  test('parents, placeholder format, and delete', async () => {
    const runtime = await open();
    await runtime.run(async (context) => {
      const shape = box(context);
      const slide = shape.getParentSlide().load('index');
      const layout = shape.getParentSlideLayoutOrNullObject().load('id');
      const ph = shape.placeholderFormat.load();
      await context.sync();
      expect(slide.index).toBe(0);
      expect(layout.isNullObject).toBe(true);
      expect(ph.type).toBe('Unsupported');
      shape.delete();
      await context.sync();
      const count = context.presentation.slides.getItemAt(0).shapes.getCount();
      await context.sync();
      expect(count.value).toBe(1);
    });
  });
});

describe('text', () => {
  test('text frame, text range, font, and paragraph format', async () => {
    const runtime = await open();
    await runtime.run(async (context) => {
      const frame = box(context).textFrame;
      frame.verticalAlignment = 'Bottom';
      frame.leftMargin = 12;
      frame.wordWrap = false;
      frame.autoSizeSetting = 'AutoSizeNone';
      const world = frame.textRange.getSubstring(6, 5);
      world.font.bold = true;
      world.font.color = '#112233';
      world.font.size = 18;
      world.font.underline = 'Double';
      world.paragraphFormat.horizontalAlignment = 'Center';
      world.paragraphFormat.bulletFormat.visible = true;
      await context.sync();
      frame.load();
      const read = world.load('text,start,length');
      const font = world.font.load();
      const paragraph = world.paragraphFormat.load('horizontalAlignment');
      const bullets = world.paragraphFormat.bulletFormat.load('visible');
      await context.sync();
      expect(frame.toJSON()).toMatchObject({
        verticalAlignment: 'Bottom',
        leftMargin: 12,
        wordWrap: false,
        autoSizeSetting: 'AutoSizeNone',
        hasText: true,
      });
      expect([read.text, read.start, read.length]).toEqual(['World', 6, 5]);
      expect(font.toJSON()).toMatchObject({
        bold: true,
        color: '#112233',
        size: 18,
        underline: 'Double',
      });
      expect(paragraph.horizontalAlignment).toBe('Center');
      expect(bullets.visible).toBe(true);
      const whole = frame.textRange.font.load('bold');
      await context.sync();
      expect(whole.bold).toBe(true);
    });
  });

  test('indent levels count from 1, one more than the outline level in the file', async () => {
    const runtime = await open();
    await runtime.run(async (context) => {
      const range = box(context).textFrame.textRange;
      const first = range.getSubstring(0, 5).paragraphFormat.load('indentLevel');
      await context.sync();
      expect(first.indentLevel).toBe(1);
      range.getSubstring(6, 5).paragraphFormat.indentLevel = 2;
      await context.sync();
    });
    const bytes = await runtime.save();
    const zip = readZip(bytes);
    if (!zip.ok) throw new Error('saved');
    const slide = new TextDecoder().decode(new Map(zip.entries).get('/ppt/slides/slide1.xml'));
    expect(slide.match(/lvl="\d"/g)).toEqual(['lvl="1"']);
    const reopened = await reopen(runtime);
    await reopened.run(async (context) => {
      const second = box(context).textFrame.textRange.getSubstring(6, 5).paragraphFormat;
      second.load('indentLevel');
      await context.sync();
      expect(second.indentLevel).toBe(2);
    });
  });

  test('write text with paragraphs and line breaks, then delete it', async () => {
    const runtime = await open();
    await runtime.run(async (context) => {
      const range = box(context).textFrame.textRange;
      range.text = 'One\rTwo\vthree';
      await context.sync();
      range.load('text');
      await context.sync();
      expect(range.text).toBe('One\rTwo\vthree');
      const tail = range.getSubstring(4).load('text');
      await context.sync();
      expect(tail.text).toBe('Two\vthree');
      box(context).textFrame.deleteText();
      const frame = box(context).textFrame.load('hasText');
      await context.sync();
      expect(frame.hasText).toBe(false);
    });
  });

  test('a substring to the end is usable for writes in the same sync', async () => {
    const runtime = await open();
    await runtime.run(async (context) => {
      const range = box(context).textFrame.textRange.getSubstring(6);
      range.font.italic = true;
      range.load('text');
      await context.sync();
      expect(range.text).toBe('World');
      const italic = box(context).textFrame.textRange.getSubstring(0, 5).font.load('italic');
      await context.sync();
      expect(italic.italic).toBe(false);
    });
  });

  test('a shape with no text frame answers a null frame', async () => {
    const runtime = await open();
    await runtime.run(async (context) => {
      const other = context.presentation.slides.getItemAt(0).shapes.getItem('3');
      const frame = other.getTextFrameOrNullObject().load('hasText');
      const boxFrame = box(context).getTextFrameOrNullObject().load('hasText');
      await context.sync();
      expect(frame.isNullObject).toBe(true);
      expect(boxFrame.isNullObject).toBe(false);
      expect(boxFrame.hasText).toBe(true);
    });
  });

  test('hyperlinks on text and on shapes', async () => {
    const runtime = await open();
    await runtime.run(async (context) => {
      const range = box(context).textFrame.textRange.getSubstring(0, 5);
      range.setHyperlink({ address: 'https://example.com/a', screenTip: 'A' });
      context.presentation.slides
        .getItemAt(0)
        .shapes.getItem('3')
        .setHyperlink({ address: 'mailto:team@example.com' });
      await context.sync();
    });
    const again = await reopen(runtime);
    await again.run(async (context) => {
      const links = context.presentation.slides.getItemAt(0).hyperlinks.load('items');
      await context.sync();
      expect(links.items.map((l) => [l.type, l.address, l.screenTip])).toEqual([
        ['TextRange', 'https://example.com/a', 'A'],
        ['Shape', 'mailto:team@example.com', ''],
      ]);
      const linked = links.items[0]!.getLinkedTextRangeOrNullObject().load('text');
      await context.sync();
      expect(linked.text).toBe('Hello');
      links.items[1]!.delete();
      await context.sync();
      const count = context.presentation.slides.getItemAt(0).hyperlinks.getCount();
      await context.sync();
      expect(count.value).toBe(1);
      links.items[0]!.address = 'https://example.com/b';
      await context.sync();
      const relinked = context.presentation.slides.getItemAt(0).hyperlinks.load('items/address');
      await context.sync();
      expect(relinked.items.map((l) => l.address)).toEqual(['https://example.com/b']);
    });
    const error = await rejection(
      again.run(async (context) => {
        box(context).setHyperlink({ address: 'javascript:alert(1)' });
        await context.sync();
      })
    );
    expect(code(error)).toBe('InvalidArgument');
  });
});

describe('tables', () => {
  test('add a table with values, edit cells, rows, and columns, and merge', async () => {
    const runtime = await open();
    await runtime.run(async (context) => {
      const shape = context.presentation.slides.getItemAt(1).shapes.addTable(2, 3, {
        values: [
          ['a', 'b', 'c'],
          ['d', 'e', 'f'],
        ],
        left: 20,
        top: 20,
      });
      const table = shape.getTable();
      table.load('rowCount,columnCount,values');
      await context.sync();
      expect([table.rowCount, table.columnCount]).toEqual([2, 3]);
      expect(table.values).toEqual([
        ['a', 'b', 'c'],
        ['d', 'e', 'f'],
      ]);
      const cell = table.getCellOrNullObject(1, 2);
      cell.text = 'F';
      table.rows.add(undefined, 1);
      table.columns.getItemAt(0).delete();
      await context.sync();
      const outside = table.getCellOrNullObject(9, 9).load('text');
      table.load('values');
      const rows = table.rows.load('items');
      await context.sync();
      expect(outside.isNullObject).toBe(true);
      expect(table.values[1]).toEqual(['e', 'F']);
      expect(rows.items).toHaveLength(3);
      table.mergeCells(0, 0, 2, 1);
      rows.items[2]!.height = 40;
      await context.sync();
      const merged = table.getCellOrNullObject(1, 0).load('rowIndex,rowCount');
      const heights = table.rows.load('items');
      await context.sync();
      expect([merged.rowIndex, merged.rowCount]).toEqual([0, 2]);
      expect(heights.items[2]!.height).toBe(40);
    });
    const error = await rejection(
      runtime.run(async (context) => {
        context.presentation.slides.getItemAt(1).shapes.addTable(1, 1, { style: 'MediumStyle2' });
        await context.sync();
      })
    );
    // The deck does not define that style, so the whole batch, the new table too, is refused.
    expect(code(error)).toBe('NotSupported');
    expect(isPptxEditorError(error) && error.debugInfo.innerError?.code).toBe(
      'edit.table-style-undefined'
    );
  });

  test('getTable on a shape that is no table refuses at the sync', async () => {
    const runtime = await open();
    const error = await rejection(
      runtime.run(async (context) => {
        box(context).getTable().load('rowCount');
        await context.sync();
      })
    );
    expect(code(error)).toBe('NotSupported');
  });
});
