/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/editor-api/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Scenarios for custom XML parts, slide import, pictures, selection, load paths, and queued
// syncs.

import type { RequestContext } from '../../runtime/public.ts';
import { outline } from './slides.ts';
import { check, pngSize, refusal, same, type Json, type Scenario } from './types.ts';

const NS = 'urn:example:review';

async function customXmlOf(context: RequestContext): Promise<Json> {
  const parts = context.presentation.customXmlParts.load('items/id,items/namespaceUri');
  const slideParts = context.presentation.slides
    .getItemAt(0)
    .customXmlParts.load('items/namespaceUri');
  await context.sync();
  const ours = parts.items.filter((p) => p.namespaceUri === NS);
  const xml = ours.map((p) => p.getXml());
  await context.sync();
  return {
    namespaces: parts.items.map((p) => p.namespaceUri),
    xml: xml.map((x) => x.value),
    slide: slideParts.items.map((p) => p.namespaceUri),
  };
}

async function titles(context: RequestContext): Promise<Json> {
  const slides = context.presentation.slides.load('items/id');
  await context.sync();
  const shapes = slides.items.map((s) =>
    s.shapes.load('items/name,items/textFrame/textRange/text')
  );
  const layouts = slides.items.map((s) => s.layout.load('name'));
  await context.sync();
  return slides.items.map((s, i) => ({
    layout: layouts[i]!.name,
    text: shapes[i]!.items.map((shape) => {
      try {
        return shape.textFrame.textRange.text;
      } catch {
        return `<${shape.name}>`;
      }
    }),
  }));
}

export const PLATFORM_SCENARIOS: readonly Scenario[] = [
  {
    name: 'custom-xml.lifecycle',
    area: 'custom-xml',
    deck: 'slide-layouts',
    writes: true,
    async script(context) {
      const parts = context.presentation.customXmlParts;
      const review = parts.add(`<review xmlns="${NS}"><status>draft</status></review>`);
      const scratch = parts.add('<scratch xmlns="urn:example:scratch"/>');
      context.presentation.slides
        .getItemAt(0)
        .customXmlParts.add('<slideNote xmlns="urn:example:slide">intro</slideNote>');
      await context.sync();
      review.load('id,namespaceUri');
      await context.sync();
      check(review.id.length > 0, 'a new part has an id');
      review.setXml(`<review xmlns="${NS}"><status>approved</status></review>`);
      scratch.delete();
      await context.sync();
      const only = parts.getByNamespace(NS).getOnlyItem();
      const none = parts.getByNamespace('urn:example:scratch').getOnlyItemOrNullObject();
      const byId = parts.getItem(review.id).getXml();
      await context.sync();
      check(none.isNullObject, 'a deleted part is gone from its namespace');
      check(byId.value.includes('approved'), 'setXml replaced the text');
      return { only: only.isNullObject, read: await customXmlOf(context) };
    },
    readBack: customXmlOf,
  },
  {
    name: 'insert.keep-source-formatting',
    area: 'insert',
    deck: 'slide-layouts',
    writes: true,
    async script(context, assets) {
      context.presentation.insertSlidesFromBase64(assets.base64['text-runs'], {
        formatting: 'KeepSourceFormatting',
        targetSlideId: '257',
      });
      await context.sync();
      const count = context.presentation.slides.getCount();
      await context.sync();
      same(count.value, 7, 'one slide was inserted');
      return titles(context);
    },
    readBack: titles,
  },
  {
    name: 'insert.use-destination-theme',
    area: 'insert',
    deck: 'slide-layouts',
    writes: true,
    async script(context, assets) {
      context.presentation.insertSlidesFromBase64(assets.base64.tables, {
        formatting: 'UseDestinationTheme',
        sourceSlideIds: ['257', '258'],
      });
      await context.sync();
      return titles(context);
    },
    readBack: titles,
  },
  {
    name: 'images.slide-and-shape',
    area: 'images',
    deck: 'pictures',
    writes: false,
    async script(context) {
      const slide = context.presentation.slides.getItemAt(0);
      const shapes = slide.shapes.load('items/id');
      await context.sync();
      const whole = slide.getImageAsBase64({ width: 480 });
      const fitted = slide.getImageAsBase64({ width: 200, height: 200 });
      const shape = shapes.items[shapes.items.length - 1]!.getImageAsBase64({ format: 'Png' });
      await context.sync();
      return {
        slide: pngSize(whole.value),
        fitted: pngSize(fitted.value),
        shape: pngSize(shape.value),
      };
    },
  },
  {
    name: 'selection.slides',
    area: 'selection',
    deck: 'slide-layouts',
    writes: false,
    host: 'selection',
    async script(context) {
      context.presentation.setSelectedSlides(['258', '259']);
      await context.sync();
      const selected = context.presentation.getSelectedSlides().load('items/id');
      await context.sync();
      const exported = context.presentation.getSelectedSlides().exportAsBase64Presentation();
      await context.sync();
      check(exported.value.length > 100, 'the selected slides export as a presentation');
      return { ids: selected.items.map((s) => s.id) };
    },
  },
  {
    name: 'selection.shapes',
    area: 'selection',
    deck: 'slide-layouts',
    writes: true,
    host: 'selection',
    async script(context) {
      const slide = context.presentation.slides.getItemAt(1);
      const badge = slide.shapes.addGeometricShape('Ellipse', {
        left: 700,
        top: 100,
        width: 60,
        height: 60,
      });
      badge.name = 'Badge';
      badge.load('id');
      await context.sync();
      slide.setSelectedShapes([badge.id, '4']);
      await context.sync();
      const selected = context.presentation.getSelectedShapes();
      selected.load('items/id,items/name');
      await context.sync();
      const names = selected.items.map((s) => s.name);
      // Office.js groups the selection without a load of its items.
      const group = context.presentation.getSelectedShapes().group();
      group.name = 'Selected group';
      await context.sync();
      group.load('type');
      const members = group.group.shapes.load('items/name');
      await context.sync();
      return { names, groupType: group.type, members: members.items.map((s) => s.name) };
    },
    async readBack(context) {
      const shapes = context.presentation.slides.getItemAt(1).shapes.load('items/name,items/type');
      await context.sync();
      return shapes.items.map((s) => `${s.name}:${s.type}`);
    },
  },
  {
    name: 'selection.text-range',
    area: 'selection',
    deck: 'slide-layouts',
    writes: true,
    host: 'selection',
    async script(context) {
      const range = context.presentation.slides.getItemAt(3).shapes.getItem('3')
        .textFrame.textRange;
      range.load('text');
      await context.sync();
      range.getSubstring(0, 4).setSelected();
      await context.sync();
      const selected = context.presentation.getSelectedTextRange();
      selected.load('text,start,length');
      await context.sync();
      selected.font.bold = true;
      await context.sync();
      const font = context.presentation.getSelectedTextRange().font.load('bold');
      await context.sync();
      return {
        text: selected.text,
        start: selected.start,
        length: selected.length,
        bold: font.bold,
        prefix: range.text.slice(0, 4),
      };
    },
  },
  {
    name: 'load.paths-and-all',
    area: 'load',
    deck: 'slide-layouts',
    writes: false,
    async script(context) {
      const slides = context.presentation.slides;
      slides.load('items/shapes/items/name');
      const second = slides.getItemAt(1);
      second.load({ shapes: { items: { name: true, textFrame: { textRange: { text: true } } } } });
      const window = context.presentation.slides
        .getItem('257')
        .shapes.load({ select: 'name', skip: 1, top: 1 });
      const all = second.shapes.getItem('4').load('$all');
      const expanded = second.shapes.getItem('4');
      expanded.load({ select: 'fill/type', expand: 'fill' });
      const deep = context.presentation.slides.getItemAt(3);
      deep.load('shapes/items/textFrame/textRange/text');
      await context.sync();
      return {
        names: slides.items.map((s) => s.shapes.items.map((shape) => shape.name)),
        tree: second.shapes.items.map((s) => s.textFrame.textRange.text.length),
        window: window.items.map((s) => s.name),
        all: { name: all.name, type: all.type, hasId: typeof all.id === 'string' },
        fill: expanded.fill.type,
        deep: deep.shapes.items.map((s) => s.textFrame.textRange.text.split('\r')[0]!),
      };
    },
  },
  {
    name: 'load.not-loaded',
    area: 'load',
    deck: 'slide-layouts',
    writes: false,
    async script(context) {
      const shape = context.presentation.slides.getItemAt(0).shapes.getItemAt(0);
      const early = await refusal(async () => shape.name);
      shape.load('name');
      const pending = await refusal(async () => shape.name);
      const count = context.presentation.slides.getCount();
      const value = await refusal(async () => count.value);
      await context.sync();
      const late = await refusal(async () => shape.left);
      return { early, pending, value, name: shape.name, count: count.value, late };
    },
  },
  {
    name: 'sync.queued',
    area: 'sync',
    deck: 'slide-layouts',
    writes: true,
    async script(context) {
      const shapes = context.presentation.slides.getItemAt(4).shapes;
      const first = shapes.addTextBox('first', { left: 10, top: 10, width: 100, height: 20 });
      first.name = 'Q1';
      const one = context.sync();
      const second = shapes.addTextBox('second', { left: 10, top: 40, width: 100, height: 20 });
      second.name = 'Q2';
      const two = context.sync();
      shapes.getItem('2').width = -1;
      const bad = context.sync();
      const third = shapes.addTextBox('third', { left: 10, top: 70, width: 100, height: 20 });
      third.name = 'Q3';
      const three = context.sync();
      const settled = await Promise.allSettled([one, two, bad, three]);
      const order = shapes.load('items/name');
      await context.sync();
      return {
        settled: settled.map((s) =>
          s.status === 'fulfilled' ? 'ok' : ((s.reason as { code?: string }).code ?? 'error')
        ),
        names: order.items.map((s) => s.name),
      };
    },
    async readBack(context) {
      const order = context.presentation.slides.getItemAt(4).shapes.load('items/name');
      await context.sync();
      return order.items.map((s) => s.name);
    },
  },
  {
    name: 'sync.tracked-objects',
    area: 'sync',
    deck: 'slide-layouts',
    writes: true,
    async script(context) {
      const shape = context.presentation.slides.getItemAt(0).shapes.getItem('2');
      context.trackedObjects.add(shape);
      shape.load('name');
      await context.sync();
      shape.name = `${shape.name} (tracked)`;
      await context.sync();
      context.trackedObjects.remove(shape);
      return outline(context);
    },
    readBack: outline,
  },
  {
    name: 'sync.retry-after-refusal',
    area: 'sync',
    deck: 'slide-layouts',
    writes: true,
    async script(context) {
      // A common add-in pattern: catch a refused sync, fix the cause, retry with the same proxies.
      const slide = context.presentation.slides.getItemAt(1);
      const note = slide.shapes.getItem('4');
      const title = slide.shapes.getItem('2').textFrame.textRange;
      note.name = 'Retried note';
      title.text = 'Retried title';
      note.width = -10;
      let first = 'applied';
      try {
        await context.sync();
      } catch (error) {
        first = (error as { code?: string }).code ?? 'error';
      }
      note.load('name');
      await context.sync();
      const afterRefusal = note.name;
      note.name = 'Retried note';
      title.text = 'Retried title';
      note.width = 200;
      note.load('name,width');
      title.load('text');
      await context.sync();
      return { first, afterRefusal, name: note.name, width: note.width, title: title.text };
    },
    async readBack(context) {
      const shapes = context.presentation.slides
        .getItemAt(1)
        .shapes.load('items/name,items/width,items/textFrame/textRange/text');
      await context.sync();
      return shapes.items.map((s) => [s.name, s.width, s.textFrame.textRange.text]);
    },
  },
  {
    name: 'sync.reads-see-earlier-writes',
    area: 'sync',
    deck: 'slide-layouts',
    writes: true,
    async script(context) {
      // One sync, in the order an add-in writes it: each read sees the writes before it.
      const p = context.presentation;
      p.tags.add('STATUS', 'draft');
      p.tags.getItemOrNullObject('STATUS').value = 'final';
      p.customXmlParts.add('<order xmlns="urn:example:order">1</order>');
      p.customXmlParts
        .getByNamespace('urn:example:order')
        .getOnlyItemOrNullObject()
        .setXml('<order xmlns="urn:example:order">2</order>');
      p.slides.add();
      const added = p.slides.getItemAt(6).load('id');
      p.slides.getItem('258').delete();
      const deleted = p.slides.getItemOrNullObject('258');
      const shapes = p.slides.getItemAt(0).shapes;
      const box = shapes.addTextBox('Fresh', { left: 40, top: 300, width: 200, height: 30 });
      box.name = 'Fresh box';
      const listed = shapes.load('items/name');
      const table = shapes.addTable(2, 2, { left: 300, top: 300 }).getTable();
      table.getCellOrNullObject(1, 1).text = 'cell';
      table.load('values');
      await context.sync();
      return {
        added: added.id,
        deleted: deleted.isNullObject,
        listed: listed.items.map((s) => s.name),
        values: table.values,
        read: await this.readBack!(context),
      };
    },
    async readBack(context) {
      const p = context.presentation;
      const tag = p.tags.getItem('STATUS').load('value');
      const xml = p.customXmlParts.getByNamespace('urn:example:order').getOnlyItem().getXml();
      const slides = p.slides.load('items/id');
      await context.sync();
      return { tag: tag.value, xml: xml.value, slides: slides.items.map((s) => s.id) };
    },
  },
  {
    name: 'selection.reads-see-earlier-writes',
    area: 'selection',
    deck: 'slide-layouts',
    writes: true,
    host: 'selection',
    async script(context) {
      const p = context.presentation;
      p.slides.getItemAt(1).setSelectedShapes(['4']);
      const shapes = p.getSelectedShapes().load('items/name');
      const text = p.slides.getItemAt(3).shapes.getItem('3').textFrame.textRange;
      text.getSubstring(0, 6).setSelected();
      const range = p.getSelectedTextRange();
      range.load('text');
      range.font.underline = 'Single';
      p.setSelectedSlides(['259', '260']);
      const slides = p.getSelectedSlides().load('items/id');
      await context.sync();
      return {
        shapes: shapes.items.map((s) => s.name),
        text: range.text,
        slides: slides.items.map((s) => s.id),
      };
    },
    async readBack(context) {
      const font = context.presentation.slides
        .getItemAt(3)
        .shapes.getItem('3')
        .textFrame.textRange.getSubstring(0, 6)
        .font.load('underline');
      await context.sync();
      return font.underline;
    },
  },
];
