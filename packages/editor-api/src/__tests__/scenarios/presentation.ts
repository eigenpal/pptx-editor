/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/editor-api/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Scenarios for the presentation, its document properties, custom properties, page setup,
// and tags.

import type { RequestContext } from '../../runtime/public.ts';
import { check, dateJson, pt, same, type Json, type Scenario } from './types.ts';

/** The shape names of the second slide of `slide-layouts`. */
async function shapeNames(context: RequestContext): Promise<Json> {
  const shapes = context.presentation.slides.getItemAt(1).shapes.load('items/name');
  await context.sync();
  return shapes.items.map((s) => s.name);
}

async function coreProperties(context: RequestContext): Promise<Json> {
  const props = context.presentation.properties;
  props.load(
    'title,author,subject,keywords,category,comments,company,manager,revisionNumber,creationDate'
  );
  await context.sync();
  return {
    title: props.title,
    author: props.author,
    subject: props.subject,
    keywords: props.keywords,
    category: props.category,
    comments: props.comments,
    company: props.company,
    manager: props.manager,
    revisionNumber: props.revisionNumber,
    creationDate: dateJson(props.creationDate),
  };
}

async function customProperties(context: RequestContext): Promise<Json> {
  const custom = context.presentation.properties.customProperties;
  custom.load('items/key,items/type,items/value');
  await context.sync();
  return custom.items.map((p) => ({
    key: p.key,
    type: p.type,
    value: p.value instanceof Date ? dateJson(p.value) : p.value,
  }));
}

async function tags(context: RequestContext): Promise<Json> {
  const presentation = context.presentation.tags.load('items/key,items/value');
  const slide = context.presentation.slides.getItemAt(0);
  const slideTags = slide.tags.load('items/key,items/value');
  const shapeTags = slide.shapes.getItemAt(0).tags.load('items/key,items/value');
  await context.sync();
  const list = (c: typeof presentation) => c.items.map((t) => [t.key, t.value]);
  return { presentation: list(presentation), slide: list(slideTags), shape: list(shapeTags) };
}

export const PRESENTATION_SCENARIOS: readonly Scenario[] = [
  {
    name: 'presentation.read-basics',
    area: 'presentation',
    deck: 'slide-layouts',
    writes: false,
    async script(context) {
      const presentation = context.presentation;
      presentation.load('id,title');
      const setup = presentation.pageSetup.load('slideWidth,slideHeight');
      const count = presentation.slides.getCount();
      const masters = presentation.slideMasters.getCount();
      await context.sync();
      check(count.value === 6, 'the deck has six slides');
      return {
        idIsText: typeof presentation.id === 'string',
        title: presentation.title,
        slideWidth: pt(setup.slideWidth),
        slideHeight: pt(setup.slideHeight),
        slides: count.value,
        masters: masters.value,
      };
    },
  },
  {
    name: 'presentation.document-properties',
    area: 'presentation',
    deck: 'slide-layouts',
    writes: true,
    async script(context) {
      const props = context.presentation.properties;
      props.title = 'Quarterly review';
      props.author = 'Automation';
      props.subject = 'Results';
      props.keywords = 'q3; review';
      props.category = 'Finance';
      props.comments = 'Written by an add-in';
      props.company = 'Example Ltd';
      props.manager = 'A. Manager';
      props.revisionNumber = 7;
      await context.sync();
      const read = (await coreProperties(context)) as Record<string, Json>;
      same(read.title, 'Quarterly review', 'title');
      same(read.revisionNumber, 7, 'revision number');
      return read;
    },
    readBack: coreProperties,
  },
  {
    name: 'presentation.custom-properties',
    area: 'presentation',
    deck: 'slide-layouts',
    writes: true,
    async script(context) {
      const custom = context.presentation.properties.customProperties;
      custom.add('Client', 'Contoso');
      custom.add('Budget', 1250000);
      custom.add('Ratio', 0.75);
      custom.add('Approved', true);
      custom.add('Due', new Date(Date.UTC(2026, 10, 30, 12, 0, 0)));
      custom.add('Draft', 'remove me');
      await context.sync();
      custom.getItem('client').value = 'Fabrikam';
      custom.getItem('Draft').delete();
      const missing = custom.getItemOrNullObject('Draft');
      const count = custom.getCount();
      await context.sync();
      check(missing.isNullObject, 'a deleted property reads as a null object');
      same(count.value, 5, 'five properties remain');
      return customProperties(context);
    },
    readBack: customProperties,
  },
  {
    name: 'presentation.bindings',
    area: 'presentation',
    deck: 'slide-layouts',
    writes: true,
    async script(context) {
      const bindings = context.presentation.bindings;
      const shapes = context.presentation.slides.getItemAt(1).shapes;
      const title = shapes.getItem('2');
      const box = shapes.addTextBox('Bound box', { left: 40, top: 400, width: 200, height: 30 });
      bindings.add(title, 'Shape', 'title');
      bindings.add(box, 'Shape', 'box');
      // No sync between: the lookups see the bindings queued before them.
      const found = bindings.getItem('title').getShape().load('name');
      const added = bindings.getItem('box').getShape().load('name');
      const count = bindings.getCount();
      await context.sync();
      same(count.value, 2, 'two bindings');
      // The same id moves the binding to another shape.
      bindings.add(box, 'Shape', 'title');
      const moved = bindings.getItem('title').getShape().load('name');
      await context.sync();
      // Delete a binding found by position; its shape stays.
      bindings.getItemAt(1).delete();
      const gone = bindings.getItemOrNullObject('box');
      const all = bindings.load('items');
      await context.sync();
      check(gone.isNullObject, 'a deleted binding reads as a null object');
      // A binding whose shape is deleted refuses getShape.
      box.delete();
      await context.sync();
      let orphan = 'answered';
      try {
        bindings.getItem('title').getShape().load('name');
        await context.sync();
      } catch (error) {
        orphan = (error as { code?: string }).code ?? 'error';
      }
      return {
        names: [found.name, added.name, moved.name],
        ids: all.items.map((b) => b.id),
        orphan,
        shapes: await shapeNames(context),
      };
    },
    // Bindings live in memory and are never saved: the read-back reads the shapes only.
    readBack: shapeNames,
  },
  {
    name: 'presentation.page-setup',
    area: 'presentation',
    deck: 'slide-layouts',
    writes: true,
    async script(context) {
      const setup = context.presentation.pageSetup;
      setup.slideWidth = 720;
      setup.slideHeight = 405;
      await context.sync();
      setup.load('slideWidth,slideHeight');
      await context.sync();
      return { slideWidth: pt(setup.slideWidth), slideHeight: pt(setup.slideHeight) };
    },
    async readBack(context) {
      const setup = context.presentation.pageSetup.load('slideWidth,slideHeight');
      await context.sync();
      return { slideWidth: pt(setup.slideWidth), slideHeight: pt(setup.slideHeight) };
    },
  },
  {
    name: 'presentation.tags',
    area: 'slides',
    deck: 'slide-layouts',
    writes: true,
    async script(context) {
      const presentation = context.presentation;
      presentation.tags.add('WORKFLOW', 'draft');
      presentation.tags.add('OWNER', 'finance');
      const slide = presentation.slides.getItemAt(0);
      slide.tags.add('SECTION', 'intro');
      slide.tags.add('TEMP', 'x');
      slide.shapes.getItemAt(0).tags.add('ROLE', 'title');
      await context.sync();
      presentation.tags.getItem('WORKFLOW').value = 'final';
      slide.tags.delete('TEMP');
      const gone = slide.tags.getItemOrNullObject('TEMP');
      const key = presentation.tags.getItemAt(0).load('key');
      await context.sync();
      check(gone.isNullObject, 'a deleted tag reads as a null object');
      return { firstKey: key.key, all: await tags(context) };
    },
    readBack: tags,
  },
];
