/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/editor-api/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Scenarios for slides: add, delete, move, layouts and masters, backgrounds, and theme colors.

import type { RequestContext } from '../../runtime/public.ts';
import { check, same, type Json, type Scenario } from './types.ts';

/** Each slide as its id, its layout name, and the names of its shapes. */
export async function outline(context: RequestContext): Promise<Json> {
  const slides = context.presentation.slides;
  slides.load('items/id');
  await context.sync();
  const layouts = slides.items.map((s) => s.layout.load('name'));
  const shapes = slides.items.map((s) => s.shapes.load('items/name'));
  await context.sync();
  return slides.items.map((s, i) => ({
    id: s.id,
    layout: layouts[i]!.name,
    shapes: shapes[i]!.items.map((shape) => shape.name),
  }));
}

export const SLIDE_SCENARIOS: readonly Scenario[] = [
  {
    name: 'slides.layouts-and-masters',
    area: 'slides',
    deck: 'slide-layouts',
    writes: false,
    async script(context) {
      const masters = context.presentation.slideMasters;
      masters.load('items/id,items/name');
      await context.sync();
      const master = masters.items[0]!;
      const layouts = master.layouts.load('items/id,items/name,items/type');
      const masterShapes = master.shapes.load('items/name,items/type');
      await context.sync();
      const placeholders = layouts.items.map((l) =>
        l.shapes.load('items/name,items/placeholderFormat/type')
      );
      const byId = master.layouts.getItemOrNullObject(layouts.items[1]!.id).load('name');
      const none = master.layouts.getItemOrNullObject('no-such-layout');
      await context.sync();
      check(!byId.isNullObject && byId.name === layouts.items[1]!.name, 'a layout by its id');
      check(none.isNullObject, 'an unknown layout id is a null object');
      return {
        masters: masters.items.map((m) => m.name),
        masterShapes: masterShapes.items.map((s) => `${s.name}:${s.type}`),
        layouts: layouts.items.map((l, i) => ({
          name: l.name,
          type: l.type,
          placeholders: placeholders[i]!.items.map((s) => s.placeholderFormat.type),
        })),
      };
    },
  },
  {
    name: 'slides.add-with-layout',
    area: 'slides',
    deck: 'slide-layouts',
    writes: true,
    async script(context) {
      const master = context.presentation.slideMasters.getItemAt(0);
      const layouts = master.layouts.load('items/id,items/name');
      master.load('id');
      await context.sync();
      const titleOnly = layouts.items.find((l) => l.name === 'Title Only');
      check(titleOnly, 'the deck has a Title Only layout');
      context.presentation.slides.add({ layoutId: titleOnly.id, slideMasterId: master.id });
      context.presentation.slides.add();
      await context.sync();
      const slides = context.presentation.slides.load('items/id');
      await context.sync();
      same(slides.items.length, 8, 'two slides were added');
      const added = slides.items[6]!;
      const layout = added.layout.load('name');
      const shapes = added.shapes.load('items/name,items/placeholderFormat/type');
      const index = added.load('index');
      await context.sync();
      return {
        index: index.index,
        layout: layout.name,
        placeholders: shapes.items.map((s) => `${s.name}:${s.placeholderFormat.type}`),
        outline: await outline(context),
      };
    },
    readBack: outline,
  },
  {
    name: 'slides.delete',
    area: 'slides',
    deck: 'slide-layouts',
    writes: true,
    async script(context) {
      const slides = context.presentation.slides;
      slides.getItemAt(2).delete();
      slides.getItem('260').delete();
      await context.sync();
      const gone = slides.getItemOrNullObject('260');
      await context.sync();
      check(gone.isNullObject, 'a deleted slide reads as a null object');
      return outline(context);
    },
    readBack: outline,
  },
  {
    name: 'slides.move',
    area: 'slides',
    deck: 'slide-layouts',
    writes: true,
    async script(context) {
      const slides = context.presentation.slides;
      slides.getItem('261').moveTo(0);
      slides.getItem('256').moveTo(5);
      await context.sync();
      const moved = slides.getItem('261').load('index');
      await context.sync();
      same(moved.index, 0, 'the moved slide is first');
      return outline(context);
    },
    readBack: outline,
  },
  {
    name: 'slides.apply-layout',
    area: 'slides',
    deck: 'slide-layouts',
    writes: true,
    async script(context) {
      const layouts = context.presentation.slideMasters.getItemAt(0).layouts;
      layouts.load('items/name');
      await context.sync();
      const twoContent = layouts.items.find((l) => l.name === 'Two Content');
      check(twoContent, 'the deck has a Two Content layout');
      context.presentation.slides.getItemAt(1).applyLayout(twoContent);
      await context.sync();
      return outline(context);
    },
    readBack: outline,
  },
  {
    name: 'slides.background-and-theme',
    area: 'slides',
    deck: 'slide-layouts',
    writes: true,
    async script(context) {
      const slide = context.presentation.slides.getItemAt(0);
      slide.background.fill.setSolidFill({ color: '#FFF2CC', transparency: 0.25 });
      slide.background.areBackgroundGraphicsHidden = true;
      const accent = slide.themeColorScheme.getThemeColor('Accent1');
      await context.sync();
      const before = accent.value;
      slide.themeColorScheme.setThemeColor('Accent2', '#7030A0');
      await context.sync();
      return { accent1: before, read: await this.readBack!(context) };
    },
    async readBack(context) {
      const slide = context.presentation.slides.getItemAt(0);
      const background = slide.background.load('areBackgroundGraphicsHidden');
      const fill = slide.background.fill.load('type');
      const solid = slide.background.fill.getSolidFillOrNullObject().load('color,transparency');
      const accent2 = slide.themeColorScheme.getThemeColor('Accent2');
      const other = context.presentation.slides.getItemAt(3).themeColorScheme;
      const sharedAccent2 = other.getThemeColor('Accent2');
      await context.sync();
      return {
        hidden: background.areBackgroundGraphicsHidden,
        type: fill.type,
        color: solid.isNullObject ? null : solid.color,
        transparency: solid.isNullObject ? null : Math.round(solid.transparency * 100) / 100,
        accent2: accent2.value,
        sharedAccent2: sharedAccent2.value,
      };
    },
  },
];
