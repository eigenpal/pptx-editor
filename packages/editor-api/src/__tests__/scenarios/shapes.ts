/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/editor-api/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Scenarios for shapes: add, geometry, z-order, groups, fill, line, and adjustments.

import type { RequestContext, Shape } from '../../runtime/public.ts';
import { check, pt, same, type Json, type Scenario } from './types.ts';

const BOX_PROPS = 'id,name,type,left,top,width,height,rotation,visible,zOrderPosition';

function boxJson(s: Shape): Json {
  return {
    id: s.id,
    name: s.name,
    type: s.type,
    left: pt(s.left),
    top: pt(s.top),
    width: pt(s.width),
    height: pt(s.height),
    rotation: pt(s.rotation),
    visible: s.visible,
    z: s.zOrderPosition,
  };
}

/** Every shape of one slide with its box. */
async function shapesOf(context: RequestContext, index: number): Promise<Json> {
  const shapes = context.presentation.slides.getItemAt(index).shapes;
  shapes.load(`items/${BOX_PROPS.split(',').join(',items/')}`);
  await context.sync();
  return shapes.items.map(boxJson);
}

async function formatOf(context: RequestContext, name: string): Promise<Json> {
  const shapes = context.presentation.slides.getItemAt(1).shapes;
  shapes.load('items/name');
  await context.sync();
  const shape = shapes.items.find((s) => s.name === name);
  check(shape, `a shape named ${name}`);
  const fill = shape.fill.load('type,foregroundColor,transparency');
  const line = shape.lineFormat.load('visible,color,weight,dashStyle,style,transparency');
  await context.sync();
  return {
    fill: {
      type: fill.type,
      color: fill.type === 'Solid' ? fill.foregroundColor : null,
      transparency: pt(fill.transparency),
    },
    line: line.visible
      ? {
          visible: true,
          color: line.color,
          weight: pt(line.weight),
          dashStyle: line.dashStyle,
          style: line.style,
          transparency: pt(line.transparency),
        }
      : { visible: false },
  };
}

export const SHAPE_SCENARIOS: readonly Scenario[] = [
  {
    name: 'shapes.add-geometric',
    area: 'shapes',
    deck: 'slide-layouts',
    writes: true,
    async script(context) {
      const shapes = context.presentation.slides.getItemAt(1).shapes;
      const star = shapes.addGeometricShape('Star5', {
        left: 500,
        top: 300,
        width: 120,
        height: 120,
      });
      star.name = 'Highlight star';
      star.fill.setSolidColor('#FFC000');
      star.lineFormat.color = '#C00000';
      star.lineFormat.weight = 2.5;
      star.lineFormat.dashStyle = 'Dash';
      const rect = shapes.addGeometricShape('Rectangle');
      rect.name = 'Default rectangle';
      await context.sync();
      star.load(BOX_PROPS);
      rect.load('width,height');
      await context.sync();
      check(rect.width > 0 && rect.height > 0, 'a shape added without options has a size');
      return {
        star: boxJson(star),
        format: await formatOf(context, 'Highlight star'),
      };
    },
    readBack: (context) => formatOf(context, 'Highlight star'),
  },
  {
    name: 'shapes.add-line',
    area: 'shapes',
    deck: 'slide-layouts',
    writes: true,
    async script(context) {
      const shapes = context.presentation.slides.getItemAt(1).shapes;
      const line = shapes.addLine('Straight', { left: 100, top: 400, width: 300, height: 0 });
      line.name = 'Divider';
      line.lineFormat.weight = 3;
      line.lineFormat.color = '#2F5597';
      const elbow = shapes.addLine('Elbow', { left: 100, top: 420, width: 200, height: 60 });
      elbow.name = 'Connector';
      await context.sync();
      return shapesOf(context, 1);
    },
    readBack: (context) => shapesOf(context, 1),
  },
  {
    name: 'shapes.add-text-box',
    area: 'shapes',
    deck: 'slide-layouts',
    writes: true,
    async script(context) {
      const slide = context.presentation.slides.getItemAt(0);
      const box = slide.shapes.addTextBox('Added by automation', {
        left: 60,
        top: 420,
        width: 360,
        height: 40,
      });
      box.name = 'Automation note';
      await context.sync();
      box.load('type,name');
      box.textFrame.textRange.load('text');
      await context.sync();
      same(box.textFrame.textRange.text, 'Added by automation', 'the text box text');
      return { type: box.type, name: box.name, shapes: await shapesOf(context, 0) };
    },
    readBack: (context) => shapesOf(context, 0),
  },
  {
    name: 'shapes.move-resize-rotate',
    area: 'shapes',
    deck: 'slide-layouts',
    writes: true,
    async script(context) {
      const note = context.presentation.slides.getItemAt(1).shapes.getItem('4');
      note.left = 40;
      note.top = 50;
      note.width = 250.5;
      note.height = 80;
      note.rotation = 15;
      note.name = 'Moved note';
      note.altTextTitle = 'Note';
      note.altTextDescription = 'A note about the chart';
      await context.sync();
      const subtitle = context.presentation.slides.getItemAt(0).shapes.getItem('3');
      subtitle.visible = false;
      await context.sync();
      note.load('altTextTitle,altTextDescription');
      await context.sync();
      return {
        alt: [note.altTextTitle, note.altTextDescription],
        slide1: await shapesOf(context, 0),
        slide2: await shapesOf(context, 1),
      };
    },
    async readBack(context) {
      return { slide1: await shapesOf(context, 0), slide2: await shapesOf(context, 1) };
    },
  },
  {
    name: 'shapes.z-order',
    area: 'shapes',
    deck: 'slide-layouts',
    writes: true,
    async script(context) {
      const shapes = context.presentation.slides.getItemAt(4).shapes;
      const back = shapes.addGeometricShape('Rectangle', {
        left: 10,
        top: 10,
        width: 50,
        height: 50,
      });
      back.name = 'Back';
      const front = shapes.addGeometricShape('Ellipse', {
        left: 20,
        top: 20,
        width: 50,
        height: 50,
      });
      front.name = 'Front';
      await context.sync();
      front.setZOrder('SendToBack');
      back.setZOrder('BringToFront');
      await context.sync();
      shapes.getItem('2').setZOrder('BringForward');
      await context.sync();
      const order = shapes.load('items/name,items/zOrderPosition');
      await context.sync();
      return order.items.map((s) => `${s.zOrderPosition}:${s.name}`);
    },
    async readBack(context) {
      const order = context.presentation.slides.getItemAt(4).shapes.load('items/name');
      await context.sync();
      return order.items.map((s) => s.name);
    },
  },
  {
    name: 'shapes.group-ungroup',
    area: 'shapes',
    deck: 'slide-layouts',
    writes: true,
    async script(context) {
      const shapes = context.presentation.slides.getItemAt(3).shapes;
      const a = shapes.addGeometricShape('Rectangle', {
        left: 50,
        top: 300,
        width: 80,
        height: 40,
      });
      a.name = 'Part A';
      const b = shapes.addGeometricShape('Ellipse', { left: 150, top: 300, width: 80, height: 40 });
      b.name = 'Part B';
      const c = shapes.addGeometricShape('Diamond', { left: 250, top: 300, width: 40, height: 40 });
      c.name = 'Part C';
      const group = shapes.addGroup([a, b, '3']);
      group.name = 'Legend';
      await context.sync();
      group.load('type,left,top,width,height');
      const members = group.group.shapes.load('items/name,items/level');
      await context.sync();
      const grouped = {
        type: group.type,
        box: [pt(group.left), pt(group.top), pt(group.width), pt(group.height)],
        members: members.items.map((s) => `${s.name}@${s.level}`),
      };
      const parent = members.items[0]!.parentGroup.load('name');
      await context.sync();
      same(parent.name, 'Legend', 'a member knows its group');
      const second = shapes.addGroup([c, group]);
      second.name = 'Temporary';
      await context.sync();
      second.group.ungroup();
      await context.sync();
      const after = shapes.load('items/name,items/type');
      await context.sync();
      return { grouped, after: after.items.map((s) => `${s.name}:${s.type}`) };
    },
    async readBack(context) {
      const shapes = context.presentation.slides.getItemAt(3).shapes.load('items/name,items/type');
      await context.sync();
      const group = shapes.items.find((s) => s.type === 'Group');
      const members = group ? group.group.shapes.load('items/name') : null;
      await context.sync();
      return {
        shapes: shapes.items.map((s) => `${s.name}:${s.type}`),
        members: members ? members.items.map((s) => s.name) : [],
      };
    },
  },
  {
    name: 'shapes.fill-and-line',
    area: 'shapes',
    deck: 'slide-layouts',
    writes: true,
    async script(context) {
      const note = context.presentation.slides.getItemAt(1).shapes.getItem('4');
      const before = await formatOf(context, 'Note box');
      note.fill.setSolidColor('#70AD47');
      note.fill.transparency = 0.4;
      note.lineFormat.visible = true;
      note.lineFormat.color = '#385723';
      note.lineFormat.weight = 4;
      note.lineFormat.style = 'ThickThin';
      note.lineFormat.dashStyle = 'SquareDot';
      await context.sync();
      const after = await formatOf(context, 'Note box');
      const title = context.presentation.slides.getItemAt(1).shapes.getItem('2');
      title.fill.setSolidColor('#000000');
      title.fill.clear();
      title.lineFormat.visible = false;
      await context.sync();
      return { before, after, cleared: await formatOf(context, 'Title 1') };
    },
    async readBack(context) {
      return {
        note: await formatOf(context, 'Note box'),
        title: await formatOf(context, 'Title 1'),
      };
    },
  },
  {
    name: 'shapes.adjustments',
    area: 'shapes',
    deck: 'slide-layouts',
    writes: true,
    async script(context) {
      const shapes = context.presentation.slides.getItemAt(4).shapes;
      const card = shapes.addGeometricShape('RoundRectangle', {
        left: 100,
        top: 200,
        width: 200,
        height: 100,
      });
      card.name = 'Card';
      const arrow = shapes.addGeometricShape('RightArrow', {
        left: 350,
        top: 200,
        width: 150,
        height: 60,
      });
      arrow.name = 'Arrow';
      await context.sync();
      card.adjustments.load('count');
      arrow.adjustments.load('count');
      const radius = card.adjustments.get(0);
      await context.sync();
      const initial = { card: card.adjustments.count, arrow: arrow.adjustments.count };
      const defaultRadius = Math.round(radius.value * 1000) / 1000;
      card.adjustments.set(0, 0.3);
      arrow.adjustments.set(1, 0.25);
      await context.sync();
      return { initial, defaultRadius, read: await this.readBack!(context) };
    },
    async readBack(context) {
      const shapes = context.presentation.slides.getItemAt(4).shapes.load('items/name');
      await context.sync();
      const byName = (n: string) => shapes.items.find((s) => s.name === n)!;
      const radius = byName('Card').adjustments.get(0);
      const head = byName('Arrow').adjustments.get(1);
      await context.sync();
      return {
        radius: Math.round(radius.value * 1000) / 1000,
        head: Math.round(head.value * 1000) / 1000,
      };
    },
  },
  {
    name: 'shapes.delete-and-placeholders',
    area: 'shapes',
    deck: 'slide-layouts',
    writes: true,
    async script(context) {
      const shapes = context.presentation.slides.getItemAt(2).shapes;
      const formats = shapes.load(
        'items/id,items/name,items/placeholderFormat/type,items/placeholderFormat/containedType'
      );
      await context.sync();
      const placeholders = formats.items.map((s) => ({
        name: s.name,
        type: s.placeholderFormat.type,
        contains: s.placeholderFormat.containedType,
      }));
      shapes.getItem('4').delete();
      await context.sync();
      const gone = shapes.getItemOrNullObject('4');
      const count = shapes.getCount();
      await context.sync();
      check(gone.isNullObject, 'a deleted shape is a null object');
      return { placeholders, count: count.value };
    },
    async readBack(context) {
      const shapes = context.presentation.slides.getItemAt(2).shapes.load('items/name');
      await context.sync();
      return shapes.items.map((s) => s.name);
    },
  },
];
