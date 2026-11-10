/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/editor-api/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
import type { RequestContext } from '../../runtime/public.ts';
import { check, same, type Json, type Scenario } from './types.ts';

async function gradientRead(context: RequestContext): Promise<Json> {
  const result = context.presentation.slides.getItemAt(0).shapes.getItemAt(0).getFormat();
  await context.sync();
  const gradient = result.value.gradient;
  check(gradient, 'gradient exists');
  return {
    angle: gradient.angle,
    stops: gradient.stops.map((s) => ({ position: s.position, opacity: s.opacity })),
  };
}

async function effectsRead(context: RequestContext): Promise<Json> {
  const shadow = context.presentation.slides.getItemAt(0).shapes.getItemAt(0).getFormat();
  const reflection = context.presentation.slides.getItemAt(1).shapes.getItemAt(0).getFormat();
  await context.sync();
  check(shadow.value.shadow, 'shadow exists');
  check(reflection.value.reflection, 'reflection exists');
  return {
    blur: shadow.value.shadow.blur,
    distance: shadow.value.shadow.distance,
    opacity: shadow.value.shadow.opacity,
    reflectionDistance: reflection.value.reflection.distance,
    reflectionOpacity: reflection.value.reflection.startOpacity,
  };
}

async function endpointsRead(context: RequestContext): Promise<Json> {
  const shapes = context.presentation.slides.getItemAt(0).shapes;
  shapes.load('items/name');
  await context.sync();
  const shape = shapes.items.find((s) => s.name === 'Inspector line');
  check(shape, 'line exists');
  shape.lineFormat.load('startArrowhead,endArrowhead,weight');
  await context.sync();
  return {
    start: shape.lineFormat.startArrowhead,
    end: shape.lineFormat.endArrowhead,
    width: shape.lineFormat.weight,
  };
}

export const FORMAT_SCENARIOS: readonly Scenario[] = [
  {
    name: 'shapes.line-endpoint-types',
    area: 'shapes',
    deck: 'slide-layouts',
    writes: true,
    script: async (context) => {
      const line = context.presentation.slides
        .getItemAt(0)
        .shapes.addLine('Straight', { left: 10, top: 20, width: 100, height: 20 });
      line.name = 'Inspector line';
      line.lineFormat.startArrowhead = 'diamond';
      line.lineFormat.endArrowhead = 'triangle';
      line.lineFormat.weight = 3.125;
      await context.sync();
      return endpointsRead(context);
    },
    readBack: endpointsRead,
  },
  {
    name: 'shapes.patch-gradient',
    area: 'shapes',
    deck: 'gradients',
    writes: true,
    async script(context) {
      const shape = context.presentation.slides.getItemAt(0).shapes.getItemAt(0);
      const read = shape.getFormat();
      await context.sync();
      const gradient = read.value.gradient;
      check(gradient, 'gradient exists before editing');
      shape.patchFormat({
        kind: 'gradient',
        angle: 120,
        stop: { id: gradient.stops[1]!.id, opacity: 0.35 },
      });
      await context.sync();
      const after = await gradientRead(context);
      same(
        after,
        {
          angle: 120,
          stops: [
            { position: 0, opacity: 1 },
            { position: 0.5, opacity: 0.35 },
            { position: 1, opacity: 1 },
          ],
        },
        'partial gradient values'
      );
      return after;
    },
    readBack: gradientRead,
  },
  {
    name: 'shapes.patch-effects',
    area: 'shapes',
    deck: 'shape-effects-presets',
    writes: true,
    async script(context) {
      context.presentation.slides
        .getItemAt(0)
        .shapes.getItemAt(0)
        .patchFormat({ kind: 'shadow', blur: 7, distance: 9, opacity: 0.6 });
      context.presentation.slides
        .getItemAt(1)
        .shapes.getItemAt(0)
        .patchFormat({ kind: 'reflection', distance: 5, startOpacity: 0.75 });
      await context.sync();
      const after = await effectsRead(context);
      same(
        after,
        { blur: 7, distance: 9, opacity: 0.6, reflectionDistance: 5, reflectionOpacity: 0.75 },
        'partial effect values'
      );
      return after;
    },
    readBack: effectsRead,
  },
];
