/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/editor-api/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Scenarios for refusals. Each one checks the Office.js error code and that the refused batch
// changed nothing: the state read before the batch equals the state read after it.

import type { RequestContext } from '../../runtime/public.ts';
import { outline } from './slides.ts';
import { refusal, same, type Json, type Scenario } from './types.ts';

/** The boxes and text of the second slide of `slide-layouts`. */
async function state(context: RequestContext): Promise<Json> {
  const shapes = context.presentation.slides
    .getItemAt(1)
    .shapes.load('items/name,items/left,items/width,items/textFrame/textRange/text');
  await context.sync();
  return {
    outline: await outline(context),
    boxes: shapes.items.map((s) => [s.name, s.left, s.width, s.textFrame.textRange.text]),
  };
}

/** The ids of the runtime's bindings, in the order they were added. */
async function bindingIds(context: RequestContext): Promise<Json> {
  const all = context.presentation.bindings.load('items');
  await context.sync();
  return all.items.map((b) => b.id);
}

/** Refuse `step`, and check that the presentation did not change. */
async function unchanged(
  context: RequestContext,
  step: () => Promise<unknown>,
  read: (context: RequestContext) => Promise<Json> = state
): Promise<Json> {
  const before = await read(context);
  const refused = await refusal(step);
  same(await read(context), before, `the refused batch (${refused.code}) changed nothing`);
  return refused;
}

/** The background fill type of the first slide, and its gradient type. */
async function gradientOf(context: RequestContext): Promise<Json> {
  const fill = context.presentation.slides.getItemAt(0).background.fill;
  fill.load('type');
  const gradient = fill.getGradientFillOrNullObject().load('type');
  await context.sync();
  return [fill.type, gradient.isNullObject ? null : gradient.type];
}

const note = (context: RequestContext) =>
  context.presentation.slides.getItemAt(1).shapes.getItem('4');

export const ERROR_SCENARIOS: readonly Scenario[] = [
  {
    name: 'errors.item-not-found',
    area: 'errors',
    deck: 'slide-layouts',
    writes: false,
    async script(context) {
      return {
        slide: await unchanged(context, async () => {
          note(context).name = 'Renamed';
          context.presentation.slides.getItem('9999').delete();
          await context.sync();
        }),
        shape: await unchanged(context, async () => {
          context.presentation.slides.getItemAt(1).shapes.getItem('99').load('name');
          await context.sync();
        }),
        index: await unchanged(context, async () => {
          context.presentation.slides.getItemAt(42).load('id');
          await context.sync();
        }),
      };
    },
  },
  {
    name: 'errors.invalid-argument',
    area: 'errors',
    deck: 'slide-layouts',
    writes: false,
    async script(context) {
      return {
        width: await unchanged(context, async () => {
          note(context).left = 10;
          note(context).width = -5;
          await context.sync();
        }),
        color: await unchanged(context, async () => {
          note(context).fill.setSolidColor('not a color');
          await context.sync();
        }),
        autosize: await unchanged(context, async () => {
          note(context).textFrame.autoSizeSetting = 'AutoSizeMixed';
          await context.sync();
        }),
        xml: await unchanged(context, async () => {
          context.presentation.customXmlParts.add('<a><b></a>');
          await context.sync();
        }),
        link: await unchanged(context, async () => {
          note(context).setHyperlink({ address: 'javascript:alert(1)' });
          await context.sync();
        }),
      };
    },
  },
  {
    name: 'errors.not-supported',
    area: 'errors',
    deck: 'pictures',
    writes: false,
    async script(context) {
      const read = async (c: RequestContext): Promise<Json> => outline(c);
      const shapes = context.presentation.slides.getItemAt(0).shapes.load('items/id,items/type');
      await context.sync();
      const picture = shapes.items.find((s) => s.type === 'Image');
      return {
        pictureText: await unchanged(
          context,
          async () => {
            picture!.name = 'Renamed';
            picture!.textFrame.textRange.text = 'text on a picture';
            await context.sync();
          },
          read
        ),
        tableStyle: await unchanged(
          context,
          async () => {
            const shape = context.presentation.slides.getItemAt(0).shapes.addTable(2, 2);
            shape.getTable().styleSettings.style = 'DarkStyle2';
            await context.sync();
          },
          read
        ),
      };
    },
  },
  {
    name: 'errors.not-implemented',
    area: 'errors',
    deck: 'slide-layouts',
    writes: true,
    async script(context) {
      context.presentation.slides.getItemAt(0).background.fill.setGradientFill({ type: 'Linear' });
      await context.sync();
      // `ShadeFromTitle` is a read state: the write refuses the sync, so nothing of it applies.
      return unchanged(
        context,
        async () => {
          const fill = context.presentation.slides.getItemAt(0).background.fill;
          fill.getGradientFillOrNullObject().type = 'ShadeFromTitle';
          await context.sync();
        },
        gradientOf
      );
    },
    readBack: gradientOf,
  },
  {
    name: 'errors.argument-at-sync',
    area: 'errors',
    deck: 'slide-layouts',
    writes: true,
    async script(context) {
      // A bad argument does not throw at the call. The sync refuses, and the good writes queued
      // before it and after it do not apply.
      const shapes = context.presentation.slides.getItemAt(1).shapes;
      let threw = false;
      const refused = await unchanged(context, async () => {
        shapes.addTextBox('Queued before', { left: 10, top: 10, width: 100, height: 20 });
        try {
          const bad = shapes.addGeometricShape('Blob' as never);
          // The refused call returns an object that takes more calls.
          bad.name = 'Never';
          bad.textFrame.textRange.text = 'Never';
        } catch {
          threw = true;
        }
        shapes.addTextBox('Queued after', { left: 10, top: 40, width: 100, height: 20 });
        await context.sync();
      });
      // The same proxies work in the next batch.
      shapes.addTextBox('Retried', { left: 10, top: 70, width: 100, height: 20 });
      await context.sync();
      return { threw, refused, after: await state(context) };
    },
    readBack: state,
  },
  {
    name: 'errors.group-placeholder',
    area: 'errors',
    deck: 'slide-layouts',
    writes: false,
    async script(context) {
      return unchanged(context, async () => {
        const shapes = context.presentation.slides.getItemAt(1).shapes;
        const box = shapes.addGeometricShape('Rectangle', {
          left: 10,
          top: 10,
          width: 20,
          height: 20,
        });
        shapes.addGroup([box, '2']);
        await context.sync();
      });
    },
  },
  {
    name: 'errors.enum-suggestion',
    area: 'errors',
    deck: 'slide-layouts',
    writes: false,
    async script(context) {
      // An add-in guesses enum names; the refusal names the accepted value it meant.
      const hint = async (step: () => void): Promise<Json> => {
        try {
          step();
          await context.sync();
        } catch (error) {
          const e = error as { code?: string; debugInfo?: { innerError?: { message?: string } } };
          const message = e.debugInfo?.innerError?.message ?? '';
          return [e.code ?? 'error', /Did you mean '([^']+)'/.exec(message)?.[1] ?? null];
        }
        return 'applied';
      };
      const body = () => note(context).textFrame.textRange.paragraphFormat;
      return {
        bullet: await hint(() => (body().bulletFormat.style = 'ArabicPeriod' as never)),
        shape: await hint(() => {
          context.presentation.slides.getItemAt(1).shapes.addGeometricShape('Rect' as never);
        }),
        align: await hint(() => (body().horizontalAlignment = 'Centre' as never)),
        far: await hint(() => (body().horizontalAlignment = 'Sideways' as never)),
        unchanged: await state(context),
      };
    },
  },
  {
    name: 'errors.bindings',
    area: 'errors',
    deck: 'slide-layouts',
    writes: false,
    async script(context) {
      const bindings = context.presentation.bindings;
      const box = context.presentation.slides.getItemAt(1).shapes.getItem('2');
      bindings.add(box, 'Shape', 'kept');
      await context.sync();
      return {
        missing: await unchanged(context, async () => {
          bindings.getItem('none').getShape().load('name');
          await context.sync();
        }),
        position: await unchanged(context, async () => {
          bindings.getItemAt(5).load('id');
          await context.sync();
        }),
        // A bad id refuses the whole sync: the binding queued before it is not added.
        badId: await unchanged(context, async () => {
          bindings.add(box, 'Shape', 'never');
          bindings.add(box, 'Shape', '');
          await context.sync();
        }),
        badType: await unchanged(context, async () => {
          bindings.add(box, 'Text' as never, 'never');
          await context.sync();
        }),
        ids: await bindingIds(context),
      };
    },
  },
];
