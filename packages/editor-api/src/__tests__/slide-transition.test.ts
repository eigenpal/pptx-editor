/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/editor-api/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// `Slide.transition`, an SDK extension over `p:transition` (`CT_SlideTransition`, ECMA-376
// Part 1, §19.3.1.50) and its markup-compatibility branches (Part 3, §7.5 `AlternateContent`).

import { describe, expect, test } from 'bun:test';
import type { AutomationHost, AutomationOperation } from '@pptx-editor.dev/core/automation';
import { createBrowserAutomationHost, createPptxEditor } from '@pptx-editor.dev/core/editor';
import {
  presentationInfo,
  readPptx,
  serializeXmlPart,
  xmlPartOf,
} from '@pptx-editor.dev/core/store';
import { buildPptx, para, run, shape } from '../../../../scripts/fixtures/builder.ts';
import { PptxEditor as BrowserEditor } from '../browser.ts';
import {
  isPptxEditorError,
  SlideTransitionDirection,
  SlideTransitionEffect,
  type RequestContext,
} from '../index.ts';
import { createRuntime } from '../runtime/runtime.ts';
import { open, refusedAtSync, reopen } from './support.ts';

const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
const P14 = 'http://schemas.microsoft.com/office/powerpoint/2010/main';
const P159 = 'http://schemas.microsoft.com/office/powerpoint/2015/09/main';

/** A fade with a 1200 ms duration and a 3000 ms advance, as a `p14` choice and a fallback. */
const FADE =
  `<mc:AlternateContent xmlns:mc="${MC}"><mc:Choice xmlns:p14="${P14}" Requires="p14">` +
  '<p:transition spd="slow" p14:dur="1200" advTm="3000"><p:fade/></p:transition></mc:Choice>' +
  '<mc:Fallback><p:transition spd="slow" advTm="3000"><p:fade/></p:transition></mc:Fallback>' +
  '</mc:AlternateContent>';
/** A standard effect that the API does not write, with no duration. */
const DISSOLVE = '<p:transition spd="slow" advClick="0"><p:dissolve/></p:transition>';
/** An effect from an extension namespace, with a fade fallback. */
const MORPH =
  `<mc:AlternateContent xmlns:mc="${MC}"><mc:Choice xmlns:p159="${P159}" Requires="p159">` +
  '<p:transition spd="slow"><p159:morph option="byObject"/></p:transition></mc:Choice>' +
  '<mc:Fallback><p:transition spd="slow"><p:fade/></p:transition></mc:Fallback></mc:AlternateContent>';

/** Four slides: no transition, a fade, a dissolve, and a morph. */
function withTransitions(): Uint8Array {
  return buildPptx({
    slides: ['', FADE, DISSOLVE, MORPH].map((ext, n) => ({
      layout: 'blank' as const,
      ext,
      shapes: shape({
        id: 2,
        name: 'Box',
        x: 914_400,
        y: 914_400,
        cx: 914_400,
        cy: 914_400,
        text: para(run(`Slide ${n + 1}`)),
      }),
    })),
  });
}

/** The XML of each slide part, in presentation order. */
function slideXml(bytes: Uint8Array): string[] {
  const read = readPptx(bytes);
  if (!read.ok) throw new Error('the saved presentation does not read back');
  return presentationInfo(read.pkg).slides.map((s) =>
    serializeXmlPart(xmlPartOf(read.pkg, s.partName)!)
  );
}

interface Read {
  effect: string;
  durationMs: number;
  advanceOnClick: boolean;
  advanceAfterMs: number | null;
}

/** Load the transition of every slide. */
async function transitionsOf(context: RequestContext): Promise<Read[]> {
  const slides = context.presentation.slides.load('items/transition');
  await context.sync();
  return slides.items.map((s) => {
    const t = s.transition;
    return {
      effect: t.effect,
      durationMs: t.durationMs,
      advanceOnClick: t.advanceOnClick,
      advanceAfterMs: t.advanceAfterMs,
    };
  });
}

/** Load the direction and the through-black option of every slide. */
async function variantsOf(context: RequestContext): Promise<[string | null, boolean][]> {
  const slides = context.presentation.slides.load(
    'items/transition/direction,items/transition/throughBlack'
  );
  await context.sync();
  return slides.items.map((s) => [s.transition.direction, s.transition.throughBlack]);
}

const NONE: Read = { effect: 'None', durationMs: 750, advanceOnClick: true, advanceAfterMs: null };

describe('Slide.transition', () => {
  test('a load reads the effect, the duration, and the advance settings', async () => {
    const runtime = await open(withTransitions());
    expect(await runtime.run(transitionsOf)).toEqual([
      NONE,
      { effect: 'Fade', durationMs: 1200, advanceOnClick: true, advanceAfterMs: 3000 },
      // A standard effect reads by its element name; an extension effect reads `Other`.
      { effect: 'Dissolve', durationMs: 1000, advanceOnClick: false, advanceAfterMs: null },
      { effect: 'Other', durationMs: 1000, advanceOnClick: true, advanceAfterMs: null },
    ]);
    const json = await runtime.run(async (context) => {
      const transition = context.presentation.slides.getItemAt(1).transition.load('effect');
      await context.sync();
      return transition.toJSON();
    });
    expect(json).toEqual({ effect: 'Fade' });
    runtime.dispose();
  });

  test('set writes every setting in one step, and survives save and reopen', async () => {
    const runtime = await open(withTransitions());
    await runtime.run(async (context) => {
      context.presentation.slides.getItemAt(0).transition.set({
        effect: SlideTransitionEffect.push,
        durationMs: 1500,
        advanceOnClick: false,
        advanceAfterMs: 5000,
      });
      await context.sync();
    });
    const xml = slideXml(await runtime.save())[0]!;
    // The `p14` choice holds the duration; the fallback holds the nearest speed.
    expect(xml).toContain('Requires="p14"');
    expect(xml).toContain('p14:dur="1500"');
    expect(xml).toContain(
      '<mc:Fallback><p:transition spd="slow" advClick="0" advTm="5000"><p:push dir="u"/>'
    );
    // `CT_Slide` puts the transition after `p:clrMapOvr`.
    expect(xml.indexOf('<mc:AlternateContent')).toBeGreaterThan(xml.indexOf('</p:clrMapOvr>'));
    const again = await reopen(runtime);
    expect((await again.run(transitionsOf))[0]).toEqual({
      effect: 'Push',
      durationMs: 1500,
      advanceOnClick: false,
      advanceAfterMs: 5000,
    });
    runtime.dispose();
    again.dispose();
  });

  test('property writes queue one write each, in the order of the calls', async () => {
    const runtime = await open(withTransitions());
    const read = await runtime.run(async (context) => {
      const transition = context.presentation.slides.getItemAt(0).transition;
      transition.effect = 'Wipe';
      transition.durationMs = 400;
      transition.advanceAfterMs = 2000;
      await context.sync();
      return (await transitionsOf(context))[0];
    });
    expect(read).toEqual({
      effect: 'Wipe',
      durationMs: 400,
      advanceOnClick: true,
      advanceAfterMs: 2000,
    });
    expect(slideXml(await runtime.save())[0]).toContain('spd="fast" p14:dur="400"');
    runtime.dispose();
  });

  test('a timing write keeps an effect that the API does not write, and its branches', async () => {
    const runtime = await open(withTransitions());
    await runtime.run(async (context) => {
      const slides = context.presentation.slides;
      slides.getItemAt(2).transition.advanceOnClick = true;
      slides.getItemAt(3).transition.set({ durationMs: 500, advanceAfterMs: 2000 });
      await context.sync();
    });
    const again = await reopen(runtime);
    const [, , dissolve, morph] = slideXml(await again.save());
    expect(dissolve).toContain('<p:dissolve/>');
    expect(dissolve).not.toContain('advClick');
    expect(morph).toContain(`<mc:Choice xmlns:p159="${P159}" Requires="p159">`);
    expect(morph).toContain('<p159:morph option="byObject"/>');
    expect(morph.match(/advTm="2000"/g)).toHaveLength(2);
    expect(morph.match(/spd="fast"/g)).toHaveLength(2);
    expect((await again.run(transitionsOf)).slice(2)).toEqual([
      { effect: 'Dissolve', durationMs: 1000, advanceOnClick: true, advanceAfterMs: null },
      { effect: 'Other', durationMs: 500, advanceOnClick: true, advanceAfterMs: 2000 },
    ]);
    runtime.dispose();
    again.dispose();
  });

  test('a load reads the direction and the through-black option', async () => {
    const bytes = buildPptx({
      slides: [
        '<p:transition><p:push/></p:transition>',
        '<p:transition><p:wipe dir="d"/></p:transition>',
        '<p:transition><p:split orient="vert" dir="in"/></p:transition>',
        '<p:transition><p:split/></p:transition>',
        '<p:transition><p:fade thruBlk="1"/></p:transition>',
        '<p:transition><p:zoom/></p:transition>',
        '',
      ].map((ext) => ({ layout: 'blank' as const, ext, shapes: '' })),
    });
    const runtime = await open(bytes);
    // Absent attributes read their schema defaults: `dir="l"`, `orient="horz"`, `dir="out"`.
    expect(await runtime.run(variantsOf)).toEqual([
      ['Left', false],
      ['Down', false],
      ['VerticalIn', false],
      ['HorizontalOut', false],
      [null, true],
      [null, false],
      [null, false],
    ]);
    expect((await runtime.run(transitionsOf))[5]!.effect).toBe('Zoom');
    runtime.dispose();
  });

  test('a direction or through-black write changes every copy, and keeps the timing', async () => {
    const runtime = await open(withTransitions());
    await runtime.run(async (context) => {
      const slides = context.presentation.slides;
      slides.getItemAt(0).transition.set({ effect: 'Push', direction: 'Left', durationMs: 900 });
      slides.getItemAt(1).transition.throughBlack = true;
      await context.sync();
      slides.getItemAt(0).transition.direction = SlideTransitionDirection.right;
      slides.getItemAt(3).transition.set({ effect: 'Split', direction: 'HorizontalIn' });
      await context.sync();
    });
    const again = await reopen(runtime);
    const xml = slideXml(await again.save());
    expect(xml[0]!.match(/<p:push dir="r"\/>/g)).toHaveLength(2);
    expect(xml[0]).toContain('p14:dur="900"');
    expect(xml[1]!.match(/<p:fade thruBlk="1"\/>/g)).toHaveLength(2);
    expect(xml[1]).toContain('p14:dur="1200"');
    // `orient="horz"` is the default, so only `dir="in"` is written.
    expect(xml[3]!.match(/<p:split dir="in"\/>/g)).toHaveLength(2);
    expect((await again.run(variantsOf)).slice(0, 4)).toEqual([
      ['Right', false],
      [null, true],
      [null, false],
      ['HorizontalIn', false],
    ]);
    runtime.dispose();
    again.dispose();
  });

  test('a direction or through-black write to another effect refuses with NotSupported', async () => {
    const runtime = await open(withTransitions());
    const before = slideXml(await runtime.save());
    const cases: ((context: RequestContext) => void)[] = [
      (c) => (c.presentation.slides.getItemAt(1).transition.direction = 'Left'),
      (c) => (c.presentation.slides.getItemAt(0).transition.throughBlack = true),
      (c) =>
        c.presentation.slides
          .getItemAt(1)
          .transition.set({ effect: 'Push', direction: 'VerticalIn' }),
      (c) => c.presentation.slides.getItemAt(3).transition.set({ direction: 'Up' }),
    ];
    for (const call of cases) {
      const error = await runtime.run((context) => refusedAtSync(context, () => call(context)));
      expect(isPptxEditorError(error) && error.code).toBe('NotSupported');
      expect(isPptxEditorError(error) && error.debugInfo.innerError?.code).toBe(
        'edit.transition-option-not-supported'
      );
    }
    expect(slideXml(await runtime.save())).toEqual(before);
    runtime.dispose();
  });

  test('a duration that an extension effect keeps only as a speed refuses', async () => {
    const runtime = await open(withTransitions());
    const before = slideXml(await runtime.save());
    const error = await runtime.run((context) =>
      refusedAtSync(context, () => {
        // The advance write of the same sync does not apply either.
        context.presentation.slides.getItemAt(3).transition.advanceAfterMs = 2000;
        context.presentation.slides.getItemAt(3).transition.durationMs = 700;
      })
    );
    expect(isPptxEditorError(error) && error.code).toBe('NotSupported');
    expect(isPptxEditorError(error) && error.debugInfo.innerError?.code).toBe(
      'edit.transition-speed-only'
    );
    expect(slideXml(await runtime.save())).toEqual(before);
    await runtime.run(async (context) => {
      context.presentation.slides.getItemAt(3).transition.durationMs = 750;
      // A plain `p:transition` gets a `p14` copy, so it keeps any duration.
      context.presentation.slides.getItemAt(2).transition.durationMs = 700;
      await context.sync();
    });
    const read = await runtime.run(transitionsOf);
    expect([read[2]!.durationMs, read[3]!.durationMs]).toEqual([700, 750]);
    expect(slideXml(await runtime.save())[2]).toContain('<p:dissolve/>');
    runtime.dispose();
  });

  test('remove clears the effect and the advance settings', async () => {
    const runtime = await open(withTransitions());
    await runtime.run(async (context) => {
      const slides = context.presentation.slides;
      for (const i of [0, 1, 2, 3]) slides.getItemAt(i).transition.remove();
      await context.sync();
    });
    const xml = slideXml(await runtime.save());
    for (const slide of xml) expect(slide).not.toContain('transition');
    expect(await runtime.run(transitionsOf)).toEqual([NONE, NONE, NONE, NONE]);
    runtime.dispose();
  });

  test('a new effect keeps the transition sound, and remove drops it', async () => {
    const SOUND = '<p:sndAc><p:endSnd/></p:sndAc>';
    const bytes = buildPptx({
      slides: [
        { layout: 'blank', shapes: '', ext: `<p:transition><p:fade/>${SOUND}</p:transition>` },
      ],
    });
    const runtime = await open(bytes);
    await runtime.run(async (context) => {
      context.presentation.slides.getItemAt(0).transition.set({ effect: 'Push', durationMs: 900 });
      await context.sync();
    });
    const again = await reopen(runtime);
    const xml = slideXml(await again.save())[0]!;
    // `p:sndAc` follows the effect element in each copy (`CT_SlideTransition`).
    expect(xml.match(new RegExp(`<p:push dir="u"/>${SOUND}`, 'g'))).toHaveLength(2);
    await again.run(async (context) => {
      context.presentation.slides.getItemAt(0).transition.remove();
      await context.sync();
    });
    expect(slideXml(await again.save())[0]).not.toContain('transition');
    runtime.dispose();
    again.dispose();
  });

  test('remove after None leaves no transition', async () => {
    const runtime = await open(withTransitions());
    await runtime.run(async (context) => {
      context.presentation.slides.getItemAt(1).transition.effect = 'None';
      await context.sync();
      context.presentation.slides.getItemAt(1).transition.remove();
      await context.sync();
    });
    expect(slideXml(await runtime.save())[1]).not.toContain('transition');
    expect((await runtime.run(transitionsOf))[1]).toEqual(NONE);
    runtime.dispose();
  });

  test('None removes the effect; with the default advance settings no transition stays', async () => {
    const runtime = await open(withTransitions());
    await runtime.run(async (context) => {
      context.presentation.slides.getItemAt(1).transition.effect = 'None';
      context.presentation.slides
        .getItemAt(3)
        .transition.set({ effect: 'None', advanceOnClick: true, advanceAfterMs: null });
      await context.sync();
    });
    const xml = slideXml(await runtime.save());
    // The advance time of the fade stays on a `p:transition` with no effect element.
    expect(xml[1]).toContain('<p:transition advTm="3000"/>');
    expect(xml[1]).not.toContain('fade');
    expect(xml[3]).not.toContain('transition');
    expect(xml[3]).not.toContain('AlternateContent');
    const again = await reopen(runtime);
    const read = await again.run(transitionsOf);
    // A `p:transition` with no `spd` reads the schema default speed, `fast` (500 ms).
    expect(read[1]).toEqual({ ...NONE, durationMs: 500, advanceAfterMs: 3000 });
    expect(read[3]).toEqual(NONE);
    runtime.dispose();
    again.dispose();
  });

  test('a write of the current settings changes nothing', async () => {
    const runtime = await open(withTransitions());
    const before = slideXml(await runtime.save());
    await runtime.run(async (context) => {
      const slides = context.presentation.slides;
      slides.getItemAt(0).transition.effect = 'None';
      slides
        .getItemAt(1)
        .transition.set({ effect: 'Fade', durationMs: 1200, advanceAfterMs: 3000 });
      slides.getItemAt(2).transition.advanceOnClick = false;
      await context.sync();
    });
    expect(slideXml(await runtime.save())).toEqual(before);
    runtime.dispose();
  });

  test('invalid arguments refuse with InvalidArgument at the sync, and nothing is written', async () => {
    const runtime = await open(withTransitions());
    const before = slideXml(await runtime.save());
    const cases: ((context: RequestContext) => void)[] = [
      (c) => (c.presentation.slides.getItemAt(0).transition.effect = 'Other'),
      (c) =>
        ((c.presentation.slides.getItemAt(0).transition as { effect: unknown }).effect = 'Spin'),
      (c) => (c.presentation.slides.getItemAt(1).transition.durationMs = 5),
      (c) => (c.presentation.slides.getItemAt(1).transition.durationMs = Number.NaN),
      (c) => (c.presentation.slides.getItemAt(1).transition.advanceAfterMs = -1),
      (c) =>
        ((
          c.presentation.slides.getItemAt(1).transition as { advanceOnClick: unknown }
        ).advanceOnClick = 'yes'),
      (c) => c.presentation.slides.getItemAt(1).transition.set({ speed: 'fast' } as never),
      (c) => c.presentation.slides.getItemAt(1).transition.set(null as never),
    ];
    for (const call of cases) {
      const error = await runtime.run((context) =>
        refusedAtSync(context, () => {
          // A valid write first: the refusal must stop the whole sync.
          context.presentation.slides.getItemAt(2).transition.effect = 'Cut';
          call(context);
        })
      );
      expect(isPptxEditorError(error) && error.code).toBe('InvalidArgument');
    }
    expect(slideXml(await runtime.save())).toEqual(before);
    runtime.dispose();
  });

  test('a duration on a slide with no effect refuses with NotSupported, and nothing is written', async () => {
    const runtime = await open(withTransitions());
    const before = slideXml(await runtime.save());
    const error = await runtime.run((context) =>
      refusedAtSync(context, () => {
        context.presentation.slides.getItemAt(1).transition.effect = 'Cut';
        context.presentation.slides.getItemAt(0).transition.durationMs = 500;
      })
    );
    expect(isPptxEditorError(error) && error.code).toBe('NotSupported');
    expect(isPptxEditorError(error) && error.debugInfo.innerError?.code).toBe(
      'edit.transition-no-effect'
    );
    expect(slideXml(await runtime.save())).toEqual(before);
    // The same duration with an effect in one `set` call applies.
    await runtime.run(async (context) => {
      context.presentation.slides.getItemAt(0).transition.set({ durationMs: 500, effect: 'Cut' });
      await context.sync();
    });
    expect((await runtime.run(transitionsOf))[0]).toMatchObject({ effect: 'Cut', durationMs: 500 });
    runtime.dispose();
  });

  test('a write to a slide that does not exist refuses with ItemNotFound', async () => {
    const runtime = await open(withTransitions());
    const error = await runtime.run((context) =>
      refusedAtSync(context, () => {
        context.presentation.slides.getItem('9999').transition.effect = 'Fade';
      })
    );
    expect(isPptxEditorError(error) && error.code).toBe('ItemNotFound');
    runtime.dispose();
  });

  test('the browser runtime reads and writes as the server runtime does, as one undo step', async () => {
    const live = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    expect((await live.load(withTransitions())).ok).toBe(true);
    const runtime = BrowserEditor.createBrowser(live);
    const server = await open(withTransitions());
    expect(await runtime.run(transitionsOf)).toEqual(await server.run(transitionsOf));
    const write = async (context: RequestContext) => {
      const slides = context.presentation.slides.load('items');
      await context.sync();
      for (const slide of slides.items) slide.transition.set({ effect: 'Split', durationMs: 900 });
      slides.items[0]!.transition.advanceAfterMs = 4000;
      await context.sync();
    };
    await runtime.run(write);
    await server.run(write);
    expect(await runtime.run(transitionsOf)).toEqual(await server.run(transitionsOf));
    expect(slideXml(live.save()!)).toEqual(slideXml(await server.save()));
    expect(live.snapshot().animations!.transition).toEqual({
      kind: 'split',
      durationMs: 900,
      advanceOnClick: true,
      advanceAfterMs: 4000,
    });
    expect(live.exec({ type: 'undo' }).ok).toBe(true);
    expect(slideXml(live.save()!)).toEqual(slideXml(withTransitions()));
    expect(live.snapshot().canUndo).toBe(false);
    server.dispose();
  });

  test('a write sends the lookups and one setTransition operation', async () => {
    // The collaboration test (`packages/pro/src/__tests__/yjs-slide-transition.test.ts`) sends
    // these operations to two linked editors; this test keeps them equal to what the API sends.
    const live = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    expect((await live.load(withTransitions())).ok).toBe(true);
    const inner = createBrowserAutomationHost(live);
    const batches: AutomationOperation[][] = [];
    const host: AutomationHost = {
      ...inner,
      execute: (request) => {
        batches.push([...request.operations]);
        return inner.execute(request);
      },
    };
    await createRuntime({ host, save: false }).run(async (context) => {
      context.presentation.slides.getItemAt(0).transition.set({ effect: 'Fade', durationMs: 700 });
      await context.sync();
    });
    expect(batches).toEqual([
      [
        { op: 'getPresentation' },
        { op: 'getSlide', key: { index: 0 } },
        {
          op: 'setTransition',
          slide: { result: 1 },
          transition: { kind: 'fade', durationMs: 700 },
        },
      ],
    ]);
  });
});
