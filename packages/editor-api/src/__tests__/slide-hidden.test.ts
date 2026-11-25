/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/editor-api/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// `Slide.hidden`, an SDK extension over `p:sld/@show` (ECMA-376 Part 1, §19.3.1.38; `CT_Slide`
// declares `show` as an optional `xsd:boolean` with default `true`).

import { describe, expect, test } from 'bun:test';
import { createPptxEditor, type ImageRasterRequest } from '@pptx-editor.dev/core/editor';
import { attr, presentationInfo, readPptx, xmlPartOf } from '@pptx-editor.dev/core/store';
import { buildPptx, para, run, shape } from '../../../../scripts/fixtures/builder.ts';
import { PptxEditor as BrowserEditor } from '../browser.ts';
import {
  isPptxEditorError,
  PptxEditor,
  type PptxEditorRuntime,
  type RequestContext,
} from '../index.ts';
import { open, refusedAtSync, reopen } from './support.ts';

/** Three slides; the second is hidden in the file. */
function withHidden(): Uint8Array {
  return buildPptx({
    slides: [1, 2, 3].map((n) => ({
      layout: 'blank' as const,
      hidden: n === 2,
      shapes: shape({
        id: 2,
        name: 'Box',
        x: 914_400,
        y: 914_400,
        cx: 914_400,
        cy: 914_400,
        text: para(run(`Slide ${n}`)),
      }),
    })),
  });
}

/** The PNG signature and an `IHDR` chunk: enough for the size check. */
function pngOf(width: number, height: number): Uint8Array {
  const b = new Uint8Array(33);
  b.set([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a, 0, 0, 0, 13, 0x49, 0x48, 0x44, 0x52]);
  const view = new DataView(b.buffer);
  view.setUint32(16, width);
  view.setUint32(20, height);
  b.set([8, 6, 0, 0, 0], 24);
  return b;
}

/** The `show` attribute of each slide root, in presentation order. */
function showOf(bytes: Uint8Array): (string | undefined)[] {
  const read = readPptx(bytes);
  if (!read.ok) throw new Error('the saved presentation does not read back');
  return presentationInfo(read.pkg).slides.map(
    (s) => attr(xmlPartOf(read.pkg, s.partName)!.root, 'show') ?? undefined
  );
}

/** Load `hidden` of every slide. */
async function hiddenOf(context: RequestContext): Promise<boolean[]> {
  const slides = context.presentation.slides.load('items/hidden');
  await context.sync();
  return slides.items.map((s) => s.hidden);
}

describe('Slide.hidden', () => {
  test('a load reads the saved state', async () => {
    const runtime = await open(withHidden());
    const read = await runtime.run(async (context) => {
      const second = context.presentation.slides.getItemAt(1).load('hidden');
      await context.sync();
      return [second.hidden, await hiddenOf(context)] as const;
    });
    expect(read).toEqual([true, [false, true, false]]);
    runtime.dispose();
  });

  test('a load of every scalar includes hidden', async () => {
    const runtime = await open(withHidden());
    const json = await runtime.run(async (context) => {
      const slide = context.presentation.slides.getItemAt(1).load();
      await context.sync();
      return slide.toJSON();
    });
    expect(json).toMatchObject({ hidden: true });
    runtime.dispose();
  });

  test('a write hides and shows slides through p:sld/@show, and survives save and reopen', async () => {
    const runtime = await open(withHidden());
    await runtime.run(async (context) => {
      context.presentation.slides.getItemAt(0).hidden = true;
      context.presentation.slides.getItemAt(1).hidden = false;
      await context.sync();
      expect(await hiddenOf(context)).toEqual([true, false, false]);
    });
    const bytes = await runtime.save();
    // A hidden slide writes `show="0"`; a shown slide drops the attribute, whose default is true.
    expect(showOf(bytes)).toEqual(['0', undefined, undefined]);
    const again = await reopen(runtime);
    expect(await again.run(hiddenOf)).toEqual([true, false, false]);
    runtime.dispose();
    again.dispose();
  });

  test('a write of the current state changes nothing', async () => {
    const runtime = await open(withHidden());
    const before = await runtime.save();
    await runtime.run(async (context) => {
      context.presentation.slides.getItemAt(1).hidden = true;
      context.presentation.slides.getItemAt(2).hidden = false;
      await context.sync();
    });
    expect(showOf(await runtime.save())).toEqual(showOf(before));
    runtime.dispose();
  });

  test('a value that is not a boolean refuses with InvalidArgument at the sync', async () => {
    const runtime = await open(withHidden());
    const error = await runtime.run((context) =>
      refusedAtSync(context, () => {
        (context.presentation.slides.getItemAt(0) as { hidden: unknown }).hidden = 'yes';
      })
    );
    expect(isPptxEditorError(error) && error.code).toBe('InvalidArgument');
    expect(await runtime.run(hiddenOf)).toEqual([false, true, false]);
    runtime.dispose();
  });

  test('a write to a slide that does not exist refuses with ItemNotFound', async () => {
    const runtime = await open(withHidden());
    const error = await runtime.run((context) =>
      refusedAtSync(context, () => {
        context.presentation.slides.getItem('9999').hidden = true;
      })
    );
    expect(isPptxEditorError(error) && error.code).toBe('ItemNotFound');
    runtime.dispose();
  });

  test('the browser runtime reads and writes as the server runtime does, as one undo step', async () => {
    const live = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    expect((await live.load(withHidden())).ok).toBe(true);
    const runtime = BrowserEditor.createBrowser(live);
    expect(await runtime.run(hiddenOf)).toEqual([false, true, false]);
    await runtime.run(async (context) => {
      context.presentation.slides.getItemAt(0).hidden = true;
      context.presentation.slides.getItemAt(1).hidden = false;
      await context.sync();
    });
    expect(live.snapshot().slides.map((s) => s.hidden)).toEqual([true, false, false]);
    expect(showOf(live.save()!)).toEqual(['0', undefined, undefined]);
    expect(live.exec({ type: 'undo' }).ok).toBe(true);
    expect(live.snapshot().slides.map((s) => s.hidden)).toEqual([false, true, false]);
    expect(live.snapshot().canUndo).toBe(false);
  });

  test('getImageAsBase64 draws a hidden slide', async () => {
    const requests: ImageRasterRequest[] = [];
    const runtime = await PptxEditor.createServer(withHidden(), {
      substitutes: false,
      rasterize: (request) => {
        requests.push(request);
        return pngOf(request.width, request.height);
      },
    });
    const image = await runtime.run(async (context) => {
      const result = context.presentation.slides.getItemAt(1).getImageAsBase64({ width: 320 });
      await context.sync();
      return result.value;
    });
    expect(image.length).toBeGreaterThan(0);
    // The picture is of the slide itself, whether a slide show skips it or not.
    expect(requests).toHaveLength(1);
    expect(requests[0]!.svg).toContain('data-slide-part="/ppt/slides/slide2.xml"');
    runtime.dispose();
  });

  test('the example of the slides and shapes page runs as written', async () => {
    // docs/site/content/editor-api/slides-and-shapes.mdx, "Hide and show slides".
    async function hideSlides(runtime: PptxEditorRuntime, positions: number[]) {
      return runtime.run(async (context) => {
        const slides = context.presentation.slides;
        for (const position of positions) slides.getItemAt(position).hidden = true;
        await context.sync();
        slides.load('items/id,items/hidden');
        await context.sync();
        return slides.items.filter((slide) => slide.hidden).map((slide) => slide.id);
      });
    }
    const runtime = await open(withHidden());
    const ids = await runtime.run(async (context) => {
      const slides = context.presentation.slides.load('items/id');
      await context.sync();
      return slides.items.map((slide) => slide.id);
    });
    expect(await hideSlides(runtime, [0, 1])).toEqual([ids[0], ids[1]]);
    runtime.dispose();
  });
});
