/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/editor-api/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Presentation.insertSlidesFromBase64 end to end on a server runtime, with a save and reopen.

import { describe, expect, test } from 'bun:test';
import { buildPptx, para, run } from '../../../../scripts/fixtures/builder.ts';
import { InsertSlideFormatting, isPptxEditorError, type RequestContext } from '../index.ts';
import { readZip, writeZip } from '@pptx-editor.dev/core/store';
import { diagnosticOf } from '../runtime/errors.ts';
import { open, refusedAtSync, rejection, reopen } from './support.ts';

const code = (error: unknown): string | undefined =>
  isPptxEditorError(error) ? error.code : String(error);
const inner = (error: unknown): string | undefined =>
  isPptxEditorError(error) ? error.debugInfo.innerError?.code : undefined;

const titled = (text: string) =>
  `<p:sp><p:nvSpPr><p:cNvPr id="2" name="Title 1"/><p:cNvSpPr/><p:nvPr><p:ph type="title"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/>${para(run(text))}</p:txBody></p:sp>`;

/** Three titled slides with notes on the first, as Base64. */
function source(): string {
  const bytes = buildPptx({
    slides: [
      { layout: 'title', shapes: titled('S1'), notes: 'Speaker notes' },
      { layout: 'blank', shapes: titled('S2') },
      { layout: 'title', shapes: titled('S3') },
    ],
  });
  let s = '';
  for (const b of bytes) s += String.fromCharCode(b);
  return btoa(s);
}

/** The title text of every slide, in order. */
async function titles(context: RequestContext): Promise<string[]> {
  const slides = context.presentation.slides;
  slides.load('items');
  await context.sync();
  const ranges = slides.items.map((slide) => {
    const shapes = slide.shapes;
    shapes.load('items/name');
    return shapes;
  });
  await context.sync();
  const texts = ranges.map((shapes) => {
    const title = shapes.items.find((s) => s.name === 'Title 1');
    return title ? title.textFrame.textRange.load('text') : null;
  });
  await context.sync();
  return texts.map((t) => (t ? t.text : ''));
}

describe('insertSlidesFromBase64', () => {
  test('every slide goes first by default and survives a save', async () => {
    const runtime = await open();
    await runtime.run(async (context) => {
      context.presentation.insertSlidesFromBase64(source());
      await context.sync();
    });
    const again = await reopen(runtime);
    await again.run(async (context) => {
      expect(await titles(context)).toEqual(['S1', 'S2', 'S3', '', '']);
      const masters = context.presentation.slideMasters;
      masters.load('items');
      await context.sync();
      expect(masters.items).toHaveLength(2);
    });
  });

  test('sourceSlideIds, targetSlideId, and the destination theme', async () => {
    const runtime = await open();
    await runtime.run(async (context) => {
      const first = context.presentation.slides.getItemAt(0).load('id');
      await context.sync();
      context.presentation.insertSlidesFromBase64(source(), {
        formatting: InsertSlideFormatting.useDestinationTheme,
        sourceSlideIds: ['258', '256'],
        targetSlideId: first.id,
      });
      await context.sync();
      expect(await titles(context)).toEqual(['', 'S3', 'S1', '']);
      const masters = context.presentation.slideMasters;
      masters.load('items');
      await context.sync();
      expect(masters.items).toHaveLength(1);
    });
  });

  test('the string formatting values work as the enum does', async () => {
    const runtime = await open();
    await runtime.run(async (context) => {
      context.presentation.insertSlidesFromBase64(source(), {
        formatting: 'KeepSourceFormatting',
        sourceSlideIds: ['257'],
      });
      await context.sync();
      expect(await titles(context)).toEqual(['S2', '', '']);
    });
  });

  test('bad arguments refuse at the sync, and change nothing', async () => {
    const runtime = await open();
    await runtime.run(async (context) => {
      const presentation = context.presentation;
      const target = 'presentation.insertSlidesFromBase64';
      expect(
        code(await refusedAtSync(context, () => presentation.insertSlidesFromBase64('')))
      ).toBe('InvalidArgument');
      expect(
        code(await refusedAtSync(context, () => presentation.insertSlidesFromBase64(1 as never)))
      ).toBe('InvalidArgument');
      const bad = await refusedAtSync(context, () =>
        presentation.insertSlidesFromBase64(source(), { formatting: 'Other' as never })
      );
      expect(isPptxEditorError(bad) && bad.debugInfo.errorLocation).toBe(`${target}.formatting`);
      expect(
        code(
          await refusedAtSync(context, () =>
            presentation.insertSlidesFromBase64(source(), { sourceSlideIds: [1 as never] })
          )
        )
      ).toBe('InvalidArgument');
      // Resolve the presentation first: a refused batch also refuses the lookups in it.
      presentation.slides.getCount();
      await context.sync();
      presentation.insertSlidesFromBase64('bm90IGEgcHB0eA==');
      let error = await rejection(context.sync());
      expect([code(error), inner(error)]).toEqual(['InvalidArgument', 'import.not-a-presentation']);
      presentation.insertSlidesFromBase64(source(), { sourceSlideIds: ['999'] });
      error = await rejection(context.sync());
      expect([code(error), inner(error)]).toEqual(['ItemNotFound', 'import.slide-not-found']);
      presentation.insertSlidesFromBase64(source(), { targetSlideId: '999' });
      error = await rejection(context.sync());
      expect([code(error), inner(error)]).toEqual(['ItemNotFound', 'import.target-not-found']);
      const count = presentation.slides.getCount();
      await context.sync();
      expect(count.value).toBe(2);
    });
  });
});

describe('slides an import cannot copy', () => {
  test('the refusal names the slides and points at sourceSlideIds', async () => {
    // Slides 257 and 258 of the source get an embedded control.
    const bytes = Uint8Array.from(atob(source()), (c) => c.charCodeAt(0));
    const zip = readZip(bytes);
    if (!zip.ok) throw new Error('fixture');
    const entries = new Map(zip.entries);
    for (const n of [2, 3]) {
      const name = [...entries.keys()].find((key) =>
        key.endsWith(`slides/_rels/slide${n}.xml.rels`)
      )!;
      const rels = new TextDecoder().decode(entries.get(name));
      entries.set(
        name,
        new TextEncoder().encode(
          rels.replace(
            '</Relationships>',
            '<Relationship Id="rIdX" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/control" Target="../media/x.bin"/></Relationships>'
          )
        )
      );
    }
    entries.set('/ppt/media/x.bin', new Uint8Array([1]));
    const types = [...entries.keys()].find((key) => key.endsWith('[Content_Types].xml'))!;
    const ct = new TextDecoder().decode(entries.get(types));
    entries.set(
      types,
      new TextEncoder().encode(
        ct.replace(
          '</Types>',
          '<Default Extension="bin" ContentType="application/octet-stream"/></Types>'
        )
      )
    );
    let s = '';
    for (const b of writeZip(entries)) s += String.fromCharCode(b);
    const withControls = btoa(s);

    const runtime = await open();
    const error = await rejection(
      runtime.run(async (context) => {
        context.presentation.insertSlidesFromBase64(withControls);
        await context.sync();
      })
    );
    expect([code(error), inner(error)]).toEqual(['NotSupported', 'import.slide-has-controls']);
    expect(isPptxEditorError(error) && error.debugInfo.innerError?.message).toBe(
      'Slides 257, 258 of the source have an embedded control (p:control). Leave them out with ' +
        'sourceSlideIds.'
    );
    await runtime.run(async (context) => {
      context.presentation.insertSlidesFromBase64(withControls, { sourceSlideIds: ['256'] });
      await context.sync();
    });
    await reopen(runtime);
  });

  test('a diagnostic for threaded comments names the slides', () => {
    expect(diagnosticOf('disabled.slideNotCopyable:import-comments:257')).toEqual({
      code: 'import.slide-has-comments',
      hint:
        'Slide 257 of the source has threaded comments, which an import cannot copy yet. Leave ' +
        'it out with sourceSlideIds.',
    });
    // Only numbers reach the hint: other text selects the general diagnostic of the reason.
    expect(diagnosticOf('disabled.slideNotCopyable:import-comments:<b>')).toEqual({
      code: 'import.slide-not-copyable',
      hint: expect.not.stringContaining('<b>'),
    });
  });
});
