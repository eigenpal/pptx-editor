/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/editor-api/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// The examples of docs/site/content/editor-api, run as written against a synthetic deck.

import { readFileSync } from 'node:fs';
import { describe, expect, spyOn, test } from 'bun:test';
import { createCanvas, GlobalFonts, loadImage } from '@napi-rs/canvas';
import { buildPptx, para, run, shape } from '../../../../scripts/fixtures/builder.ts';
import {
  isPptxEditorError,
  PptxEditor,
  ShapeZOrder,
  type PptxEditorRuntime,
  type RequestContext,
} from '../index.ts';
import { deck, IN, rejection, reopen } from './support.ts';

test('the overview example edits and saves', async () => {
  const runtime = await PptxEditor.createServer(deck(), { substitutes: false });
  try {
    await runtime.run(async (context) => {
      const slide = context.presentation.slides.getItemAt(0);
      const box = slide.shapes.addTextBox('Quarterly results', { left: 40, top: 40, width: 400 });
      box.textFrame.textRange.font.bold = true;
      const table = slide.shapes
        .addTable(2, 2, {
          values: [
            ['Region', 'Total'],
            ['North', '42'],
          ],
        })
        .getTable();
      await context.sync();

      const cell = table.getCellOrNullObject(1, 1);
      cell.horizontalAlignment = 'Right';
      cell.fill.setSolidColor('#FFF2CC');
      await context.sync();
    });
    const again = await reopen(runtime);
    await again.run(async (context) => {
      const shapes = context.presentation.slides.getItemAt(0).shapes.load('items/name,items/type');
      await context.sync();
      expect(shapes.items.map((s) => s.type)).toContain('Table');
    });
  } finally {
    runtime.dispose();
  }
});

test('the compatibility example renames shapes, and enums take both spellings', async () => {
  const runtime = await PptxEditor.createServer(deck(), { substitutes: false });
  await runtime.run(async (context) => {
    const shapes = context.presentation.slides.getItemAt(0).shapes;
    shapes.load('items/name');
    await context.sync();
    for (const shape of shapes.items) shape.name = shape.name.toUpperCase();
    shapes.items[0]!.setZOrder(ShapeZOrder.bringToFront);
    shapes.items[1]!.setZOrder('BringToFront');
    await context.sync();
    shapes.load('items/name');
    await context.sync();
    expect(shapes.items.map((s) => s.name)).toEqual(['BOX', 'OTHER']);
  });
  runtime.dispose();
});

test('the common mistakes examples behave as the docs say', async () => {
  const runtime = await PptxEditor.createServer(deck(), { substitutes: false });
  try {
    await runtime.run(async (context) => {
      const slide = context.presentation.slides.getItemAt(0);
      // A read with no load.
      const box = slide.shapes.getItem('2');
      expect(() => box.name).toThrow();
      box.load('name');
      await context.sync();
      expect(box.name).toBe('Box');
      // A sync inside a loop: one sync for all the writes.
      const shapes = slide.shapes.load('items/name');
      await context.sync();
      for (const shape of shapes.items) shape.name = shape.name.trim();
      await context.sync();
      // A retry after a refusal: the same proxy works.
      let refused = false;
      try {
        box.fill.setSolidColor('notacolor');
        await context.sync();
      } catch {
        refused = true;
        box.fill.setSolidColor('#1F4E79');
        await context.sync();
      }
      expect(refused).toBe(true);
      const fill = box.fill.load('foregroundColor');
      await context.sync();
      expect(fill.foregroundColor).toBe('#1F4E79');
      // Bindings are in memory only.
      context.presentation.bindings.add(box, 'Shape', 'kpi');
      await context.sync();
    });
    const again = await reopen(runtime);
    await again.run(async (context) => {
      const all = context.presentation.bindings.load('items');
      await context.sync();
      expect(all.items).toEqual([]);
    });
  } finally {
    runtime.dispose();
  }
});

/**
 * The `rasterize` example of docs/site/content/editor-api/runtime.mdx, as written. Keep the two
 * equal: the next test compares them.
 */
async function openWithImages(bytes: Uint8Array) {
  return PptxEditor.createServer(bytes, {
    rasterize: async ({ svg, width, height, fonts }) => {
      // Register the faces for this picture only, and remove them before returning.
      const keys = fonts.map((font) => GlobalFonts.register(Buffer.from(font.bytes), font.family));
      try {
        const canvas = createCanvas(width, height);
        canvas.getContext('2d').drawImage(await loadImage(Buffer.from(svg)), 0, 0, width, height);
        return new Uint8Array(canvas.toBuffer('image/png'));
      } finally {
        for (const key of keys) if (key) GlobalFonts.remove(key);
      }
    },
  });
}

test('the rasterize example of the test equals the example of runtime.mdx', () => {
  // The lines from `return PptxEditor.createServer(` to the end of its call, without indentation.
  const body = (source: string): string[] => {
    const lines = source.split('\n').map((line) => line.trim());
    const start = lines.indexOf('return PptxEditor.createServer(bytes, {');
    const end = lines.indexOf('});', start);
    expect(start).toBeGreaterThan(-1);
    expect(end).toBeGreaterThan(start);
    return lines.slice(start, end + 1);
  };
  const guide = readFileSync(
    new URL('../../../../docs/site/content/editor-api/runtime.mdx', import.meta.url),
    'utf8'
  );
  const own = readFileSync(new URL(import.meta.url), 'utf8');
  expect(body(own)).toEqual(body(guide));
  const canvasImport = "import { createCanvas, GlobalFonts, loadImage } from '@napi-rs/canvas';";
  // A whole line: the literal above is part of a longer line, so it does not match.
  expect(own.split('\n')).toContain(canvasImport);
  expect(guide.split('\n')).toContain(canvasImport);
  // The minimum release that the guide states is the floor of the optional peer range.
  const manifest = JSON.parse(
    readFileSync(new URL('../../package.json', import.meta.url), 'utf8')
  ) as { peerDependencies: Record<string, string> };
  const floor = manifest.peerDependencies['@napi-rs/canvas']!.replace(/^\^/, '');
  expect(guide).toContain(`Install \`@napi-rs/canvas\` ${floor} or a later 1.x release`);
  expect(manifest.peerDependencies['@napi-rs/canvas']).toBe(`^${floor}`);
});

test('the runtime example draws the embedded pictures of a slide', async () => {
  // A red PNG in a picture (`p:pic`, ECMA-376 Part 1, 19.3.1.37; its `a:blip`, 20.1.8.13) two
  // inches by one inch at one inch from the top left corner. The engine paints it as a data URL
  // in an SVG `<image>`. `loadImage` of `@napi-rs/canvas` draws it from 1.0.12 only: a probe
  // that draws a red data URL `<image>` over a blue SVG background leaves the pixel blue in
  // 0.1.100 and 1.0.0 to 1.0.11, and draws it red in 1.0.12 and later. The lockfile installs
  // a later release, so this test checks the engine output, not the floor itself.
  const red = createCanvas(8, 8);
  const pen = red.getContext('2d');
  pen.fillStyle = '#FF0000';
  pen.fillRect(0, 0, 8, 8);
  const bytes = buildPptx({
    slides: [
      {
        layout: 'blank',
        rels: [
          [
            'rIdImg',
            'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
            '../media/image1.png',
          ],
        ],
        shapes:
          '<p:pic><p:nvPicPr><p:cNvPr id="2" name="Red"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>' +
          '<p:blipFill><a:blip r:embed="rIdImg"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>' +
          `<p:spPr><a:xfrm><a:off x="${IN}" y="${IN}"/><a:ext cx="${2 * IN}" cy="${IN}"/></a:xfrm>` +
          '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>',
      },
    ],
    media: [
      {
        name: 'image1.png',
        bytes: new Uint8Array(red.toBuffer('image/png')),
        contentType: 'image/png',
      },
    ],
  });
  const runtime = await openWithImages(bytes);
  try {
    const png = await runtime.run(async (context) => {
      const image = context.presentation.slides.getItemAt(0).getImageAsBase64({ width: 480 });
      await context.sync();
      return image.value;
    });
    const picture = await loadImage(Buffer.from(png, 'base64'));
    expect([picture.width, picture.height]).toEqual([480, 270]);
    const canvas = createCanvas(480, 270);
    const ctx = canvas.getContext('2d');
    ctx.drawImage(picture, 0, 0);
    // 480 pixels across a 13.33-inch slide is 36 pixels per inch: the picture covers x 36-108
    // and y 36-72. Its center is red, and the slide background outside it is white.
    expect([...ctx.getImageData(72, 54, 1, 1).data]).toEqual([255, 0, 0, 255]);
    expect([...ctx.getImageData(200, 200, 1, 1).data]).toEqual([255, 255, 255, 255]);
  } finally {
    runtime.dispose();
  }
});

test('the duplicate example copies the slide after its source and tags the copy', async () => {
  const runtime = await PptxEditor.createServer(deck(), { substitutes: false });
  try {
    const answer = await runtime.run(async (context) => {
      const copy = context.presentation.slides.getItemAt(0).duplicate();
      copy.tags.add('VARIANT', 'B');
      copy.load('id,index');
      await context.sync();
      return { id: copy.id, index: copy.index };
    });
    expect(answer.index).toBe(1);
    await runtime.run(async (context) => {
      const slides = context.presentation.slides.load('items/id');
      await context.sync();
      expect(slides.items.map((s) => s.id)).toEqual(['256', answer.id, '257']);
      const tags = slides.items[1]!.tags.load('items/key,items/value');
      await context.sync();
      expect(tags.items.map((t) => [t.key, t.value])).toEqual([['VARIANT', 'B']]);
    });
  } finally {
    runtime.dispose();
  }
});

test('the text fit example keeps text that fits and shortens text that overflows', async () => {
  // The default substitutes: a runtime with no face at all refuses the read.
  const runtime = await PptxEditor.createServer(deck());
  try {
    // From text-and-ranges.mdx, "Check whether text fits its shape".
    const writeIfItFits = (shapeId: string, text: string, shorterText: string) =>
      runtime.run(async (context) => {
        const frame = context.presentation.slides.getItemAt(0).shapes.getItem(shapeId).textFrame;
        frame.textRange.text = text;
        const fit = frame.getTextFit();
        await context.sync();
        if (!fit.value.overflows) return true;
        frame.textRange.text = shorterText;
        await context.sync();
        return false;
      });
    // Shape 3 is one inch square, and a write gives it a text body with no autofit.
    expect(await writeIfItFits('3', 'Fits', 'Short')).toBe(true);
    expect(await writeIfItFits('3', 'Too long for the box. '.repeat(20), 'Short')).toBe(false);
    const text = await runtime.run(async (context) => {
      const range = context.presentation.slides.getItemAt(0).shapes.getItem('3')
        .textFrame.textRange;
      range.load('text');
      await context.sync();
      return range.text;
    });
    expect(text).toBe('Short');
  } finally {
    runtime.dispose();
  }
});

test('the text-and-ranges Replace All example replaces whole words and counts them', async () => {
  const runtime = await PptxEditor.createServer(deck(), { substitutes: false });
  try {
    await runtime.run(async (context) => {
      context.presentation.slides.getItemAt(0).shapes.addTextBox('FY25 plan, not FY250');
      await context.sync();
      const count = context.presentation.replaceAll('FY25', 'FY26', { matchWholeWord: true });
      await context.sync();
      expect(count.value).toBe(1);
      context.presentation.slides.getItemAt(0).shapes.addTextBox('For {{client}}');
      await context.sync();
      const slide = context.presentation.slides.getItemAt(0);
      const onSlide = slide.replaceAll('{{client}}', 'Northwind', { matchCase: true });
      await context.sync();
      expect(onSlide.value).toBe(1);
    });
  } finally {
    runtime.dispose();
  }
});

/** Two slides, each with one text box: shape 2 says 'One' on slide 1 and 'Two' on slide 2. */
function twoSlideDeck(): Uint8Array {
  const box = (text: string) =>
    shape({ id: 2, name: 'Title', x: IN, y: IN, cx: 4 * IN, cy: IN, text: para(run(text)) });
  return buildPptx({
    slides: [
      { layout: 'blank', shapes: box('One') },
      { layout: 'blank', shapes: box('Two') },
    ],
  });
}

async function slideIds(runtime: PptxEditorRuntime): Promise<string[]> {
  return runtime.run(async (context) => {
    const slides = context.presentation.slides.load('items/id');
    await context.sync();
    return slides.items.map((slide) => slide.id);
  });
}

async function texts(runtime: PptxEditorRuntime, ids: readonly string[]): Promise<string[]> {
  return runtime.run(async (context) => {
    const ranges = ids.map((id) =>
      context.presentation.slides.getItem(id).shapes.getItem('2').textFrame.textRange.load('text')
    );
    await context.sync();
    return ranges.map((range) => range.text);
  });
}

// The example of batching-and-errors.mdx, "Run parallel work", as written.

export interface TextTarget {
  readonly slideId: string;
  readonly shapeId: string;
}

// One queue for each runtime. A task starts when the task before it settles.
const queues = new WeakMap<PptxEditorRuntime, Promise<unknown>>();

export function enqueue<T>(runtime: PptxEditorRuntime, task: () => Promise<T>): Promise<T> {
  const result = (queues.get(runtime) ?? Promise.resolve()).then(task);
  const settled = result.catch(() => undefined);
  queues.set(runtime, settled);
  return result;
}

// Run a task again after StaleDocument, at most maxAttempts times in total.
async function retryStale<T>(task: () => Promise<T>, maxAttempts = 3): Promise<T> {
  for (let attempt = 1; ; attempt += 1) {
    try {
      return await task();
    } catch (error) {
      if (!isPptxEditorError(error) || error.code !== 'StaleDocument') throw error;
      // No revisions: earlier writes of that sync applied. Do not run it again.
      if (error.expectedRevision === undefined) throw error;
      if (attempt >= maxAttempts) throw error;
      console.warn('StaleDocument: reading again', {
        attempt,
        expectedRevision: error.expectedRevision,
        actualRevision: error.actualRevision,
      });
    }
  }
}

// A target whose slide or shape is gone reads as null after the sync.
function loadTexts(context: RequestContext, targets: readonly TextTarget[]) {
  return targets.map((target) => {
    const slide = context.presentation.slides.getItemOrNullObject(target.slideId);
    const shape = slide.shapes.getItemOrNullObject(target.shapeId);
    const range = shape.textFrame.textRange.load('text');
    return { target, text: () => (shape.isNullObject ? null : range.text), range };
  });
}

export async function rewriteTexts(
  runtime: PptxEditorRuntime,
  targets: readonly TextTarget[],
  rewrite: (text: string) => Promise<string>
) {
  // 1. Read every target in one short run.
  const before = await retryStale(() =>
    enqueue(runtime, () =>
      runtime.run(async (context) => {
        const items = loadTexts(context, targets);
        await context.sync();
        return items.map((item) => item.text());
      })
    )
  );
  // 2. Do the slow work outside run(), in parallel. A failed call gives null.
  const outcomes = await Promise.allSettled(
    before.map((text) => (text === null ? Promise.resolve(null) : rewrite(text)))
  );
  const after = outcomes.map((outcome) => (outcome.status === 'fulfilled' ? outcome.value : null));
  // 3. Apply every result in one run. Read again, and write only where the text is unchanged.
  return retryStale(() =>
    enqueue(runtime, () =>
      runtime.run(async (context) => {
        const items = loadTexts(context, targets);
        await context.sync();
        const skipped: TextTarget[] = [];
        items.forEach((item, index) => {
          const result = after[index];
          if (typeof result === 'string' && item.text() === before[index]) item.range.text = result;
          else skipped.push(item.target);
        });
        await context.sync();
        return { skipped };
      })
    )
  );
}

// End of the example.

/** Each interfering write uses a new name, so that each one commits. */
let interferences = 0;

/**
 * `runtime`, except that after the first sync of each of the `runs` listed (counted from 1), a
 * separate run renames shape 2 of `slideId`: another writer moves the presentation mid-run.
 */
function interfered(runtime: PptxEditorRuntime, slideId: string, runs: readonly number[]) {
  let count = 0;
  const run = (callback: (context: RequestContext) => Promise<unknown>) => {
    count += 1;
    const interfere = runs.includes(count);
    return runtime.run(async (context) => {
      const sync = context.sync.bind(context);
      let syncs = 0;
      context.sync = (async () => {
        await sync();
        syncs += 1;
        if (interfere && syncs === 1)
          await runtime.run(async (other) => {
            other.presentation.slides.getItem(slideId).shapes.getItem('2').name =
              `moved ${(interferences += 1)}`;
            await other.sync();
          });
      }) as RequestContext['sync'];
      return callback(context);
    });
  };
  return { ...runtime, run } as PptxEditorRuntime;
}

/** A stand-in for a model call: slow, and outside the presentation. */
const shout = (text: string) =>
  new Promise<string>((resolve) => setTimeout(() => resolve(text.toUpperCase()), 5));

describe('the parallel work rules', () => {
  const rename = (runtime: PptxEditorRuntime, name: string) =>
    runtime.run(async (context) => {
      context.presentation.slides.getItemAt(1).shapes.getItem('2').name = name;
      await context.sync();
    });

  test('the first sync of a run writes without a check', async () => {
    const runtime = await PptxEditor.createServer(twoSlideDeck(), { substitutes: false });
    try {
      let release!: () => void;
      const gate = new Promise<void>((resolve) => (release = resolve));
      const late = runtime.run(async (context) => {
        await gate;
        context.presentation.slides.getItemAt(0).shapes.getItem('2').name = 'Late';
        await context.sync();
      });
      await rename(runtime, 'First');
      release();
      await late;
      await runtime.run(async (context) => {
        const shape = context.presentation.slides.getItemAt(0).shapes.getItem('2').load('name');
        await context.sync();
        expect(shape.name).toBe('Late');
      });
    } finally {
      runtime.dispose();
    }
  });

  test('a first sync that reads in several host batches refuses before its writes', async () => {
    const runtime = await PptxEditor.createServer(twoSlideDeck(), { substitutes: false });
    try {
      const mine = runtime.run(async (context) => {
        const shapes = context.presentation.slides.getItemAt(0).shapes;
        shapes.load('items/name');
        shapes.getItem('2').name = 'Mine';
        await context.sync();
      });
      const write = rename(runtime, 'Between');
      const error = await rejection(mine);
      await write;
      if (!isPptxEditorError(error)) throw error;
      expect(error.code).toBe('StaleDocument');
      expect([error.expectedRevision, error.actualRevision]).toEqual([0, 1]);
    } finally {
      runtime.dispose();
    }
  });

  test('a read in several host batches refuses when another run writes between them', async () => {
    const runtime = await PptxEditor.createServer(twoSlideDeck(), { substitutes: false });
    try {
      // A listing and its item reads: two host batches, the second pinned to the first.
      const read = runtime.run(async (context) => {
        context.presentation.slides.getItemAt(0).shapes.load('items/name');
        await context.sync();
      });
      // A write from another run, sent after the listing and before the item reads.
      const write = rename(runtime, 'Between');
      const error = await rejection(read);
      await write;
      expect(isPptxEditorError(error) && error.code).toBe('StaleDocument');
    } finally {
      runtime.dispose();
    }
  });
});

describe('the parallel work example', () => {
  test('runs that hold a read across slow work conflict, even on different slides', async () => {
    const runtime = await PptxEditor.createServer(twoSlideDeck(), { substitutes: false });
    try {
      const ids = await slideIds(runtime);
      const edit = (slideId: string) =>
        runtime.run(async (context) => {
          const range = context.presentation.slides
            .getItem(slideId)
            .shapes.getItem('2')
            .textFrame.textRange.load('text');
          await context.sync();
          range.text = await shout(range.text); // The slow work inside run(): the mistake.
          await context.sync();
        });
      const outcomes = await Promise.allSettled(ids.map(edit));
      expect(outcomes[0]!.status).toBe('fulfilled');
      const failure = outcomes[1]!;
      expect(failure.status).toBe('rejected');
      const error = (failure as PromiseRejectedResult).reason as unknown;
      if (!isPptxEditorError(error)) throw error;
      expect(error.code).toBe('StaleDocument');
      expect(error.expectedRevision).toBe(0);
      expect(error.actualRevision).toBe(1);
      expect(await texts(runtime, ids)).toEqual(['ONE', 'Two']);
    } finally {
      runtime.dispose();
    }
  });

  test('jobs that fan out and apply through the queue all succeed, with no retry', async () => {
    const runtime = await PptxEditor.createServer(twoSlideDeck(), { substitutes: false });
    const warn = spyOn(console, 'warn').mockImplementation(() => {});
    try {
      const ids = await slideIds(runtime);
      // A slow call and an immediate one, so that the runs of the two jobs interleave.
      const instant = async (text: string) => text.toUpperCase();
      for (const rewrite of [shout, instant]) {
        const jobs = ids.map((slideId) =>
          rewriteTexts(runtime, [{ slideId, shapeId: '2' }], rewrite)
        );
        expect(await Promise.all(jobs)).toEqual([{ skipped: [] }, { skipped: [] }]);
      }
      expect(await texts(runtime, ids)).toEqual(['ONE', 'TWO']);
      expect(warn).not.toHaveBeenCalled();
    } finally {
      warn.mockRestore();
      runtime.dispose();
    }
  });

  test('without the queue, the same jobs conflict', async () => {
    const runtime = await PptxEditor.createServer(twoSlideDeck(), { substitutes: false });
    try {
      const ids = await slideIds(runtime);
      // The steps of rewriteTexts with no enqueue and no retry.
      const job = async (slideId: string) => {
        const targets = [{ slideId, shapeId: '2' }];
        const before = await runtime.run(async (context) => {
          const items = loadTexts(context, targets);
          await context.sync();
          return items.map((item) => item.text());
        });
        const after = before.map((text) => text?.toUpperCase());
        return runtime.run(async (context) => {
          const items = loadTexts(context, targets);
          await context.sync();
          items.forEach((item, index) => {
            if (item.text() === before[index]) item.range.text = after[index]!;
          });
          await context.sync();
        });
      };
      const outcomes = await Promise.allSettled(ids.map(job));
      const codes = outcomes.map((o) =>
        o.status === 'rejected' && isPptxEditorError(o.reason) ? o.reason.code : o.status
      );
      expect(codes).toEqual(['fulfilled', 'StaleDocument']);
    } finally {
      runtime.dispose();
    }
  });

  test('a target that changed during the slow work is skipped, not overwritten', async () => {
    const runtime = await PptxEditor.createServer(twoSlideDeck(), { substitutes: false });
    try {
      const ids = await slideIds(runtime);
      const targets = ids.map((slideId) => ({ slideId, shapeId: '2' }));
      const rewrite = async (text: string) => {
        if (text === 'One')
          await runtime.run(async (context) => {
            context.presentation.slides
              .getItem(ids[0]!)
              .shapes.getItem('2').textFrame.textRange.text = 'Edited by a person';
            await context.sync();
          });
        return shout(text);
      };
      expect(await rewriteTexts(runtime, targets, rewrite)).toEqual({ skipped: [targets[0]] });
      expect(await texts(runtime, ids)).toEqual(['Edited by a person', 'TWO']);
    } finally {
      runtime.dispose();
    }
  });

  test('a target deleted during the slow work is skipped, and the others apply', async () => {
    const runtime = await PptxEditor.createServer(twoSlideDeck(), { substitutes: false });
    try {
      const ids = await slideIds(runtime);
      const targets = [
        ...ids.map((slideId) => ({ slideId, shapeId: '2' })),
        { slideId: ids[0]!, shapeId: '99' },
      ];
      const rewrite = async (text: string) => {
        if (text === 'Two')
          await runtime.run(async (context) => {
            context.presentation.slides.getItem(ids[1]!).delete();
            await context.sync();
          });
        return shout(text);
      };
      expect(await rewriteTexts(runtime, targets, rewrite)).toEqual({
        skipped: [targets[1], targets[2]],
      });
      expect(await texts(runtime, [ids[0]!])).toEqual(['ONE']);
    } finally {
      runtime.dispose();
    }
  });

  test('a failed model call skips its target, and the others apply', async () => {
    const runtime = await PptxEditor.createServer(twoSlideDeck(), { substitutes: false });
    try {
      const ids = await slideIds(runtime);
      const targets = ids.map((slideId) => ({ slideId, shapeId: '2' }));
      const rewrite = async (text: string) => {
        if (text === 'One') throw new Error('the model call failed');
        return shout(text);
      };
      expect(await rewriteTexts(runtime, targets, rewrite)).toEqual({ skipped: [targets[0]] });
      expect(await texts(runtime, ids)).toEqual(['One', 'TWO']);
    } finally {
      runtime.dispose();
    }
  });

  test('a stale apply is logged with its revisions and retried, within the bound', async () => {
    const runtime = await PptxEditor.createServer(twoSlideDeck(), { substitutes: false });
    const warn = spyOn(console, 'warn').mockImplementation(() => {});
    try {
      const ids = await slideIds(runtime);
      const targets = [{ slideId: ids[0]!, shapeId: '2' }];
      // Run 1 reads; run 2, the first apply, is interfered with; run 3 applies.
      const once = interfered(runtime, ids[1]!, [2]);
      expect(await rewriteTexts(once, targets, shout)).toEqual({ skipped: [] });
      expect(await texts(runtime, ids)).toEqual(['ONE', 'Two']);
      expect(warn).toHaveBeenCalledTimes(1);
      expect(warn.mock.calls[0]).toEqual([
        'StaleDocument: reading again',
        { attempt: 1, expectedRevision: 0, actualRevision: 1 },
      ]);

      // Every apply is interfered with: the third refusal is thrown, and nothing applies.
      warn.mockClear();
      const always = interfered(runtime, ids[1]!, [2, 3, 4]);
      const error = await rejection(
        rewriteTexts(always, [{ slideId: ids[1]!, shapeId: '2' }], shout)
      );
      expect(isPptxEditorError(error) && error.code).toBe('StaleDocument');
      expect(warn).toHaveBeenCalledTimes(2);
      expect(await texts(runtime, ids)).toEqual(['ONE', 'Two']);
    } finally {
      warn.mockRestore();
      runtime.dispose();
    }
  });
});

test('the section examples group, list, move, and rename sections as written', async () => {
  // docs/site/content/editor-api/sections.mdx.
  async function listSections(runtime: PptxEditorRuntime) {
    return runtime.run(async (context) => {
      const sections = context.presentation.sections.load('items/name,items/slideIds');
      await context.sync();
      return sections.items.map((section) => ({ name: section.name, slideIds: section.slideIds }));
    });
  }
  async function groupSlides(
    runtime: PptxEditorRuntime,
    groups: { name: string; firstSlideId: string }[]
  ) {
    return runtime.run(async (context) => {
      const sections = context.presentation.sections;
      const added = groups.map((group) => sections.add(group.name, group.firstSlideId));
      for (const section of added) section.load('id');
      await context.sync();
      return added.map((section) => section.id);
    });
  }
  async function moveToEnd(runtime: PptxEditorRuntime, sectionId: string, name: string) {
    await runtime.run(async (context) => {
      const sections = context.presentation.sections;
      const count = sections.getCount();
      await context.sync();
      const section = sections.getItem(sectionId);
      section.moveTo(count.value - 1);
      section.name = name;
      await context.sync();
    });
  }
  const bytes = buildPptx({
    slides: [1, 2, 3, 4].map(() => ({ layout: 'blank' as const, shapes: '' })),
  });
  const runtime = await PptxEditor.createServer(bytes, { substitutes: false });
  try {
    expect(await listSections(runtime)).toEqual([]);
    const [intro] = await groupSlides(runtime, [
      { name: 'Intro', firstSlideId: '256' },
      { name: 'Results', firstSlideId: '258' },
    ]);
    expect(await listSections(runtime)).toEqual([
      { name: 'Intro', slideIds: ['256', '257'] },
      { name: 'Results', slideIds: ['258', '259'] },
    ]);
    await moveToEnd(runtime, intro!, 'Appendix');
    expect(await listSections(runtime)).toEqual([
      { name: 'Results', slideIds: ['258', '259'] },
      { name: 'Appendix', slideIds: ['256', '257'] },
    ]);
  } finally {
    runtime.dispose();
  }
});
