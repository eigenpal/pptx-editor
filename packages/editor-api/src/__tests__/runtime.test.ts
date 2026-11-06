/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/editor-api/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// The runtime rules: batching, load, null objects, errors, stale reads, lifetimes.

import { describe, expect, test } from 'bun:test';
import { isPptxEditorError, PptxEditor, type RequestContext, type Shape } from '../index.ts';
import { buildPptx, para, run, shape } from '../../../../scripts/fixtures/builder.ts';
import type { AutomationHost } from '@pptx-editor.dev/core/automation';
import { createRuntime } from '../runtime/runtime.ts';
import { deck, IN, open, rejection, reopen, thrown } from './support.ts';

const code = (error: unknown): string | undefined =>
  isPptxEditorError(error) ? error.code : String(error);

describe('opening', () => {
  test('refuses bytes that are not a presentation', async () => {
    expect(code(await rejection(PptxEditor.createServer(new Uint8Array([1, 2, 3]))))).toBe(
      'InvalidArgument'
    );
    expect(code(await rejection(PptxEditor.createServer(new Uint8Array())))).toBe(
      'InvalidArgument'
    );
  });

  test('a server runtime saves and has no selection', async () => {
    const runtime = await open();
    expect(runtime.capabilities).toMatchObject({ save: true, selection: false, autofit: true });
    const error = await rejection(
      runtime.run(async (context) => {
        context.presentation.getSelectedSlides().load('items');
        await context.sync();
      })
    );
    expect(code(error)).toBe('NotSupported');
    expect(isPptxEditorError(error) && error.debugInfo.innerError?.code).toBe('host.no-selection');
    runtime.dispose();
  });
});

describe('load and sync', () => {
  test('a property is readable only after its load synced', async () => {
    const runtime = await open();
    await runtime.run(async (context) => {
      const slide = context.presentation.slides.getItemAt(0);
      expect(code(thrown(() => slide.id))).toBe('PropertyNotLoaded');
      slide.load('id');
      expect(code(thrown(() => slide.id))).toBe('PropertyNotLoaded');
      await context.sync();
      expect(slide.id).toMatch(/^\d+$/);
      expect(code(thrown(() => slide.index))).toBe('PropertyNotLoaded');
    });
  });

  test('collections load items and item properties in one sync', async () => {
    const runtime = await open();
    const names = await runtime.run(async (context) => {
      const shapes = context.presentation.slides.getItemAt(0).shapes;
      shapes.load('items/name,items/left');
      const count = shapes.getCount();
      expect(code(thrown(() => count.value))).toBe('ValueNotLoaded');
      await context.sync();
      expect(count.value).toBe(2);
      expect(shapes.items[0]!.left).toBe(72);
      return shapes.items.map((s) => s.name);
    });
    expect(names).toEqual(['Box', 'Other']);
  });

  test('top and skip window the items', async () => {
    const runtime = await open();
    await runtime.run(async (context) => {
      const shapes = context.presentation.slides.getItemAt(0).shapes;
      shapes.load({ select: 'name', skip: 1, top: 5 });
      await context.sync();
      expect(shapes.items.map((s) => s.name)).toEqual(['Other']);
    });
  });

  test('load options are checked', async () => {
    const runtime = await open();
    await runtime.run(async (context) => {
      const shape = context.presentation.slides.getItemAt(0).shapes.getItemAt(0);
      expect(code(thrown(() => shape.load('nope')))).toBe('InvalidArgument');
      expect(code(thrown(() => shape.load('items/name')))).toBe('InvalidArgument');
      expect(code(thrown(() => shape.load({ select: 'name', bogus: 1 } as never)))).toBe(
        'InvalidArgument'
      );
      shape.load('fill');
      await context.sync();
      expect(shape.fill.type).toBe('Solid');
    });
  });

  test('toJSON gives the loaded properties', async () => {
    const runtime = await open();
    await runtime.run(async (context) => {
      const shape = context.presentation.slides.getItemAt(0).shapes.getItem('2');
      shape.load('name,width');
      await context.sync();
      expect(shape.toJSON()).toEqual({ name: 'Box', width: 216 });
    });
  });

  test('toJSON gives plain data for items and loaded navigation objects', async () => {
    const runtime = await open();
    await runtime.run(async (context) => {
      const shapes = context.presentation.slides.getItemAt(0).shapes;
      shapes.load('items/name');
      const shape = shapes.getItem('2');
      shape.load('name,fill');
      await context.sync();
      const listed = shapes.toJSON() as { items: unknown[] };
      expect(listed).toEqual({ items: [{ name: 'Box' }, { name: 'Other' }] });
      expect(Object.getPrototypeOf(listed.items[0])).toBe(Object.prototype);
      expect(shape.toJSON()).toEqual({
        name: 'Box',
        fill: { type: 'Solid', foregroundColor: '#4472C4', transparency: 0 },
      });
    });
  });

  test('an empty sync is not a round trip', async () => {
    const runtime = await open();
    await runtime.run(async (context) => {
      await context.sync();
      await context.sync();
    });
  });
});

describe('null objects and missing items', () => {
  test('getItemOrNullObject answers a null object, getItem refuses', async () => {
    const runtime = await open();
    await runtime.run(async (context) => {
      const missing = context.presentation.slides.getItemOrNullObject('999');
      expect(code(thrown(() => missing.isNullObject))).toBe('PropertyNotLoaded');
      missing.load('id');
      await context.sync();
      expect(missing.isNullObject).toBe(true);
      const found = context.presentation.slides.getItemOrNullObject(
        (await idOfFirstSlide(context)) as string
      );
      found.load('index');
      await context.sync();
      expect(found.isNullObject).toBe(false);
      expect(found.index).toBe(0);
    });
    const error = await rejection(
      runtime.run(async (context) => {
        context.presentation.slides.getItem('999').load('id');
        await context.sync();
      })
    );
    expect(code(error)).toBe('ItemNotFound');
    expect(isPptxEditorError(error) && error.target).toBe('presentation.slides.getItem()');
  });

  test('a load on a known null object passes and loads nothing', async () => {
    const runtime = await open();
    await runtime.run(async (context) => {
      const missing = context.presentation.slides.getItemOrNullObject('999');
      missing.load('id');
      await context.sync();
      expect(missing.load('id')).toBe(missing);
      await context.sync();
      expect(missing.isNullObject).toBe(true);
      expect(missing.toJSON()).toEqual({});
      expect(code(thrown(() => missing.delete()))).toBe('InvalidObjectPath');
    });
  });

  test('a write to a null object says that it is a null object', async () => {
    const runtime = await open();
    // The lookup and the write go in one sync: the write waits for the lookup, then refuses.
    const error = await rejection(
      runtime.run(async (context) => {
        const missing = context.presentation.slides.getItemAt(0).shapes.getItemOrNullObject('99');
        missing.name = 'Never';
        await context.sync();
      })
    );
    expect(code(error)).toBe('InvalidObjectPath');
    expect(isPptxEditorError(error) && error.debugInfo.innerError?.code).toBe('edit.null-object');
    expect(String((error as Error).message)).toContain('null object');
    // A known null object refuses at the call, with the same reason.
    await runtime.run(async (context) => {
      const missing = context.presentation.slides.getItemOrNullObject('999').load('id');
      await context.sync();
      const known = thrown(() => missing.delete());
      expect(isPptxEditorError(known) && known.debugInfo.innerError?.code).toBe('edit.null-object');
    });
  });
});

async function idOfFirstSlide(context: {
  presentation: { slides: { getItemAt(i: number): { load(o: string): { id: string } } } };
  sync(): Promise<void>;
}): Promise<string> {
  const slide = context.presentation.slides.getItemAt(0).load('id');
  await context.sync();
  return slide.id;
}

describe('atomic batches', () => {
  test('a refused call commits nothing of its batch', async () => {
    const runtime = await open();
    const error = await rejection(
      runtime.run(async (context) => {
        const box = context.presentation.slides.getItemAt(0).shapes.getItem('2');
        box.textFrame.textRange.text = 'Changed';
        box.name = 'Renamed';
        context.presentation.slides.getItem('999').delete();
        await context.sync();
      })
    );
    expect(code(error)).toBe('ItemNotFound');
    await runtime.run(async (context) => {
      const box = context.presentation.slides.getItemAt(0).shapes.getItem('2');
      const range = box.textFrame.textRange.load('text');
      box.load('name');
      await context.sync();
      expect(range.text).toBe('Hello\rWorld');
      expect(box.name).toBe('Box');
    });
  });

  test('a bad argument refuses the whole sync, so the calls before it do not apply', async () => {
    const runtime = await open();
    await runtime.run(async (context) => {
      const box = context.presentation.slides.getItemAt(0).shapes.getItem('2');
      box.name = 'Renamed';
      const fill = context.presentation.slides.getItemAt(0).background.fill;
      // The call does not throw: the refusal waits for the sync, as in Office.js.
      fill.getGradientFillOrNullObject().type = 'ShadeFromTitle';
      box.textFrame.textRange.text = 'Changed';
      const error = await rejection(context.sync());
      expect(code(error)).toBe('NotImplemented');
      expect(isPptxEditorError(error) && error.debugInfo.errorLocation).toBe(
        'presentation.slides.getItemAt(0).background.fill.getGradientFillOrNullObject().type'
      );
      const range = box.textFrame.textRange.load('text');
      box.load('name');
      await context.sync();
      expect([box.name, range.text]).toEqual(['Box', 'Hello\rWorld']);
    });
  });

  test('a refused call that answers an object answers one that takes further calls', async () => {
    const runtime = await open();
    await runtime.run(async (context) => {
      const shapes = context.presentation.slides.getItemAt(0).shapes;
      const bad = shapes.addGeometricShape('Blob' as never);
      bad.name = 'Never';
      bad.fill.setSolidColor('#FF0000');
      bad.textFrame.textRange.font.bold = true;
      bad.load('name');
      expect(code(await rejection(context.sync()))).toBe('InvalidArgument');
      shapes.load('items/name');
      await context.sync();
      expect(shapes.items.map((s) => s.name)).toEqual(['Box', 'Other']);
    });
  });

  test('what still throws at the call: a read before its load, and load options', async () => {
    const runtime = await open();
    await runtime.run(async (context) => {
      const box = context.presentation.slides.getItemAt(0).shapes.getItem('2');
      box.name = 'Renamed';
      expect(code(thrown(() => box.left))).toBe('PropertyNotLoaded');
      expect(code(thrown(() => box.load('nope/items')))).toBe('InvalidArgument');
      // A call that throws queues nothing, and the calls before it stay queued.
      box.load('name');
      await context.sync();
      expect(box.name).toBe('Renamed');
    });
  });

  test('proxies from a refused sync stay usable: a retry with the same objects works', async () => {
    const runtime = await open();
    await runtime.run(async (context) => {
      const slide = context.presentation.slides.getItemAt(0);
      const box = slide.shapes.getItem('2');
      const added = slide.shapes.addTextBox('never added');
      const missing = context.presentation.slides.getItem('999');
      box.name = 'Renamed';
      missing.delete();
      expect(code(await rejection(context.sync()))).toBe('ItemNotFound');
      // Only that sync's calls were dropped: the box was not renamed.
      box.load('name');
      await context.sync();
      expect(box.name).toBe('Box');
      // The retry with the same proxy applies.
      box.name = 'Renamed';
      box.load('name');
      await context.sync();
      expect(box.name).toBe('Renamed');
      // An object that the refused sync would have created does not exist.
      added.load('name');
      const gone = await rejection(context.sync());
      expect(code(gone)).toBe('InvalidObjectPath');
      expect(isPptxEditorError(gone) && gone.debugInfo.innerError).toEqual({
        code: 'sync.object-from-refused-sync',
        message:
          'This object came from a sync that was refused, so it was never created. Create it again.',
      });
      // A write to it says the same.
      added.name = 'still?';
      const write = await rejection(context.sync());
      expect(isPptxEditorError(write) && write.debugInfo.innerError?.code).toBe(
        'sync.object-from-refused-sync'
      );
      // A lookup that found nothing still finds nothing.
      missing.load('id');
      expect(code(await rejection(context.sync()))).toBe('ItemNotFound');
      const shapes = slide.shapes.load('items/name');
      await context.sync();
      expect(shapes.items.map((s) => s.name)).toEqual(['Renamed', 'Other']);
    });
  });

  test('a lookup chain from a refused sync is looked up again, deepest first', async () => {
    const runtime = await open();
    await runtime.run(async (context) => {
      const range = context.presentation.slides.getItem('256').shapes.getItem('2')
        .textFrame.textRange;
      range.text = 'First try';
      context.presentation.slides.getItem('999').delete();
      expect(code(await rejection(context.sync()))).toBe('ItemNotFound');
      range.text = 'Second try';
      range.load('text');
      await context.sync();
      expect(range.text).toBe('Second try');
    });
  });

  test('reads and lookups see the writes queued before them in the same sync', async () => {
    const runtime = await open();
    await runtime.run(async (context) => {
      const p = context.presentation;
      // Tags: a lookup that a write uses, after the add.
      p.tags.add('STATUS', 'draft');
      p.tags.getItemOrNullObject('STATUS').value = 'final';
      const tag = p.tags.getItem('STATUS').load('value');
      // Custom XML: the part a write adds, found by namespace and changed.
      p.customXmlParts.add('<a xmlns="urn:order"/>');
      p.customXmlParts
        .getByNamespace('urn:order')
        .getOnlyItemOrNullObject()
        .setXml('<b xmlns="urn:order"/>');
      const xml = p.customXmlParts.getByNamespace('urn:order').getOnlyItem().getXml();
      // Slides: the slide a write adds, and a deleted slide that is gone.
      p.slides.add();
      const added = p.slides.getItemAt(2).load('id');
      p.slides.getItemAt(1).delete();
      const deleted = p.slides.getItemOrNullObject('257');
      // Shapes: a table cell of a table the same sync adds.
      const table = p.slides.getItemAt(0).shapes.addTable(2, 2).getTable();
      table.getCellOrNullObject(1, 1).text = 'cell';
      table.load('values');
      await context.sync();
      expect(tag.value).toBe('final');
      expect(xml.value).toBe('<b xmlns="urn:order"/>');
      expect(added.id).toBe('258');
      expect(deleted.isNullObject).toBe(true);
      expect(table.values).toEqual([
        ['', ''],
        ['', 'cell'],
      ]);
    });
  });

  test('a refusal after a split keeps its cause and says the writes before it applied', async () => {
    const runtime = await open();
    const error = await rejection(
      runtime.run(async (context) => {
        const tags = context.presentation.tags;
        tags.add('A', '1');
        tags.getItemOrNullObject('A').value = '2';
        context.presentation.slides.getItem('999').delete();
        await context.sync();
      })
    );
    expect(code(error)).toBe('ItemNotFound');
    expect(isPptxEditorError(error) && error.debugInfo.innerError?.message).toContain(
      'those writes applied first'
    );
    await runtime.run(async (context) => {
      const tag = context.presentation.tags.getItem('A').load('value');
      await context.sync();
      expect(tag.value).toBe('1');
    });
  });

  test('an object a write creates is usable in the same sync', async () => {
    const runtime = await open();
    await runtime.run(async (context) => {
      const shapes = context.presentation.slides.getItemAt(1).shapes;
      const box = shapes.addTextBox('Created', { left: 10, top: 20, width: 300, height: 40 });
      box.name = 'Fresh';
      box.textFrame.textRange.font.italic = true;
      box.load('name,left,top,type');
      const font = box.textFrame.textRange.font.load('italic');
      await context.sync();
      expect(box.name).toBe('Fresh');
      expect([box.left, box.top]).toEqual([10, 20]);
      expect(box.type).toBe('TextBox');
      expect(font.italic).toBe(true);
    });
    const again = await reopen(runtime);
    await again.run(async (context) => {
      const shapes = context.presentation.slides.getItemAt(1).shapes.load('items/name');
      await context.sync();
      expect(shapes.items.map((s) => s.name)).toContain('Fresh');
    });
  });

  test('a write to a shape whose id the file repeats refuses with a diagnostic', async () => {
    const twin = (name: string): string =>
      shape({ id: 2, name, x: 0, y: 0, cx: IN, cy: IN, text: para(run(name)) });
    const runtime = await open(
      buildPptx({ slides: [{ layout: 'blank', shapes: twin('A') + twin('B') }] })
    );
    const error = await rejection(
      runtime.run(async (context) => {
        const shapes = context.presentation.slides.getItemAt(0).shapes.load('items');
        await context.sync();
        shapes.items[1]!.name = 'Renamed';
        await context.sync();
      })
    );
    expect(code(error)).toBe('NotSupported');
    expect(isPptxEditorError(error) && error.debugInfo.innerError?.code).toBe(
      'edit.duplicate-shape-id'
    );
  });

  test('item reads of a listing go before a write that removes an item', async () => {
    const runtime = await open();
    const ids = await runtime.run(async (context) => {
      const slides = context.presentation.slides;
      slides.load('items/id');
      slides.getItemAt(1).delete();
      await context.sync();
      return slides.items.map((slide) => slide.id);
    });
    expect(ids).toHaveLength(2);
    const count = await runtime.run(async (context) => {
      const result = context.presentation.slides.getCount();
      await context.sync();
      return result.value;
    });
    expect(count).toBe(1);
  });

  test('an item read that fails after the writes committed says that they applied', async () => {
    const runtime = await open();
    const error = await rejection(
      runtime.run(async (context) => {
        const slides = context.presentation.slides;
        slides.add();
        slides.load('items/id');
        slides.getItemAt(0).delete();
        await context.sync();
      })
    );
    expect(code(error)).toBe('InvalidObjectPath');
    expect(isPptxEditorError(error) && error.debugInfo.innerError?.code).toBe(
      'edit.read-after-commit'
    );
  });

  test('a queued write without a sync is dropped when the run ends', async () => {
    const runtime = await open();
    await runtime.run(async (context) => {
      context.presentation.slides.getItemAt(0).shapes.getItem('2').name = 'Never';
    });
    await runtime.run(async (context) => {
      const box = context.presentation.slides.getItemAt(0).shapes.getItem('2').load('name');
      await context.sync();
      expect(box.name).toBe('Box');
    });
  });
});

describe('observers', () => {
  test('onRefusal sees refusals and onException sees engine errors, then both propagate', async () => {
    const seen: string[] = [];
    const defect = new Error('engine defect');
    let mode: 'refuse' | 'throw' = 'refuse';
    const host: AutomationHost = {
      capabilities: {
        document: true,
        save: false,
        events: false,
        selection: false,
        autofit: false,
        images: false,
      },
      revision: () => 1,
      execute: (request) => {
        if (mode === 'throw') throw defect;
        return {
          ok: false,
          results: request.operations.map(() => ({
            status: 'error' as const,
            error: { code: 'transaction-refused' as const, message: 'no' },
          })),
          revision: 1,
          changed: false,
        };
      },
      save: () => ({ ok: false, error: { code: 'unsupported-capability', message: 'no' } }),
      subscribe: () => () => {},
      dispose: () => {},
    };
    const runtime = createRuntime({
      host,
      save: false,
      onRefusal: (error) => seen.push(`refusal:${error.code}`),
      onException: (error) => {
        seen.push(`exception:${(error as Error).message}`);
        throw new Error('an observer that throws is ignored');
      },
    });
    const refuse = () =>
      runtime.run(async (context) => {
        context.presentation.slides.getCount();
        await context.sync();
      });
    expect(code(await rejection(refuse()))).toBe('GeneralException');
    mode = 'throw';
    expect(await rejection(refuse())).toBe(defect);
    expect(seen).toEqual(['refusal:GeneralException', 'exception:engine defect']);
  });
});

describe('stale reads', () => {
  test('a write decided from an old read refuses with StaleDocument', async () => {
    const runtime = await open();
    let release!: () => void;
    const gate = new Promise<void>((resolve) => (release = resolve));
    const slow = runtime.run(async (context) => {
      const box = context.presentation.slides.getItemAt(0).shapes.getItem('2').load('name');
      await context.sync();
      await gate;
      box.name = `${box.name} (checked)`;
      await context.sync();
    });
    await runtime.run(async (context) => {
      context.presentation.slides.getItemAt(0).shapes.getItem('3').name = 'Moved first';
      await context.sync();
    });
    release();
    const error = await rejection(slow);
    expect(code(error)).toBe('StaleDocument');
    expect(isPptxEditorError(error) && error.expectedRevision).toBeLessThan(
      (isPptxEditorError(error) && error.actualRevision) || 0
    );
  });
});

describe('lifetimes', () => {
  test('objects die with their run unless tracked', async () => {
    const runtime = await open();
    let loose!: Shape;
    let kept!: Shape;
    await runtime.run(async (context) => {
      loose = context.presentation.slides.getItemAt(0).shapes.getItem('2').load('name');
      kept = context.presentation.slides.getItemAt(0).shapes.getItem('3');
      context.trackedObjects.add(kept);
      await context.sync();
    });
    expect(loose.name).toBe('Box');
    expect(code(thrown(() => loose.load('left')))).toBe('InvalidObjectPath');
    await runtime.run(kept, async (context) => {
      kept.name = 'Adopted';
      await context.sync();
      kept.load('name');
      await context.sync();
      expect(kept.name).toBe('Adopted');
      context.trackedObjects.remove(kept);
    });
    expect(code(await rejection(runtime.run(loose, async () => {})))).toBe('InvalidObjectPath');
  });

  test('a finished context and a disposed runtime refuse', async () => {
    const runtime = await open();
    let saved!: RequestContext;
    await runtime.run(async (context) => {
      saved = context;
    });
    expect(code(thrown(() => saved.presentation))).toBe('InvalidRequestContext');
    runtime.dispose();
    runtime.dispose();
    expect(code(await rejection(runtime.run(async () => {})))).toBe('RuntimeDisposed');
    expect(code(await rejection(runtime.save()))).toBe('RuntimeDisposed');
  });

  test('run refuses a list that holds something other than a proxy', async () => {
    const runtime = await open();
    const error = await rejection(
      runtime.run([{} as unknown as Shape], async () => {
        /* never called */
      })
    );
    expect(code(error)).toBe('InvalidArgument');
  });

  test('adoption is all or nothing: a refused list moves no object', async () => {
    const runtime = await open();
    let kept!: Shape;
    let loose!: Shape;
    let first!: RequestContext;
    await runtime.run(async (context) => {
      first = context;
      kept = context.presentation.slides.getItemAt(0).shapes.getItem('2');
      loose = context.presentation.slides.getItemAt(0).shapes.getItem('3');
      context.trackedObjects.add(kept);
      await context.sync();
    });
    expect(code(await rejection(runtime.run([kept, loose], async () => {})))).toBe(
      'InvalidObjectPath'
    );
    // `kept` stayed with its first run, so a later run can still adopt it.
    expect(kept.context).toBe(first);
    const name = await runtime.run(kept, async (context) => {
      kept.load('name');
      await context.sync();
      return kept.name;
    });
    expect(name).toBe('Box');
  });

  test('two runtimes keep their objects apart', async () => {
    const first = await open();
    const second = await open(deck());
    let shape!: Shape;
    await first.run(async (context) => {
      shape = context.presentation.slides.getItemAt(0).shapes.getItem('2');
      context.trackedObjects.add(shape);
      await context.sync();
    });
    expect(code(await rejection(second.run(shape, async () => {})))).toBe('InvalidObjectPath');
  });
});
