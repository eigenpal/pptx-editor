/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/editor-api/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Multi-level load paths: `a/b/c` strings, the property tree, `expand`, and `$all`, with the
// number of host batches each one costs.

import { describe, expect, test } from 'bun:test';
import { createStoreAutomationHost, type AutomationHost } from '@pptx-editor.dev/core/automation';
import { PptxDocumentStore, readPptx, readZip } from '@pptx-editor.dev/core/store';
import { isPptxEditorError, type RequestContext } from '../index.ts';
import { createRuntime } from '../runtime/runtime.ts';
import { readdirSync, readFileSync } from 'node:fs';
import { join } from 'node:path';
import { buildPptx, cell, para, run, shape, table } from '../../../../scripts/fixtures/builder.ts';
import { deck, IN, open, rejection, thrown } from './support.ts';

/** One slide with three text boxes. */
function texts(): Uint8Array {
  const box = (id: number, words: string) =>
    shape({ id, name: `Box ${id}`, x: IN, y: id * IN, cx: IN, cy: IN, text: para(run(words)) });
  return buildPptx({
    slides: [{ layout: 'blank', shapes: box(2, 'A') + box(3, 'B') + box(4, 'C') }],
  });
}

const code = (error: unknown): string | undefined =>
  isPptxEditorError(error) ? error.code : String(error);
const target = (error: unknown): string | undefined =>
  isPptxEditorError(error) ? error.target : undefined;

/** A server runtime whose host records the operations of each batch it executes. */
function counting(bytes: Uint8Array = deck()) {
  const read = readPptx(bytes);
  if (!read.ok) throw new Error('fixture');
  const inner = createStoreAutomationHost(new PptxDocumentStore(read.pkg));
  const batches: string[][] = [];
  const host: AutomationHost = {
    ...inner,
    execute: (request) => {
      batches.push(request.operations.map((operation) => operation.op));
      return inner.execute(request);
    },
  };
  return { runtime: createRuntime({ host, save: true }), batches };
}

describe('path strings', () => {
  test('a path through a collection and two navigation objects reads in two batches', async () => {
    const { runtime, batches } = counting(texts());
    const read = await runtime.run(async (context) => {
      const slide = context.presentation.slides.getItemAt(0);
      slide.load('shapes/items/textFrame/textRange/text');
      await context.sync();
      return slide.shapes.items.map((item) => item.textFrame.textRange.text);
    });
    expect(read).toEqual(['A', 'B', 'C']);
    // The listing, then one batch with a text read for each shape: no round trip per shape.
    expect(batches).toHaveLength(2);
    expect(batches[1]).toHaveLength(3);
  });

  test('a collection under the items adds one round, not one per item', async () => {
    const { runtime, batches } = counting();
    const names = await runtime.run(async (context) => {
      const slides = context.presentation.slides;
      slides.load('items/shapes/items/name');
      await context.sync();
      return slides.items.map((slide) => slide.shapes.items.map((shape) => shape.name));
    });
    expect(names).toEqual([['Box', 'Other'], []]);
    expect(batches).toHaveLength(3);
  });

  test('a collection path names its own and its items properties together', async () => {
    const { runtime } = counting();
    await runtime.run(async (context) => {
      const masters = context.presentation.slideMasters;
      masters.load('id, name, layouts/items/name, layouts/items/id');
      await context.sync();
      const layouts = masters.items[0]!.layouts.items;
      expect(masters.items[0]!.name).toEqual(expect.any(String));
      expect(layouts.length).toBeGreaterThan(0);
      expect(layouts.every((layout) => typeof layout.name === 'string')).toBe(true);
      expect(layouts.every((layout) => typeof layout.id === 'string')).toBe(true);
    });
  });

  test('a path on an object loads only the property it ends on', async () => {
    const { runtime } = counting();
    await runtime.run(async (context) => {
      const shape = context.presentation.slides.getItemAt(0).shapes.getItem('2');
      shape.load('name, fill/foregroundColor, textFrame/textRange/font/bold');
      await context.sync();
      expect(shape.toJSON()).toEqual({
        name: 'Box',
        fill: { foregroundColor: '#4472C4' },
        textFrame: { textRange: { font: { bold: null } } },
      });
      expect(code(thrown(() => shape.fill.type))).toBe('PropertyNotLoaded');
      expect(code(thrown(() => shape.left))).toBe('PropertyNotLoaded');
    });
  });

  test('a navigation property named alone and with a path loads its default set', async () => {
    const { runtime } = counting();
    await runtime.run(async (context) => {
      const shape = context.presentation.slides.getItemAt(0).shapes.getItem('2');
      shape.load('fill, fill/type');
      await context.sync();
      expect(shape.fill.toJSON()).toEqual({
        type: 'Solid',
        foregroundColor: '#4472C4',
        transparency: 0,
      });
    });
  });

  test('a path without the items prefix means each item on a collection', async () => {
    const { runtime } = counting();
    await runtime.run(async (context) => {
      const shapes = context.presentation.slides.getItemAt(0).shapes;
      shapes.load('fill/type');
      await context.sync();
      expect(shapes.items.map((shape) => shape.fill.type)).toEqual(['Solid', 'NoFill']);
    });
  });
});

describe('a bare collection load', () => {
  test('load() and load("items") load every scalar property of each item in two batches', async () => {
    for (const option of [undefined, 'items'] as const) {
      const { runtime, batches } = counting();
      await runtime.run(async (context) => {
        const shapes = context.presentation.slides.getItemAt(0).shapes;
        shapes.load(option);
        await context.sync();
        expect(shapes.items.map((shape) => [shape.name, shape.left, shape.width])).toEqual([
          ['Box', 72, 216],
          ['Other', 360, 72],
        ]);
        expect(shapes.items[0]!.toJSON()).not.toHaveProperty('fill');
      });
      // The listing, then one batch with a read for each shape.
      expect(batches).toHaveLength(2);
      expect(batches[1]).toHaveLength(2);
    }
  });

  test('items with a path loads every scalar property and the path', async () => {
    const { runtime } = counting();
    await runtime.run(async (context) => {
      const shapes = context.presentation.slides.getItemAt(0).shapes;
      shapes.load('items, items/fill/type');
      await context.sync();
      expect(shapes.items.map((shape) => [shape.name, shape.fill.type])).toEqual([
        ['Box', 'Solid'],
        ['Other', 'NoFill'],
      ]);
    });
  });
});

describe('$all', () => {
  test('loads every scalar property and no navigation property', async () => {
    const { runtime } = counting();
    await runtime.run(async (context) => {
      const shape = context.presentation.slides.getItemAt(0).shapes.getItem('2');
      shape.load('$all');
      await context.sync();
      expect(shape.name).toBe('Box');
      expect(shape.left).toBe(72);
      expect(shape.toJSON()).not.toHaveProperty('fill');
    });
  });

  test('items/$all loads every scalar property of each item, with more paths', async () => {
    const { runtime } = counting();
    await runtime.run(async (context) => {
      const shapes = context.presentation.slides.getItemAt(0).shapes;
      shapes.load('items/$all, items/fill/type');
      await context.sync();
      expect(shapes.items.map((shape) => [shape.name, shape.width, shape.fill.type])).toEqual([
        ['Box', 216, 'Solid'],
        ['Other', 72, 'NoFill'],
      ]);
    });
  });

  test('$all after a path segment loads that object', async () => {
    const { runtime } = counting();
    await runtime.run(async (context) => {
      const shape = context.presentation.slides.getItemAt(0).shapes.getItem('2');
      shape.load('fill/$all');
      await context.sync();
      expect(shape.fill.type).toBe('Solid');
      expect(code(thrown(() => shape.name))).toBe('PropertyNotLoaded');
    });
  });
});

describe('the property tree', () => {
  test('loads a collection through an items key or straight item keys', async () => {
    const { runtime, batches } = counting();
    await runtime.run(async (context) => {
      const slide = context.presentation.slides.getItemAt(0);
      slide.load({ shapes: { items: { name: true, fill: { type: true } } } });
      await context.sync();
      expect(slide.shapes.items.map((shape) => [shape.name, shape.fill.type])).toEqual([
        ['Box', 'Solid'],
        ['Other', 'NoFill'],
      ]);
      const other = context.presentation.slides.getItemAt(0);
      other.load({ id: true, shapes: { name: true, $skip: 1, $top: 1 } });
      await context.sync();
      expect(other.id).toMatch(/^\d+$/);
      expect(other.shapes.items.map((shape) => shape.name)).toEqual(['Other']);
    });
    expect(batches).toHaveLength(4);
  });

  test('an empty object loads the default set, and false loads nothing', async () => {
    const { runtime } = counting();
    await runtime.run(async (context) => {
      const shape = context.presentation.slides.getItemAt(0).shapes.getItem('2');
      shape.load({ name: true, left: false, fill: {} });
      await context.sync();
      expect(shape.toJSON()).toEqual({
        name: 'Box',
        fill: { type: 'Solid', foregroundColor: '#4472C4', transparency: 0 },
      });
    });
  });

  test('$all in a tree loads the scalars of that level', async () => {
    const { runtime } = counting();
    await runtime.run(async (context) => {
      const shapes = context.presentation.slides.getItemAt(0).shapes;
      shapes.load({ $all: true, $top: 1 });
      await context.sync();
      expect(shapes.items).toHaveLength(1);
      expect(shapes.items[0]!.name).toBe('Box');
      expect(shapes.items[0]!.top).toBe(72);
    });
  });
});

describe('expand', () => {
  test('loads a navigation object with its default set', async () => {
    const { runtime } = counting();
    await runtime.run(async (context) => {
      const shapes = context.presentation.slides.getItemAt(0).shapes;
      shapes.load({ select: 'items/name', expand: 'items/fill' });
      await context.sync();
      expect(shapes.items[0]!.toJSON()).toEqual({
        name: 'Box',
        fill: { type: 'Solid', foregroundColor: '#4472C4', transparency: 0 },
      });
    });
  });

  test('a path that select names deeper loads only what select names', async () => {
    const { runtime } = counting();
    await runtime.run(async (context) => {
      const shape = context.presentation.slides.getItemAt(0).shapes.getItem('2');
      shape.load({ select: 'fill/type', expand: 'fill' });
      await context.sync();
      expect(shape.fill.toJSON()).toEqual({ type: 'Solid' });
    });
  });
});

describe('refusals', () => {
  test('a name that is not a property of the object it reaches refuses and queues nothing', async () => {
    const { runtime, batches } = counting();
    await runtime.run(async (context) => {
      const shape = context.presentation.slides.getItemAt(0).shapes.getItem('2');
      const before = batches.length;
      for (const bad of ['fill/nope', 'name/length', 'textFrame/textRange/nope/bold']) {
        const error = thrown(() => shape.load(bad));
        expect(code(error)).toBe('InvalidArgument');
      }
      expect(target(thrown(() => shape.load('fill/nope')))).toBe(
        'presentation.slides.getItemAt(0).shapes.getItem().fill.nope'
      );
      expect(target(thrown(() => shape.load('name/length')))).toBe(
        'presentation.slides.getItemAt(0).shapes.getItem().name'
      );
      await context.sync();
      // The lookups of getItemAt and getItem only; the refused loads left nothing.
      expect(batches.length - before).toBe(1);
      expect(batches.at(-1)).not.toContain('readFill');
    });
  });

  test('an unknown name suggests the nearest property and lists every one', async () => {
    const runtime = await open();
    await runtime.run(async (context) => {
      const shape = context.presentation.slides.getItemAt(0).shapes.getItem('2');
      const typo = thrown(() => shape.load('nmae'));
      if (!isPptxEditorError(typo)) throw typo;
      expect([typo.code, typo.debugInfo.innerError?.code]).toEqual([
        'InvalidArgument',
        'argument.not-accepted',
      ]);
      expect(typo.message).toContain("Did you mean 'name'?");
      expect(typo.message).toContain("'textFrame'");
      expect(typo.message).toContain("'parentGroup'");
      // A name far from every property gets the list only.
      const far = thrown(() => shape.load('somethingElse'));
      expect(isPptxEditorError(far) && far.message).not.toContain('Did you mean');
      // Below a path, the names are those of the object the path reaches.
      const deep = thrown(() => shape.load('textFrame/textRnage/text'));
      expect(isPptxEditorError(deep) && deep.message).toContain("Did you mean 'textRange'?");
    });
  });

  test('items on an object that is not a collection refuses', async () => {
    const { runtime } = counting();
    await runtime.run(async (context) => {
      const shape = context.presentation.slides.getItemAt(0).shapes.getItemAt(0);
      expect(code(thrown(() => shape.load('items/name')))).toBe('InvalidArgument');
      expect(code(thrown(() => shape.load({ items: { name: true } })))).toBe('InvalidArgument');
    });
  });

  test('a bad item path throws at the call, with the nearest name', async () => {
    const { runtime, batches } = counting();
    await runtime.run(async (context) => {
      const shapes = context.presentation.slides.getItemAt(0).shapes;
      const hint = (error: unknown): string | undefined =>
        isPptxEditorError(error) ? error.debugInfo.innerError?.message : undefined;
      const deep = thrown(() => shapes.load('items/fill/nope'));
      expect([code(deep), target(deep)]).toEqual([
        'InvalidArgument',
        'presentation.slides.getItemAt(0).shapes.items.fill.nope',
      ]);
      const typo = thrown(() => shapes.load('itmes/name'));
      expect([code(typo), target(typo)]).toEqual([
        'InvalidArgument',
        'presentation.slides.getItemAt(0).shapes.itmes',
      ]);
      expect(hint(typo)).toContain("Did you mean 'items'?");
      const name = thrown(() => shapes.load('items/nmae'));
      expect(hint(name)).toContain("Did you mean 'name'?");
      await context.sync();
    });
    // The calls queued nothing: the sync sends only the lookups of the slide.
    expect(batches).toEqual([['getPresentation', 'getSlide']]);
  });

  test('bounds and value types are checked as the option is read', async () => {
    const { runtime } = counting();
    await runtime.run(async (context) => {
      const shape = context.presentation.slides.getItemAt(0).shapes.getItemAt(0);
      const deep = Array.from({ length: 9 }, () => 'fill').join('/');
      expect(code(thrown(() => shape.load(deep)))).toBe('InvalidArgument');
      let tree: Record<string, unknown> = { name: true };
      for (let i = 0; i < 9; i += 1) tree = { fill: tree };
      expect(code(thrown(() => shape.load(tree as never)))).toBe('InvalidArgument');
      const wide = Array.from({ length: 257 }, () => 'name');
      expect(code(thrown(() => shape.load(wide)))).toBe('InvalidArgument');
      expect(code(thrown(() => shape.load({ name: 'yes' } as never)))).toBe('InvalidArgument');
      expect(code(thrown(() => shape.load({ $all: 1 } as never)))).toBe('InvalidArgument');
      expect(code(thrown(() => shape.load({ '__proto__/x': true } as never)))).toBe(
        'InvalidArgument'
      );
      expect(code(thrown(() => shape.load('$all/name')))).toBe('InvalidArgument');
      expect(code(thrown(() => shape.load('fill//type')))).toBe('InvalidArgument');
      expect(code(thrown(() => shape.load({ select: 'name', expand: 3 } as never)))).toBe(
        'InvalidArgument'
      );
    });
  });
});

describe('text paths over shapes with and without text', () => {
  const FIXTURES = join(import.meta.dir, '../../../../fixtures');

  test('the documented text path loads on the first slide of every fixture', async () => {
    const decks = readdirSync(FIXTURES).filter((name) => name.endsWith('.pptx'));
    expect(decks.length).toBeGreaterThan(10);
    for (const name of decks) {
      const runtime = await open(new Uint8Array(readFileSync(join(FIXTURES, name))));
      try {
        await runtime.run(async (context) => {
          const slide = context.presentation.slides.getItemAt(0);
          slide.load('shapes/items/textFrame/textRange/text');
          await context.sync();
          for (const item of slide.shapes.items) {
            const range = item.textFrame.textRange;
            const loaded = range.toJSON() as { text?: unknown };
            if (loaded.text !== undefined) expect(typeof loaded.text).toBe('string');
          }
        });
      } catch (error) {
        throw new Error(`${name}: ${String(error)}`);
      } finally {
        runtime.dispose();
      }
    }
  });

  /** A rectangle with no text body, a connector, and a table. */
  function mixed(): Uint8Array {
    const connector =
      '<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="3" name="Line"/><p:cNvCxnSpPr/><p:nvPr/></p:nvCxnSpPr>' +
      '<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="0"/></a:xfrm><a:prstGeom prst="line"><a:avLst/></a:prstGeom></p:spPr></p:cxnSp>';
    const grid = table({
      id: 4,
      x: 0,
      y: IN,
      cx: IN,
      cy: IN,
      columns: [IN],
      tblPr: '<a:tblPr/>',
      rows: `<a:tr h="370840">${cell('T')}</a:tr>`,
    });
    return buildPptx({
      slides: [
        {
          layout: 'blank',
          shapes: shape({ id: 2, name: 'Empty', x: 0, y: 0, cx: IN, cy: IN }) + connector + grid,
        },
      ],
    });
  }

  test('a shape with no text body reads as empty text; others are skipped on a collection', async () => {
    const runtime = await open(mixed());
    await runtime.run(async (context) => {
      const shapes = context.presentation.slides.getItemAt(0).shapes;
      shapes.load('items/name, items/textFrame/hasText, items/textFrame/textRange/text');
      await context.sync();
      const [empty, line, grid] = shapes.items;
      expect(empty!.textFrame.textRange.text).toBe('');
      expect(empty!.textFrame.hasText).toBe(false);
      // The connector and the table hold no text frame: their text is not loaded.
      expect(code(thrown(() => line!.textFrame.textRange.text))).toBe('PropertyNotLoaded');
      expect(code(thrown(() => grid!.textFrame.hasText))).toBe('PropertyNotLoaded');
      // The refusal says why: the shape has no text frame.
      const why = thrown(() => grid!.textFrame.hasText);
      expect(isPptxEditorError(why) && why.debugInfo.innerError?.code).toBe('edit.no-text-frame');
      expect(isPptxEditorError(why) && why.debugInfo.innerError?.message).toContain('text frame');
      // A property the load did not ask for says nothing more.
      const plain = thrown(() => empty!.textFrame.autoSizeSetting);
      expect(isPptxEditorError(plain) && plain.debugInfo.innerError).toBeUndefined();
      expect(line!.name).toBe('Line');
    });
  });

  test('a direct text read of a shape with no text frame still refuses', async () => {
    const runtime = await open(mixed());
    const error = await rejection(
      runtime.run(async (context) => {
        const line = context.presentation.slides.getItemAt(0).shapes.getItem('3');
        line.textFrame.textRange.load('text');
        await context.sync();
      })
    );
    expect(code(error)).toBe('NotSupported');
  });

  test('a write to a shape with no text body adds one', async () => {
    const runtime = await open(mixed());
    await runtime.run(async (context) => {
      const empty = context.presentation.slides.getItemAt(0).shapes.getItem('2');
      empty.textFrame.textRange.text = 'Now';
      await context.sync();
      const range = empty.textFrame.textRange.load('text');
      await context.sync();
      expect(range.text).toBe('Now');
    });
  });
});

describe('group paths over shapes', () => {
  /** A box at the top level, and a group of two boxes. */
  function grouped(): Uint8Array {
    const box = (id: number, name: string) =>
      shape({ id, name, x: id * IN, y: 0, cx: IN, cy: IN, text: para(run(name)) });
    const group =
      '<p:grpSp><p:nvGrpSpPr><p:cNvPr id="10" name="Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
      '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="5486400" cy="914400"/>' +
      '<a:chOff x="0" y="0"/><a:chExt cx="5486400" cy="914400"/></a:xfrm></p:grpSpPr>' +
      box(11, 'Inner A') +
      box(12, 'Inner B') +
      '</p:grpSp>';
    return buildPptx({ slides: [{ layout: 'blank', shapes: box(2, 'Top') + group }] });
  }

  test('group and parentGroup load through the items; other shapes are skipped', async () => {
    const runtime = await open(grouped());
    await runtime.run(async (context) => {
      const shapes = context.presentation.slides.getItemAt(0).shapes;
      shapes.load('items/name, items/group/shapes/items/name, items/parentGroup/name');
      await context.sync();
      const [top, group] = shapes.items;
      expect(group!.group.shapes.items.map((s) => s.name)).toEqual(['Inner A', 'Inner B']);
      expect(group!.parentGroup.isNullObject).toBe(true);
      expect(top!.parentGroup.isNullObject).toBe(true);
      // The box is no group: its members are not loaded, and the refusal says why.
      const why = thrown(() => top!.group.shapes.items);
      expect(code(why)).toBe('PropertyNotLoaded');
      expect(isPptxEditorError(why) && why.debugInfo.innerError?.code).toBe('edit.not-a-group');
      expect(code(thrown(() => top!.group.id))).toBe('PropertyNotLoaded');

      const inner = group!.group.shapes;
      inner.load('items/parentGroup/name');
      await context.sync();
      expect(inner.items.map((s) => s.parentGroup.name)).toEqual(['Group', 'Group']);
    });
  });

  test('a direct read keeps the Office.js refusals', async () => {
    const runtime = await open(grouped());
    const notGroup = await rejection(
      runtime.run(async (context) => {
        context.presentation.slides.getItemAt(0).shapes.getItem('2').group.load('id');
        await context.sync();
      })
    );
    expect(isPptxEditorError(notGroup) && notGroup.debugInfo.innerError?.code).toBe(
      'edit.not-a-group'
    );
    const topLevel = await rejection(
      runtime.run(async (context) => {
        context.presentation.slides.getItemAt(0).shapes.getItem('2').load('parentGroup');
        await context.sync();
      })
    );
    expect(code(topLevel)).toBe('ItemNotFound');
  });
});

describe('a shape with no text body reads through the cascade', () => {
  const ph = (id: number, body: string) =>
    `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="Title ${id}"/><p:cNvSpPr/><p:nvPr><p:ph type="ctrTitle"/></p:nvPr></p:nvSpPr><p:spPr/>${body}</p:sp>`;

  /** A layout title anchored at the bottom; one slide title has no body, one an empty body. */
  function inherited(): Uint8Array {
    return buildPptx({
      layouts: [
        {
          name: 'Anchored',
          type: 'title',
          shapes: `<p:sp><p:nvSpPr><p:cNvPr id="2" name="T"/><p:cNvSpPr/><p:nvPr><p:ph type="ctrTitle"/></p:nvPr></p:nvSpPr><p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="5000000" cy="1000000"/></a:xfrm></p:spPr><p:txBody><a:bodyPr anchor="b" lIns="0" wrap="none"><a:normAutofit/></a:bodyPr><a:lstStyle><a:lvl1pPr><a:defRPr sz="4400"/></a:lvl1pPr></a:lstStyle>${para(run('Prompt'))}</p:txBody></p:sp>`,
        },
      ],
      slides: [
        {
          layoutName: 'Anchored',
          shapes:
            ph(2, '') +
            ph(3, '<p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody>') +
            shape({ id: 4, name: 'Plain', x: 0, y: 2 * IN, cx: IN, cy: IN }),
        },
      ],
    } as Parameters<typeof buildPptx>[0]);
  }

  /** What a load of the text frame, the font, and the paragraph format of a shape reads. */
  async function described(context: RequestContext, id: string) {
    const frame = context.presentation.slides.getItemAt(0).shapes.getItem(id).textFrame;
    frame.load('verticalAlignment, leftMargin, wordWrap, autoSizeSetting');
    frame.textRange.font.load('size, name, bold');
    frame.textRange.paragraphFormat.load('horizontalAlignment');
    await context.sync();
    return {
      frame: frame.toJSON(),
      font: frame.textRange.font.toJSON(),
      paragraph: frame.textRange.paragraphFormat.toJSON(),
    };
  }

  test('a placeholder with no body reads what an empty body inherits', async () => {
    const runtime = await open(inherited());
    await runtime.run(async (context) => {
      const none = await described(context, '2');
      expect(none).toEqual(await described(context, '3'));
      expect(none.frame).toMatchObject({
        verticalAlignment: 'Bottom',
        leftMargin: 0,
        wordWrap: false,
      });
      expect(none.font.size).toBe(44);
    });
  });

  test('a write adds the body that the read described', async () => {
    const runtime = await open(inherited());
    await runtime.run(async (context) => {
      const before = { title: await described(context, '2'), plain: await described(context, '4') };
      const shapes = context.presentation.slides.getItemAt(0).shapes;
      shapes.getItem('2').textFrame.textRange.text = 'X';
      shapes.getItem('4').textFrame.textRange.text = 'X';
      await context.sync();
      expect({
        title: await described(context, '2'),
        plain: await described(context, '4'),
      }).toEqual(before);
      // A shape that is no placeholder gets plain body properties: its text is at the top.
      expect(before.plain.frame).toMatchObject({ verticalAlignment: 'Top' });
    });
    const saved = await runtime.save();
    const zip = readZip(saved);
    if (!zip.ok) throw new Error('saved');
    const slide = [...zip.entries].find(([name]) => name.endsWith('slides/slide1.xml'))!;
    const xml = new TextDecoder().decode(slide[1]);
    // Each shape now has a body with no body property attributes. (The title's inherited
    // autofit re-fits its text, which can add an autofit child.)
    expect(xml.match(/<p:txBody>/g)).toHaveLength(3);
    expect(xml.match(/<a:bodyPr(?:\/>|>)/g)).toHaveLength(3);
    expect(xml).not.toContain('anchor="ctr"');
  });
});
