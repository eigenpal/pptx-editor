/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/editor-api/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// `ShapeFont.highlightColor`, an SDK extension over the run highlight (`a:highlight`, CT_Color,
// ECMA-376 Part 1, 21.1.2.3.4) in the run properties (`a:rPr`, CT_TextCharacterProperties,
// ECMA-376 Part 1, 21.1.2.3.9): read, set, and remove on a text range, a text frame, and a table
// cell, in the server and the browser runtime.

import { describe, expect, test } from 'bun:test';
import { createStoreAutomationHost, type AutomationHost } from '@pptx-editor.dev/core/automation';
import { createPptxEditor } from '@pptx-editor.dev/core/editor';
import { PptxDocumentStore, readPptx, readZip } from '@pptx-editor.dev/core/store';
import { buildPptx, para, run, shape } from '../../../../scripts/fixtures/builder.ts';
import { isPptxEditorError, PptxEditor, type RequestContext } from '../browser.ts';
import { createRuntime } from '../runtime/runtime.ts';
import { IN, open, refusedAtSync, reopen } from './support.ts';

const YELLOW = '<a:highlight><a:srgbClr val="FFFF00"/></a:highlight>';

/** One box: "Marked" with a highlight, bold, a red fill, and a typeface, then " plain". */
const highlighted = (): Uint8Array =>
  buildPptx({
    slides: [
      {
        layout: 'blank',
        shapes: shape({
          id: 2,
          name: 'Box',
          x: IN,
          y: IN,
          cx: 4 * IN,
          cy: IN,
          text: para(
            run(
              'Marked',
              'b="1" sz="2000"',
              `<a:solidFill><a:srgbClr val="C00000"/></a:solidFill>${YELLOW}<a:latin typeface="Carlito"/>`
            ) + run(' plain')
          ),
        }),
      },
      { layout: 'blank', shapes: '' },
    ],
  });

/** The XML of the first slide in saved bytes. */
function slideXml(bytes: Uint8Array): string {
  const zip = readZip(bytes);
  if (!zip.ok) throw new Error('saved');
  return new TextDecoder().decode(new Map(zip.entries).get('/ppt/slides/slide1.xml'));
}

/** The `a:rPr` elements of a slide, in document order. */
const runProps = (xml: string): string[] => xml.match(/<a:rPr\b[^>]*?(?:\/>|>.*?<\/a:rPr>)/g) ?? [];

const box = (context: RequestContext) =>
  context.presentation.slides.getItemAt(0).shapes.getItem('2');

describe('ShapeFont.highlightColor', () => {
  test('reads the color, null for no highlight, and null for a mixed range', async () => {
    const runtime = await open(highlighted());
    const read = await runtime.run(async (context) => {
      const range = box(context).textFrame.textRange;
      const marked = range.getSubstring(0, 6).font.load('highlightColor');
      const plain = range.getSubstring(6, 6).font.load('highlightColor');
      const mixed = range.font.load('highlightColor');
      await context.sync();
      return [marked.highlightColor, plain.highlightColor, mixed.highlightColor];
    });
    expect(read).toEqual(['#FFFF00', null, null]);
  });

  test('a color sets a:highlight on the range only, in schema order', async () => {
    const runtime = await open(highlighted());
    await runtime.run(async (context) => {
      box(context).textFrame.textRange.getSubstring(7, 5).font.highlightColor = '#00ff00';
      await context.sync();
    });
    const props = runProps(slideXml(await runtime.save()));
    // " plain" splits at "p": the space keeps no highlight, "plain" takes the new one.
    expect(props).toHaveLength(3);
    expect(props[0]).toContain(YELLOW);
    expect(props[1]).not.toContain('a:highlight');
    expect(props[2]).toContain('<a:highlight><a:srgbClr val="00FF00"/></a:highlight>');
    const again = await reopen(runtime);
    const read = await again.run(async (context) => {
      const font = box(context).textFrame.textRange.getSubstring(7, 5).font.load('highlightColor');
      await context.sync();
      return font.highlightColor;
    });
    expect(read).toBe('#00FF00');
  });

  test('a preset color name writes a:prstClr and reads back as #RRGGBB', async () => {
    const runtime = await open(highlighted());
    const read = await runtime.run(async (context) => {
      const font = box(context).textFrame.textRange.getSubstring(6, 6).font;
      font.highlightColor = 'cyan';
      font.load('highlightColor');
      await context.sync();
      return font.highlightColor;
    });
    expect(read).toBe('#00FFFF');
    expect(slideXml(await runtime.save())).toContain(
      '<a:highlight><a:prstClr val="cyan"/></a:highlight>'
    );
  });

  test('null removes the highlight and keeps every other run property', async () => {
    const runtime = await open(highlighted());
    const before = runProps(slideXml(await (await open(highlighted())).save()));
    await runtime.run(async (context) => {
      box(context).textFrame.textRange.font.highlightColor = null;
      await context.sync();
    });
    const xml = slideXml(await runtime.save());
    expect(xml).not.toContain('highlight');
    // The first run keeps its attributes, its fill, and its typeface, in order.
    expect(runProps(xml)).toEqual([before[0]!.replace(YELLOW, ''), before[1]!]);
    const again = await reopen(runtime);
    const read = await again.run(async (context) => {
      const font = box(context)
        .textFrame.textRange.getSubstring(0, 6)
        .font.load('highlightColor,bold,color,name,size');
      await context.sync();
      return [font.highlightColor, font.bold, font.color, font.name, font.size];
    });
    expect(read).toEqual([null, true, '#C00000', 'Carlito', 20]);
  });

  test('a text frame write marks every run, and its read returns the shared color', async () => {
    const runtime = await open(highlighted());
    const read = await runtime.run(async (context) => {
      const font = box(context).textFrame.textRange.font;
      font.highlightColor = '#FFC000';
      font.load('highlightColor');
      await context.sync();
      return font.highlightColor;
    });
    expect(read).toBe('#FFC000');
    const props = runProps(slideXml(await runtime.save()));
    expect(
      props.every((p) => p.includes('<a:highlight><a:srgbClr val="FFC000"/></a:highlight>'))
    ).toBe(true);
    expect(props.join('')).not.toContain('FFFF00');
  });

  test('works on a table cell and survives a save', async () => {
    const runtime = await open();
    await runtime.run(async (context) => {
      const table = context.presentation.slides
        .getItemAt(1)
        .shapes.addTable(2, 2, { left: 20, top: 20 })
        .getTable();
      await context.sync();
      const cell = table.getCellOrNullObject(0, 1);
      cell.text = 'Due';
      cell.font.highlightColor = '#FFFF00';
      await context.sync();
    });
    const again = await reopen(runtime);
    const read = await again.run(async (context) => {
      const table = context.presentation.slides.getItemAt(1).shapes.getItemAt(0).getTable();
      const marked = table.getCellOrNullObject(0, 1).font.load('highlightColor');
      const plain = table.getCellOrNullObject(0, 0).font.load('highlightColor');
      await context.sync();
      const before = marked.highlightColor;
      table.getCellOrNullObject(0, 1).font.highlightColor = null;
      marked.load('highlightColor');
      await context.sync();
      return [before, plain.highlightColor, marked.highlightColor];
    });
    expect(read).toEqual(['#FFFF00', null, null]);
  });

  test('a removal adds no run properties to a run that has none', async () => {
    const bare = buildPptx({
      slides: [
        {
          layout: 'blank',
          shapes: shape({
            id: 2,
            name: 'Box',
            x: IN,
            y: IN,
            cx: 4 * IN,
            cy: IN,
            text: para(
              '<a:r><a:t>Hallo </a:t></a:r>' +
                run('Welt', '', YELLOW) +
                '<a:br/><a:r><a:t>x</a:t></a:r>'
            ),
          }),
        },
      ],
    });
    const runtime = await open(bare);
    await runtime.run(async (context) => {
      box(context).textFrame.textRange.font.highlightColor = null;
      await context.sync();
    });
    const xml = slideXml(await runtime.save());
    expect(xml).not.toContain('highlight');
    // The bare runs keep no a:rPr, so they keep the language they inherit.
    expect(xml).toContain('<a:r><a:t>Hallo </a:t></a:r>');
    expect(xml).toContain('<a:br/><a:r><a:t>x</a:t></a:r>');
  });

  test('a removal at a caret in text with no highlight leaves the file unchanged', async () => {
    const runtime = await open(highlighted());
    await runtime.run(async (context) => {
      box(context).textFrame.textRange.getSubstring(9, 0).font.highlightColor = null;
      await context.sync();
    });
    expect(slideXml(await runtime.save())).toBe(slideXml(await (await open(highlighted())).save()));
  });

  test('a removal at a caret after a highlighted run arms text with no highlight', async () => {
    const runtime = await open(highlighted());
    await runtime.run(async (context) => {
      box(context).textFrame.textRange.getSubstring(6, 0).font.highlightColor = null;
      await context.sync();
    });
    const props = runProps(slideXml(await runtime.save()));
    // "Marked", the empty run armed at the caret (it copies "Marked" less the highlight), " plain".
    expect(props).toHaveLength(3);
    expect(props[0]).toContain(YELLOW);
    expect(props[1]).toBe(props[0]!.replace(YELLOW, ''));
    expect(props[2]).not.toContain('a:highlight');
  });

  for (const syncs of [1, 2])
    test(`a removal at a caret clears a highlight armed there, in ${syncs} sync(s)`, async () => {
      const runtime = await open(highlighted());
      await runtime.run(async (context) => {
        const caret = box(context).textFrame.textRange.getSubstring(9, 0);
        caret.font.highlightColor = '#00FF00';
        if (syncs === 2) await context.sync();
        caret.font.highlightColor = null;
        await context.sync();
      });
      expect(slideXml(await runtime.save())).not.toContain('00FF00');
    });

  test('a write with other font properties keeps the schema order and removes a field highlight', async () => {
    const fielded = buildPptx({
      slides: [
        {
          layout: 'blank',
          shapes: shape({
            id: 2,
            name: 'Box',
            x: IN,
            y: IN,
            cx: 4 * IN,
            cy: IN,
            text: para(
              run('Page ') +
                `<a:fld id="{B6F15528-21DE-4FAA-801E-634DDDAF4B2B}" type="slidenum"><a:rPr lang="en-US">${YELLOW}</a:rPr><a:t>1</a:t></a:fld>`
            ),
          }),
        },
      ],
    });
    const runtime = await open(fielded);
    await runtime.run(async (context) => {
      const range = box(context).textFrame.textRange;
      range.getSubstring(5, 1).font.highlightColor = null;
      const font = range.getSubstring(0, 4).font;
      font.highlightColor = '#00FF00';
      font.color = '#C00000';
      font.name = 'Carlito';
      await context.sync();
    });
    const props = runProps(slideXml(await runtime.save()));
    expect(props[0]).toBe(
      '<a:rPr lang="en-US" dirty="0"><a:solidFill><a:srgbClr val="C00000"/></a:solidFill>' +
        '<a:highlight><a:srgbClr val="00FF00"/></a:highlight><a:latin typeface="Carlito"/>' +
        '<a:ea typeface="Carlito"/><a:cs typeface="Carlito"/></a:rPr>'
    );
    expect(props.join('')).not.toContain('FFFF00');
  });

  test('a highlight from the list style still reads after a removal', async () => {
    const inherited = buildPptx({
      slides: [
        {
          layout: 'blank',
          shapes: shape({
            id: 2,
            name: 'Box',
            x: IN,
            y: IN,
            cx: 4 * IN,
            cy: IN,
            lstStyle: `<a:lvl1pPr><a:defRPr>${YELLOW}</a:defRPr></a:lvl1pPr>`,
            text: para(run('Styled')),
          }),
        },
      ],
    });
    const runtime = await open(inherited);
    const read = await runtime.run(async (context) => {
      const font = box(context).textFrame.textRange.font;
      font.highlightColor = null;
      const loaded = font.load('highlightColor');
      await context.sync();
      return loaded.highlightColor;
    });
    expect(read).toBe('#FFFF00');
    // A removal at a caret in that text has no highlight of its own to remove: no change.
    const before = slideXml(await runtime.save());
    await runtime.run(async (context) => {
      box(context).textFrame.textRange.getSubstring(3, 0).font.highlightColor = null;
      await context.sync();
    });
    expect(slideXml(await runtime.save())).toBe(before);
  });

  test('an empty paragraph keeps the highlight on its paragraph mark', async () => {
    const empty = buildPptx({
      slides: [
        {
          layout: 'blank',
          shapes: shape({
            id: 2,
            name: 'Box',
            x: IN,
            y: IN,
            cx: 4 * IN,
            cy: IN,
            text: para(run('Text')) + '<a:p><a:endParaRPr lang="en-US"/></a:p>',
          }),
        },
      ],
    });
    const runtime = await open(empty);
    const read = await runtime.run(async (context) => {
      const range = box(context).textFrame.textRange;
      range.font.highlightColor = '#FFFF00';
      await context.sync();
      // "Text" and the paragraph break: the empty paragraph starts at offset 5.
      const mark = range.getSubstring(5, 0).font.load('highlightColor');
      await context.sync();
      return mark.highlightColor;
    });
    expect(read).toBe('#FFFF00');
    expect(slideXml(await runtime.save())).toContain(
      `<a:endParaRPr lang="en-US">${YELLOW}</a:endParaRPr>`
    );
  });

  test('reads a scheme color highlight as #RRGGBB', async () => {
    const scheme = buildPptx({
      slides: [
        {
          layout: 'blank',
          shapes: shape({
            id: 2,
            name: 'Box',
            x: IN,
            y: IN,
            cx: 4 * IN,
            cy: IN,
            text: para(
              run('Accent', '', '<a:highlight><a:schemeClr val="accent1"/></a:highlight>')
            ),
          }),
        },
      ],
    });
    const runtime = await open(scheme);
    const read = await runtime.run(async (context) => {
      const font = box(context).textFrame.textRange.font.load('highlightColor');
      await context.sync();
      return font.highlightColor;
    });
    expect(read).toMatch(/^#[0-9A-F]{6}$/);
  });

  test('refuses a value that is not a color', async () => {
    const runtime = await open(highlighted());
    for (const value of ['#FFF', 'not-a-color', 42]) {
      const error = await runtime.run((context) =>
        refusedAtSync(context, () => {
          box(context).textFrame.textRange.font.highlightColor = value as string;
        })
      );
      expect(isPptxEditorError(error) && error.code).toBe('InvalidArgument');
    }
    // Nothing changed.
    expect(slideXml(await runtime.save())).toBe(slideXml(await (await open(highlighted())).save()));
  });

  test('the server and the browser runtime write the same run properties', async () => {
    const edit = async (context: RequestContext) => {
      const range = box(context).textFrame.textRange;
      range.getSubstring(0, 6).font.highlightColor = null;
      range.getSubstring(7, 5).font.highlightColor = '#92D050';
      await context.sync();
    };
    const server = await open(highlighted());
    await server.run(edit);
    const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    expect((await editor.load(highlighted())).ok).toBe(true);
    await PptxEditor.createBrowser(editor).run(edit);
    const browserXml = slideXml(editor.save()!);
    expect(runProps(browserXml)).toEqual(runProps(slideXml(await server.save())));
    expect(browserXml).toContain('<a:highlight><a:srgbClr val="92D050"/></a:highlight>');
  });

  // The collaboration test in `packages/pro` sends this exact sequence to two peers.
  test('emits one setFont operation with highlightColor for a set and for a removal', async () => {
    const read = readPptx(highlighted());
    if (!read.ok) throw new Error('fixture');
    const inner = createStoreAutomationHost(new PptxDocumentStore(read.pkg));
    const sent: unknown[] = [];
    const host: AutomationHost = {
      ...inner,
      execute: (request) => {
        for (const operation of request.operations)
          if (operation.op === 'setFont') sent.push(operation);
        return inner.execute(request);
      },
    };
    await createRuntime({ host, save: true }).run(async (context) => {
      const range = box(context).textFrame.textRange;
      range.getSubstring(7, 5).font.highlightColor = '#00B0F0';
      range.font.highlightColor = null;
      await context.sync();
    });
    expect(sent).toEqual([
      {
        op: 'setFont',
        shape: expect.anything(),
        span: { start: 7, length: 5 },
        font: { highlightColor: '#00B0F0' },
      },
      { op: 'setFont', shape: expect.anything(), font: { highlightColor: null } },
    ]);
  });
});
