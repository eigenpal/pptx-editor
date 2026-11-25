/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/editor-api/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// `Presentation.replaceAll`, an SDK extension over the editor's find (`commands/find.ts`): shapes,
// group children, table cells, and speaker notes; run properties kept; fields atomic; the same
// answer on a server runtime and over a live editor, with one undo step for the sync.

import { describe, expect, test } from 'bun:test';
import { strFromU8, unzipSync } from 'fflate';
import type { AutomationHost, AutomationOperation } from '@pptx-editor.dev/core/automation';
import {
  createBrowserAutomationHost,
  createPptxEditor,
  type PptxEditorInstance,
} from '@pptx-editor.dev/core/editor';
import {
  buildPptx,
  cell,
  group,
  para,
  run,
  shape,
  table,
} from '../../../../scripts/fixtures/builder.ts';
import { PptxEditor as BrowserEditor } from '../browser.ts';
import { isPptxEditorError, type RequestContext } from '../index.ts';
import { createRuntime } from '../runtime/runtime.ts';
import { IN, open, rejection, reopen } from './support.ts';

const FIELD =
  '<a:fld id="{B6F15528-21DE-4FAA-801E-634DDDAF4B2B}" type="slidenum"><a:rPr lang="en-US"/><a:t>1</a:t></a:fld>';

/**
 * Slide 1: a box with a bold "Acme" and a field, a group child, and a table. Slide 2: a box
 * with the word in other cases and inside a longer word, and speaker notes.
 */
function deck(): Uint8Array {
  return buildPptx({
    slides: [
      {
        layout: 'blank',
        shapes:
          shape({
            id: 2,
            name: 'Title',
            x: IN,
            y: IN,
            cx: 6 * IN,
            cy: IN,
            text: para(run('Acme', 'b="1" sz="3200"') + run(' plan, slide ') + FIELD),
          }) +
          group(
            { id: 3, x: IN, y: 3 * IN, cx: 2 * IN, cy: IN },
            shape({
              id: 4,
              name: 'Inner',
              x: IN,
              y: 3 * IN,
              cx: 2 * IN,
              cy: IN,
              text: para(run('by Acme')),
            })
          ) +
          table({
            id: 5,
            x: 4 * IN,
            y: 3 * IN,
            cx: 4 * IN,
            cy: IN,
            columns: [2 * IN, 2 * IN],
            tblPr: '<a:tblPr/>',
            rows: `<a:tr h="370840">${cell('Acme')}${cell('Other')}</a:tr>`,
          }),
        notes: 'Say Acme twice: Acme.',
      },
      {
        layout: 'blank',
        shapes: shape({
          id: 2,
          name: 'Body',
          x: IN,
          y: IN,
          cx: 6 * IN,
          cy: IN,
          text: para(run('ACME acme Acmeworks')),
        }),
      },
    ],
  });
}

const box = (context: RequestContext, slide: number, id: string) =>
  context.presentation.slides.getItemAt(slide).shapes.getItem(id);

/** The text of each place the deck holds the word, after a sync. */
async function texts(context: RequestContext) {
  const title = box(context, 0, '2').textFrame.textRange.load('text');
  const inner = box(context, 0, '3').group.shapes.getItem('4').textFrame.textRange.load('text');
  const tableCell = box(context, 0, '5').getTable().getCellOrNullObject(0, 0).load('text');
  const body = box(context, 1, '2').textFrame.textRange.load('text');
  await context.sync();
  return [title.text, inner.text, tableCell.text, body.text];
}

/** The notes text of the first slide, from saved bytes. */
function notesOf(bytes: Uint8Array): string {
  const xml = strFromU8(unzipSync(bytes)['ppt/notesSlides/notesSlide1.xml']!);
  return [...xml.matchAll(/<a:t>([^<]*)<\/a:t>/g)].map((m) => m[1]).join('');
}

async function liveEditor(): Promise<PptxEditorInstance> {
  const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await editor.load(deck())).ok).toBe(true);
  return editor;
}

describe('Presentation.replaceAll', () => {
  test('replaces in shapes, group children, table cells, and notes, and answers the count', async () => {
    const runtime = await open(deck());
    await runtime.run(async (context) => {
      const count = context.presentation.replaceAll('Acme', 'Globex', { matchCase: true });
      await context.sync();
      // Title, group child, table cell, two in the notes; slide 2 has "Acmeworks" only.
      expect(count.value).toBe(6);
      expect(await texts(context)).toEqual([
        'Globex plan, slide 1',
        'by Globex',
        'Globex',
        'ACME acme Globexworks',
      ]);
    });
    const saved = await runtime.save();
    expect(notesOf(saved)).toBe('Say Globex twice: Globex.');
    const again = await reopen(runtime);
    await again.run(async (context) => {
      expect((await texts(context))[0]).toBe('Globex plan, slide 1');
    });
  });

  test('the new text keeps the run properties of the text it replaces', async () => {
    const runtime = await open(deck());
    await runtime.run(async (context) => {
      context.presentation.replaceAll('Acme plan', 'Initech roadmap', { matchCase: true });
      await context.sync();
      const range = box(context, 0, '2').textFrame.textRange;
      const first = range.getSubstring(0, 7).font.load('bold,size');
      const rest = range.getSubstring(7, 8).font.load('bold,size');
      range.load('text');
      await context.sync();
      expect(range.text).toBe('Initech roadmap, slide 1');
      // The match started in the bold 32 pt run: all of the new text is that run's.
      expect([first.bold, first.size]).toEqual([true, 32]);
      expect([rest.bold, rest.size]).toEqual([true, 32]);
    });
  });

  test('matches any case by default, and whole words only on request', async () => {
    for (const [options, expected, body] of [
      [undefined, 8, 'X X Xworks'],
      [{ matchWholeWord: true }, 7, 'X X Acmeworks'],
      [{ matchCase: true, matchWholeWord: true }, 5, 'ACME acme Acmeworks'],
    ] as const) {
      const runtime = await open(deck());
      await runtime.run(async (context) => {
        const count = context.presentation.replaceAll('Acme', 'X', options);
        await context.sync();
        expect(count.value).toBe(expected);
        expect((await texts(context))[3]).toBe(body);
      });
    }
  });

  test('a field stays whole: a match never covers it', async () => {
    const runtime = await open(deck());
    await runtime.run(async (context) => {
      // The field shows "1" on slide 1; the search does not see into it.
      const field = context.presentation.replaceAll('slide 1', 'page');
      const around = context.presentation.replaceAll('slide ', 'page ');
      await context.sync();
      expect([field.value, around.value]).toEqual([0, 1]);
      expect((await texts(context))[0]).toBe('Acme plan, page 1');
    });
    const xml = strFromU8(unzipSync(await runtime.save())['ppt/slides/slide1.xml']!);
    expect(xml).toContain('type="slidenum"');
  });

  test('a line break is searched as \\v, and the replacement is one line', async () => {
    const bytes = buildPptx({
      slides: [
        {
          layout: 'blank',
          shapes: shape({
            id: 2,
            x: IN,
            y: IN,
            cx: 4 * IN,
            cy: IN,
            text: '<a:p><a:r><a:rPr lang="en-US"/><a:t>one</a:t></a:r><a:br><a:rPr lang="en-US"/></a:br><a:r><a:rPr lang="en-US"/><a:t>two</a:t></a:r></a:p>',
          }),
        },
      ],
    });
    const runtime = await open(bytes);
    await runtime.run(async (context) => {
      const count = context.presentation.replaceAll('one\vtwo', 'a\tb\vc');
      await context.sync();
      expect(count.value).toBe(1);
      const range = box(context, 0, '2').textFrame.textRange.load('text');
      await context.sync();
      expect(range.text).toBe('a b c');
    });
  });

  test('no match changes nothing and answers 0', async () => {
    const runtime = await open(deck());
    const before = await runtime.save();
    await runtime.run(async (context) => {
      const count = context.presentation.replaceAll('Nowhere', 'X');
      await context.sync();
      expect(count.value).toBe(0);
    });
    expect(await runtime.save()).toEqual(before);
  });

  test('refuses arguments it cannot honor, at the call', async () => {
    const runtime = await open(deck());
    const cases: [(context: RequestContext) => unknown, string][] = [
      [(c) => c.presentation.replaceAll('', 'x'), 'presentation.replaceAll.searchText'],
      [(c) => c.presentation.replaceAll('a\rb', 'x'), 'presentation.replaceAll.searchText'],
      [(c) => c.presentation.replaceAll('a\nb', 'x'), 'presentation.replaceAll.searchText'],
      [
        (c) => c.presentation.replaceAll('a'.repeat(1001), 'x'),
        'presentation.replaceAll.searchText',
      ],
      [(c) => c.presentation.replaceAll(1 as never, 'x'), 'presentation.replaceAll.searchText'],
      [
        (c) => c.presentation.replaceAll('a', 'x'.repeat(1001)),
        'presentation.replaceAll.replacementText',
      ],
      [
        (c) => c.presentation.replaceAll('a', null as never),
        'presentation.replaceAll.replacementText',
      ],
      [
        (c) => c.presentation.replaceAll('a', 'x', 'yes' as never),
        'presentation.replaceAll.options',
      ],
      [
        (c) => c.presentation.replaceAll('a', 'x', { matchcase: true } as never),
        'presentation.replaceAll.options.matchcase',
      ],
      [
        (c) => c.presentation.replaceAll('a', 'x', { matchWildcards: true } as never),
        'presentation.replaceAll.options.matchWildcards',
      ],
      [
        (c) => c.presentation.replaceAll('a', 'x', { matchCase: 1 } as never),
        'presentation.replaceAll.options.matchCase',
      ],
    ];
    // U+FFFC stands for a field in the search text, so it could match one.
    cases.push(
      [(c) => c.presentation.replaceAll('\uFFFC', 'x'), 'presentation.replaceAll.searchText'],
      [(c) => c.presentation.replaceAll('a\u0001', 'x'), 'presentation.replaceAll.searchText'],
      [(c) => c.presentation.replaceAll('a\uD83D', 'x'), 'presentation.replaceAll.searchText'],
      [(c) => c.presentation.replaceAll('a', 'x\u0001'), 'presentation.replaceAll.replacementText'],
      [(c) => c.presentation.replaceAll('a', '\uDE00'), 'presentation.replaceAll.replacementText']
    );
    for (const [call, target] of cases) {
      const error = await rejection(
        runtime.run(async (context) => {
          call(context);
          await context.sync();
        })
      );
      expect(isPptxEditorError(error) && [error.code, error.debugInfo.errorLocation]).toEqual([
        'InvalidArgument',
        target,
      ]);
    }
  });

  test('sends one replaceText operation for the call', async () => {
    const editor = await liveEditor();
    const real = createBrowserAutomationHost(editor);
    const sent: AutomationOperation[] = [];
    const host: AutomationHost = {
      ...real,
      capabilities: real.capabilities,
      revision: () => real.revision(),
      execute: (request) => {
        sent.push(...request.operations);
        return real.execute(request);
      },
    };
    const runtime = createRuntime({ host, save: false });
    await runtime.run(async (context) => {
      context.presentation.replaceAll('Acme', 'Globex', { matchCase: true, matchWholeWord: true });
      await context.sync();
    });
    // `packages/pro/src/__tests__/yjs-replace-text.test.ts` sends this operation to two peers.
    expect(sent.filter((op) => op.op !== 'getPresentation')).toEqual([
      {
        op: 'replaceText',
        query: 'Acme',
        replacement: 'Globex',
        options: { matchCase: true, wholeWord: true },
      },
    ]);
  });

  test('a live editor answers as a server runtime, and one undo step takes the sync back', async () => {
    const server = await open(deck());
    const editor = await liveEditor();
    const live = BrowserEditor.createBrowser(editor);
    const answers: unknown[] = [];
    for (const runtime of [server, live])
      await runtime.run(async (context) => {
        const count = context.presentation.replaceAll('acme', 'Globex', { matchWholeWord: true });
        await context.sync();
        answers.push([count.value, await texts(context)]);
      });
    expect(answers[1]).toEqual(answers[0]);
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    await live.run(async (context) => {
      expect(await texts(context)).toEqual([
        'Acme plan, slide 1',
        'by Acme',
        'Acme',
        'ACME acme Acmeworks',
      ]);
    });
  });

  test('text that grows re-fits its shape and its table, the same on both runtimes', async () => {
    const long = 'Acme Corporation of the Western Hemisphere and Beyond';
    const heights: number[][] = [];
    const runtimes = [await open(deck()), BrowserEditor.createBrowser(await liveEditor())];
    for (const runtime of runtimes)
      await runtime.run(async (context) => {
        const read = async () => {
          const title = box(context, 0, '2').load('height');
          const table = box(context, 0, '5').load('height');
          await context.sync();
          return [title.height, table.height];
        };
        const before = await read();
        context.presentation.replaceAll('Acme', long, { matchCase: true });
        await context.sync();
        const after = await read();
        // The box fits its text (`a:spAutoFit`), and the table row grows with its cell.
        expect(after[0]!).toBeGreaterThan(before[0]!);
        expect(after[1]!).toBeGreaterThan(before[1]!);
        heights.push(after);
      });
    expect(heights[1]).toEqual(heights[0]);
  });
});
