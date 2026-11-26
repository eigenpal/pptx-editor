/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/editor-api/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// `TextRange.search`, an SDK extension in the shape of `Range.search` of the DOCX Editor API:
// the matches of a range's text as a `TextRangeCollection`, by the rules of `replaceAll`, the
// same on a server runtime and over a live editor.

import { describe, expect, test } from 'bun:test';
import { createPptxEditor } from '@pptx-editor.dev/core/editor';
import { PptxEditor as BrowserEditor } from '../browser.ts';
import {
  isPptxEditorError,
  type PptxEditorRuntime,
  type PptxEditorServerRuntime,
  type RequestContext,
} from '../index.ts';
import { deck, open, rejection } from './support.ts';

const box = (context: RequestContext) =>
  context.presentation.slides.getItemAt(0).shapes.getItem('2');

/** Two paragraphs: "Acme and acme" with a line break and "ACME", then "acmeworks acme". */
const TEXT = 'Acme and acme\vACME\racmeworks acme';

async function written(): Promise<PptxEditorServerRuntime> {
  const runtime = await open();
  await runtime.run(async (context) => {
    box(context).textFrame.textRange.text = TEXT;
    await context.sync();
  });
  return runtime;
}

/** The text of each match, loaded on the matches. */
async function texts(
  context: RequestContext,
  matches: { items: { load(name: string): { text: string } }[] }
): Promise<string[]> {
  const loaded = matches.items.map((m) => m.load('text'));
  await context.sync();
  return loaded.map((m) => m.text);
}

describe('TextRange.search', () => {
  test('finds each match in text order, in either case by default', async () => {
    const runtime = await written();
    await runtime.run(async (context) => {
      const matches = box(context).textFrame.textRange.search('acme').load('items');
      const count = box(context).textFrame.textRange.search('acme').getCount();
      await context.sync();
      expect(count.value).toBe(5);
      // "acmeworks" holds a match too: matchWholeWord is off.
      expect(await texts(context, matches)).toEqual(['Acme', 'acme', 'ACME', 'acme', 'acme']);
    });
  });

  test('passes matchCase and matchWholeWord', async () => {
    const runtime = await written();
    await runtime.run(async (context) => {
      const range = box(context).textFrame.textRange;
      const exact = range.search('acme', { matchCase: true }).getCount();
      const words = range.search('acme', { matchWholeWord: true }).getCount();
      const both = range.search('acme', { matchCase: true, matchWholeWord: true }).getCount();
      await context.sync();
      expect([exact.value, words.value, both.value]).toEqual([3, 4, 2]);
    });
  });

  test('a match is a range of the shape text that reads and writes in place', async () => {
    const runtime = await written();
    await runtime.run(async (context) => {
      const range = box(context).textFrame.textRange;
      range.search('ACME', { matchCase: true }).getFirst().text = 'Globex';
      await context.sync();
      const all = range.load('text');
      await context.sync();
      expect(all.text).toBe('Acme and acme\vGlobex\racmeworks acme');
    });
  });

  test('writes of several matches hold from the last match to the first', async () => {
    const runtime = await open();
    await runtime.run(async (context) => {
      const range = box(context).textFrame.textRange;
      range.text = 'aa X bb X cc X';
      const matches = range.search('X', { matchCase: true }).load('items');
      await context.sync();
      for (const match of [...matches.items].reverse()) match.text = 'LONG';
      await context.sync();
      const all = range.load('text');
      await context.sync();
      expect(all.text).toBe('aa LONG bb LONG cc LONG');
    });
  });

  test('a range that a read finds searches after that read', async () => {
    const runtime = await written();
    await runtime.run(async (context) => {
      // To the end of the text: the span is known only after the read of the text.
      const tail = box(context).textFrame.textRange.getSubstring(14);
      const matches = tail.search('acme').load('items');
      await context.sync();
      expect(await texts(context, matches)).toEqual(['ACME', 'acme', 'acme']);
    });
  });

  test('a sub-range finds only the matches wholly inside it', async () => {
    const runtime = await written();
    await runtime.run(async (context) => {
      // "and acme\vAC": the match "ACME" crosses the end of the range.
      const sub = box(context).textFrame.textRange.getSubstring(5, 11);
      const matches = sub.search('acme').load('items');
      await context.sync();
      expect(await texts(context, matches)).toEqual(['acme']);
    });
  });

  test('finds a line break written as \\v, and never across paragraphs', async () => {
    const runtime = await written();
    await runtime.run(async (context) => {
      const range = box(context).textFrame.textRange;
      const lineBreak = range.search('acme\vACME', { matchCase: true }).getCount();
      const paragraph = range.search('ACME acmeworks').getCount();
      await context.sync();
      expect([lineBreak.value, paragraph.value]).toEqual([1, 0]);
    });
  });

  test('getItemAt, getFirst, and getLast name single matches', async () => {
    const runtime = await written();
    await runtime.run(async (context) => {
      const matches = box(context).textFrame.textRange.search('acme', { matchWholeWord: true });
      const second = matches.getItemAt(1).load('text');
      const first = matches.getFirst().load('text');
      const last = matches.getLast().load('text');
      await context.sync();
      expect([first.text, second.text, last.text]).toEqual(['Acme', 'acme', 'acme']);
    });
  });

  test('the search reads a write made earlier in the same batch', async () => {
    const runtime = await open();
    await runtime.run(async (context) => {
      const range = box(context).textFrame.textRange;
      range.text = 'one two one';
      const count = range.search('one').getCount();
      await context.sync();
      expect(count.value).toBe(2);
    });
  });

  test('no match: no items, a null object, and ItemNotFound for a required match', async () => {
    const runtime = await written();
    await runtime.run(async (context) => {
      const matches = box(context).textFrame.textRange.search('Globex').load('items');
      const none = box(context).textFrame.textRange.search('Globex').getFirstOrNullObject();
      await context.sync();
      expect(matches.items).toEqual([]);
      expect(none.isNullObject).toBe(true);
    });
    for (const call of [
      (c: RequestContext) => box(c).textFrame.textRange.search('Globex').getFirst().load('text'),
      (c: RequestContext) => box(c).textFrame.textRange.search('acme').getItemAt(5).load('text'),
    ]) {
      const error = await rejection(
        runtime.run(async (context) => {
          call(context);
          await context.sync();
        })
      );
      expect(isPptxEditorError(error) && error.code).toBe('ItemNotFound');
    }
  });

  test('refuses search text and options it cannot honor, and names them', async () => {
    const runtime = await written();
    const target = 'presentation.slides.getItemAt(0).shapes.getItem().textFrame.textRange';
    const cases: [(c: RequestContext) => unknown, string][] = [
      [(c) => box(c).textFrame.textRange.search(''), `${target}.search.searchText`],
      [(c) => box(c).textFrame.textRange.search('a\rb'), `${target}.search.searchText`],
      [(c) => box(c).textFrame.textRange.search('￼'), `${target}.search.searchText`],
      [(c) => box(c).textFrame.textRange.search('x'.repeat(1001)), `${target}.search.searchText`],
      [
        (c) => box(c).textFrame.textRange.search('a', { matchWildcards: true } as never),
        `${target}.search.options.matchWildcards`,
      ],
      [
        (c) =>
          box(c)
            .textFrame.textRange.search('acme')
            .load('text' as never),
        `${target}.search().text`,
      ],
    ];
    for (const [call, location] of cases) {
      const error = await rejection(
        runtime.run(async (context) => {
          call(context);
          await context.sync();
        })
      );
      expect(isPptxEditorError(error) && [error.code, error.debugInfo.errorLocation]).toEqual([
        'InvalidArgument',
        location,
      ]);
    }
  });

  test('a live editor answers as a server runtime', async () => {
    const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    expect((await editor.load(deck())).ok).toBe(true);
    const runtimes: PptxEditorRuntime[] = [await open(), BrowserEditor.createBrowser(editor)];
    const answers: unknown[] = [];
    for (const runtime of runtimes)
      await runtime.run(async (context) => {
        box(context).textFrame.textRange.text = TEXT;
        const matches = box(context).textFrame.textRange.search('acme', { matchWholeWord: true });
        matches.load('items');
        await context.sync();
        answers.push(await texts(context, matches));
      });
    expect(answers[0]).toEqual(['Acme', 'acme', 'ACME', 'acme']);
    expect(answers[1]).toEqual(answers[0]);
  });
});
