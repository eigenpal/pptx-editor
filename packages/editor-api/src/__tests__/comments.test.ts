/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/editor-api/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// `Slide.comments`, an SDK extension over both comment formats: the legacy comments of ECMA-376
// Part 1, 19.4 (`p:cm`, `CT_Comment`, 19.4.1, in a `p:cmLst`, 19.4.4) and the threaded comment
// parts that a slide names in its `p:extLst` (19.3.1.20).

import { describe, expect, test } from 'bun:test';
import type { AutomationHost, AutomationOperation } from '@pptx-editor.dev/core/automation';
import {
  createBrowserAutomationHost,
  createPptxEditor,
  readCommentThreads,
} from '@pptx-editor.dev/core/editor';
import { presentationInfo, readPptx } from '@pptx-editor.dev/core/store';
import { buildPptx, para, run, shape } from '../../../../scripts/fixtures/builder.ts';
import { buildSlideStructureDeck } from '../../../../scripts/fixtures/slide-structure.ts';
import { PptxEditor as BrowserEditor } from '../browser.ts';
import {
  isPptxEditorError,
  PptxEditor,
  type Comment,
  type PptxEditorRuntime,
  type RequestContext,
} from '../index.ts';
import { createRuntime } from '../runtime/runtime.ts';
import { refusedAtSync } from './support.ts';

const IN = 914_400;

/** Two slides; the first has a text box with two paragraphs. */
const plainDeck = (): Uint8Array =>
  buildPptx({
    slides: [
      {
        layout: 'blank',
        shapes: shape({
          id: 2,
          name: 'Title',
          x: IN,
          y: IN,
          cx: 4 * IN,
          cy: IN,
          text: para(run('Quarterly plan')) + para(run('Draft')),
        }),
      },
      { layout: 'blank', shapes: '' },
    ],
  });

const server = (bytes = plainDeck(), author: string | null = 'Review Agent') =>
  PptxEditor.createServer(bytes, { substitutes: false, ...(author ? { author } : {}) });

/** Every thread of the first slide, as plain values. */
async function readAll(runtime: PptxEditorRuntime, index = 0) {
  return runtime.run(async (context) => {
    const comments = context.presentation.slides.getItemAt(index).comments;
    comments.load(
      'items/id,items/authorName,items/authorInitials,items/content,items/resolved,items/anchor'
    );
    await context.sync();
    const replies = comments.items.map((c) => c.replies.load('items/id,items/content'));
    await context.sync();
    return comments.items.map((c, i) => ({
      id: c.id,
      authorName: c.authorName,
      authorInitials: c.authorInitials,
      content: c.content,
      resolved: c.resolved,
      anchor: c.anchor,
      replies: replies[i]!.items.map((r) => r.content),
    }));
  });
}

async function codeOf(
  runtime: PptxEditorRuntime,
  call: (context: RequestContext) => unknown
): Promise<[string | undefined, string | undefined]> {
  const error = await runtime.run((context) => refusedAtSync(context, () => call(context)));
  if (!isPptxEditorError(error)) throw error;
  return [error.code, error.debugInfo.innerError?.code];
}

describe('Slide.comments: threaded comments', () => {
  test('a slide with no comments has none', async () => {
    const runtime = await server();
    expect(await readAll(runtime)).toEqual([]);
    expect(
      await runtime.run(async (context) => {
        const count = context.presentation.slides.getItemAt(0).comments.getCount();
        await context.sync();
        return count.value;
      })
    ).toBe(0);
    runtime.dispose();
  });

  test('add on the slide, on a shape, and on a text range, then read them back', async () => {
    const runtime = await server();
    const ids = await runtime.run(async (context) => {
      const slide = context.presentation.slides.getItemAt(0);
      const title = slide.shapes.getItemAt(0);
      const a = slide.comments.add('Overall: good.');
      const b = slide.comments.add('Shorten this.', { shape: title });
      const c = slide.comments.add('Say "Q3".', {
        shape: '2',
        textRange: { start: 0, length: 9 },
        authorName: 'Ada Lovelace',
        authorInitials: 'ADA',
      });
      await context.sync();
      return [a.id, b.id, c.id];
    });
    expect(ids.every((id) => /^\{[0-9A-F-]{36}\}$/.test(id))).toBe(true);
    const all = await readAll(runtime);
    expect(all).toEqual([
      {
        id: ids[0],
        authorName: 'Review Agent',
        authorInitials: 'RA',
        content: 'Overall: good.',
        resolved: false,
        anchor: { type: 'slide' },
        replies: [],
      },
      {
        id: ids[1],
        authorName: 'Review Agent',
        authorInitials: 'RA',
        content: 'Shorten this.',
        resolved: false,
        anchor: { type: 'shape', shapeId: '2' },
        replies: [],
      },
      {
        id: ids[2],
        authorName: 'Ada Lovelace',
        authorInitials: 'ADA',
        content: 'Say "Q3".',
        resolved: false,
        anchor: { type: 'text', shapeId: '2', start: 0, length: 9 },
        replies: [],
      },
    ]);
    // Save and reopen keep every thread, author, and anchor.
    const again = await PptxEditor.createServer(await runtime.save(), { substitutes: false });
    expect(await readAll(again)).toEqual(all);
    runtime.dispose();
    again.dispose();
  });

  test('reply, edit, resolve, reopen, and delete', async () => {
    const runtime = await server();
    const id = await runtime.run(async (context) => {
      const added = context.presentation.slides.getItemAt(0).comments.add('Check the numbers.');
      await context.sync();
      return added.id;
    });
    await runtime.run(async (context) => {
      const comment = context.presentation.slides.getItemAt(0).comments.getItem(id);
      comment.reply('Checked.\nAll fine.', { authorName: 'Grace Hopper' });
      comment.content = 'Check the numbers, please.';
      comment.resolved = true;
      await context.sync();
    });
    let [thread] = await readAll(runtime);
    expect(thread).toMatchObject({
      content: 'Check the numbers, please.',
      resolved: true,
      replies: ['Checked.\nAll fine.'],
    });
    const replyAuthor = await runtime.run(async (context) => {
      const reply = context.presentation.slides
        .getItemAt(0)
        .comments.getItem(id)
        .replies.getItemAt(0)
        .load('authorName,authorInitials,creationDate');
      await context.sync();
      return [reply.authorName, reply.authorInitials, reply.creationDate instanceof Date];
    });
    expect(replyAuthor).toEqual(['Grace Hopper', 'GH', true]);
    await runtime.run(async (context) => {
      const comment = context.presentation.slides.getItemAt(0).comments.getItem(id);
      comment.resolved = false;
      comment.replies.getItemAt(0).delete();
      await context.sync();
    });
    [thread] = await readAll(runtime);
    expect(thread).toMatchObject({ resolved: false, replies: [] });
    await runtime.run(async (context) => {
      context.presentation.slides.getItemAt(0).comments.getItemAt(0).delete();
      await context.sync();
    });
    expect(await readAll(runtime)).toEqual([]);
    const saved = readPptx(await runtime.save());
    expect(saved.ok).toBe(true);
    runtime.dispose();
  });

  test('a new comment and its reply are addressable after the sync that adds them', async () => {
    const runtime = await server();
    const json = await runtime.run(async (context) => {
      const comment = context.presentation.slides.getItemAt(0).comments.add('First');
      await context.sync();
      const reply = comment.reply('Second');
      await context.sync();
      comment.load();
      reply.load('content');
      await context.sync();
      return [comment.toJSON(), reply.content];
    });
    expect(json[0]).toMatchObject({ content: 'First', resolved: false, anchor: { type: 'slide' } });
    expect(json[1]).toBe('Second');
    runtime.dispose();
  });

  test('getItemOrNullObject answers a null object for an unknown id', async () => {
    const runtime = await server();
    const isNull = await runtime.run(async (context) => {
      const found = context.presentation.slides.getItemAt(0).comments.getItemOrNullObject('nope');
      await context.sync();
      return found.isNullObject;
    });
    expect(isNull).toBe(true);
    runtime.dispose();
  });
});

describe('Slide.comments: author identity', () => {
  test('a server runtime without an author option refuses an add with no authorName', async () => {
    const runtime = await server(plainDeck(), null);
    expect(
      await codeOf(runtime, (context) =>
        context.presentation.slides.getItemAt(0).comments.add('Hi')
      )
    ).toEqual(['NotSupported', 'edit.no-comment-author']);
    await runtime.run(async (context) => {
      context.presentation.slides.getItemAt(0).comments.add('Hi', { authorName: 'Lin Yu' });
      await context.sync();
    });
    expect((await readAll(runtime))[0]).toMatchObject({ authorName: 'Lin Yu' });
    runtime.dispose();
  });

  test('an author option that is not a name refuses at creation', async () => {
    for (const author of ['', '   ', '\u0007\u0007', 'x'.repeat(256), 7]) {
      const error = await PptxEditor.createServer(plainDeck(), {
        substitutes: false,
        author: author as string,
      }).catch((e: unknown) => e);
      expect(isPptxEditorError(error) && error.code).toBe('InvalidArgument');
    }
  });

  test('text and author names stay plain text', async () => {
    const runtime = await server();
    await runtime.run(async (context) => {
      context.presentation.slides
        .getItemAt(0)
        .comments.add('<img src=x onerror=alert(1)> & "quotes"', { authorName: '<b>Bold</b>' });
      await context.sync();
    });
    const again = await PptxEditor.createServer(await runtime.save(), { substitutes: false });
    expect((await readAll(again))[0]).toMatchObject({
      authorName: '<b>Bold</b>',
      content: '<img src=x onerror=alert(1)> & "quotes"',
    });
    runtime.dispose();
    again.dispose();
  });
});

describe('Slide.comments: refusals', () => {
  test('invalid input refuses with a clear code and writes nothing', async () => {
    const runtime = await server();
    const slide = (context: RequestContext) => context.presentation.slides.getItemAt(0);
    const cases: [(context: RequestContext) => unknown, string, string | undefined][] = [
      [(c) => slide(c).comments.add(''), 'InvalidArgument', 'edit.comment-text'],
      [(c) => slide(c).comments.add('x'.repeat(10_001)), 'InvalidArgument', 'edit.comment-text'],
      [(c) => slide(c).comments.add(7 as unknown as string), 'InvalidArgument', undefined],
      [
        (c) => slide(c).comments.add('x', { authorName: 'y'.repeat(256) }),
        'InvalidArgument',
        'edit.comment-author',
      ],
      [
        (c) => slide(c).comments.add('x', { textRange: { start: 0, length: 1 } }),
        'InvalidArgument',
        undefined,
      ],
      [
        (c) => slide(c).comments.add('x', { shape: '2', textRange: { start: 40, length: 2 } }),
        'InvalidArgument',
        undefined,
      ],
      [(c) => slide(c).comments.add('x', { shape: '99' }), 'ItemNotFound', undefined],
      [(c) => slide(c).comments.getItem('nope').delete(), 'ItemNotFound', undefined],
    ];
    for (const [i, [call, code, diagnostic]] of cases.entries()) {
      const [got, inner] = await codeOf(runtime, (context) => {
        // A valid add first: a refusal later in the same sync must drop it too.
        slide(context).comments.add('Kept?');
        call(context);
      });
      expect([i, got]).toEqual([i, code]);
      if (diagnostic) expect([i, inner]).toEqual([i, diagnostic]);
      expect([i, await readAll(runtime)]).toEqual([i, []]);
    }
    runtime.dispose();
  });

  test('a tracked comment refuses a write after another call changed the slide', async () => {
    const runtime = await server();
    const comment = await runtime.run(async (context) => {
      const slide = context.presentation.slides.getItemAt(0);
      const added = slide.comments.add('Original');
      await context.sync();
      // The slide is addressable now, so the comment has a path of its own to track.
      const found = slide.comments.getItem(added.id);
      found.load('content');
      await context.sync();
      context.trackedObjects.add(found);
      return found;
    });
    // Another call deletes the thread: the tracked proxy read an older revision.
    await runtime.run(async (context) => {
      context.presentation.slides.getItemAt(0).comments.getItemAt(0).delete();
      await context.sync();
    });
    const error = await runtime.run(comment, (context) =>
      refusedAtSync(context, () => {
        comment.resolved = true;
      })
    );
    expect(isPptxEditorError(error) && error.code).toBe('StaleDocument');
    runtime.dispose();
  });
});

describe('Slide.comments: scope of a proxy', () => {
  test('a write through a proxy never reaches a comment outside its thread', async () => {
    const runtime = await server();
    const [first, second, reply] = await runtime.run(async (context) => {
      const comments = context.presentation.slides.getItemAt(0).comments;
      const a = comments.add('First');
      const b = comments.add('Second');
      await context.sync();
      const r = a.reply('Reply to first');
      await context.sync();
      return [a.id, b.id, r.id];
    });
    const cases: ((context: RequestContext) => void)[] = [
      // The replies of the first thread do not hold the second thread.
      (c) =>
        c.presentation.slides.getItemAt(0).comments.getItem(first).replies.getItem(second).delete(),
      // A reply id is not a thread of the slide.
      (c) => {
        c.presentation.slides.getItemAt(0).comments.getItem(reply).resolved = true;
      },
      (c) => void c.presentation.slides.getItemAt(0).comments.getItem(reply).reply('x'),
      (c) => c.presentation.slides.getItemAt(0).comments.getItem(reply).delete(),
      (c) => {
        c.presentation.slides.getItemAt(0).comments.getItem(first).replies.getItem(first).content =
          'x';
      },
    ];
    for (const [i, call] of cases.entries())
      expect([i, ...(await codeOf(runtime, call))]).toEqual([
        i,
        'ItemNotFound',
        'edit.comment-not-found',
      ]);
    const all = await readAll(runtime);
    expect(all.map((t) => [t.content, t.resolved, t.replies])).toEqual([
      ['First', false, ['Reply to first']],
      ['Second', false, []],
    ]);
    runtime.dispose();
  });

  test('a legacy reply proxy does not reach its thread', async () => {
    const runtime = await server(buildSlideStructureDeck());
    expect(
      await codeOf(runtime, (context) => {
        context.presentation.slides
          .getItemAt(0)
          .comments.getItem('cm-0-1')
          .replies.getItem('cm-0-1').content = 'hijack';
      })
    ).toEqual(['ItemNotFound', 'edit.comment-not-found']);
    expect((await readAll(runtime))[0]!.content).toBe('Check the chart numbers.');
    runtime.dispose();
  });
});

describe('Slide.comments: loads and edge cases', () => {
  test('replies load through a navigation path, on a thread and on the collection', async () => {
    const runtime = await server();
    await runtime.run(async (context) => {
      const comment = context.presentation.slides.getItemAt(0).comments.add('Thread');
      await context.sync();
      comment.reply('One');
      comment.reply('Two');
      await context.sync();
    });
    const read = await runtime.run(async (context) => {
      const comments = context.presentation.slides.getItemAt(0).comments;
      comments.load('items/content,items/replies/items/content');
      const first = comments.getItemAt(0);
      first.load('replies/items/authorName');
      await context.sync();
      return [
        comments.items.map((c) => [c.content, c.replies.items.map((r) => r.content)]),
        first.replies.items.map((r) => r.authorName),
      ];
    });
    expect(read).toEqual([[['Thread', ['One', 'Two']]], ['Review Agent', 'Review Agent']]);
    runtime.dispose();
  });

  test('the replies of a thread that does not exist refuse with ItemNotFound', async () => {
    const runtime = await server();
    const calls: ((context: RequestContext) => void)[] = [
      (c) => void c.presentation.slides.getItemAt(0).comments.getItem('nope').replies.getCount(),
      (c) =>
        void c.presentation.slides
          .getItemAt(0)
          .comments.getItem('nope')
          .replies.load('items/content'),
      (c) =>
        void c.presentation.slides
          .getItemAt(0)
          .comments.getItem('nope')
          .replies.getItemAt(0)
          .load('content'),
    ];
    for (const [i, call] of calls.entries()) {
      const error = await runtime.run((context) => refusedAtSync(context, () => call(context)));
      expect([i, isPptxEditorError(error) && error.code]).toEqual([i, 'ItemNotFound']);
    }
    runtime.dispose();
  });

  test('a reply of a thread found in the same sync writes to that thread', async () => {
    const runtime = await server();
    const replyId = await runtime.run(async (context) => {
      const comment = context.presentation.slides.getItemAt(0).comments.add('Thread');
      await context.sync();
      const reply = comment.reply('Remove me');
      await context.sync();
      return reply.id;
    });
    await runtime.run(async (context) => {
      context.presentation.slides
        .getItemAt(0)
        .comments.getItemAt(0)
        .replies.getItem(replyId)
        .delete();
      await context.sync();
    });
    expect((await readAll(runtime))[0]!.replies).toEqual([]);
    runtime.dispose();
  });

  test('a blank author name or long initials refuse, not fall back or cut', async () => {
    const runtime = await server();
    for (const options of [
      { authorName: '   ' },
      { authorName: '' },
      { authorName: '\u0001\u0002' },
      { authorInitials: 'ABCDEFGHIJ' },
    ]) {
      const [code] = await codeOf(runtime, (context) =>
        context.presentation.slides.getItemAt(0).comments.add('Hi', options)
      );
      expect([options, code]).toEqual([options, 'InvalidArgument']);
    }
    expect(await readAll(runtime)).toEqual([]);
    runtime.dispose();
  });

  test('a write of the current resolved state changes nothing', async () => {
    const live = createPptxEditor({ fixedMeasurer: true, substitutes: false, author: 'Editor' });
    expect((await live.load(plainDeck())).ok).toBe(true);
    const runtime = BrowserEditor.createBrowser(live);
    await runtime.run(async (context) => {
      context.presentation.slides.getItemAt(0).comments.add('Open');
      await context.sync();
    });
    const revision = live.session!.store.revision;
    await runtime.run(async (context) => {
      const comments = context.presentation.slides.getItemAt(0).comments.load('items/resolved');
      await context.sync();
      comments.items[0]!.resolved = false;
      await context.sync();
    });
    expect(live.session!.store.revision).toBe(revision);
    // Undo removes the add: the write of the current state left no step of its own.
    expect(live.exec({ type: 'undo' }).ok).toBe(true);
    expect(await readAll(runtime)).toEqual([]);
  });
});

describe('Slide.comments: legacy comments', () => {
  test('read a legacy thread and reply; resolve refuses with NotSupported', async () => {
    const runtime = await server(buildSlideStructureDeck());
    const [thread] = await readAll(runtime);
    expect(thread).toEqual({
      id: 'cm-0-1',
      authorName: 'Fixture Author',
      authorInitials: 'FA',
      content: 'Check the chart numbers.',
      resolved: false,
      anchor: { type: 'point', left: 1.25, top: 1.25 },
      replies: ['Checked.'],
    });
    const date = await runtime.run(async (context) => {
      const c = context.presentation.slides.getItemAt(0).comments.getItem('cm-0-1');
      c.load('creationDate');
      await context.sync();
      return c.creationDate?.toISOString();
    });
    // The file stores the date with no zone: it counts as UTC.
    expect(date).toBe('2026-01-01T09:00:00.000Z');
    const added = await runtime.run(async (context) => {
      const comments = context.presentation.slides.getItemAt(0).comments;
      comments.getItem('cm-0-1').reply('Fixed.');
      const fresh = comments.add('New legacy thread', {
        shape: comments.context.presentation.slides.getItemAt(0).shapes.getItemAt(0),
      });
      await context.sync();
      return fresh.id;
    });
    expect(added).toMatch(/^cm-\d+-\d+$/);
    const all = await readAll(runtime);
    expect(all[0]!.replies).toEqual(['Checked.', 'Fixed.']);
    expect(all[1]).toMatchObject({
      id: added,
      content: 'New legacy thread',
      authorName: 'Review Agent',
    });
    expect(
      await codeOf(runtime, (context) => {
        context.presentation.slides.getItemAt(0).comments.getItem('cm-0-1').resolved = true;
      })
    ).toEqual(['NotSupported', 'edit.comment-no-status']);
    expect(
      await codeOf(runtime, (context) =>
        context.presentation.slides
          .getItemAt(0)
          .comments.add('x', { shape: '2', textRange: { start: 0, length: 2 } })
      )
    ).toEqual(['NotSupported', 'edit.comment-no-text-range']);
    runtime.dispose();
  });
});

describe('Slide.comments: browser runtime', () => {
  test('reads and writes as the server runtime does, as one undo step, signed by the editor author', async () => {
    const live = createPptxEditor({
      fixedMeasurer: true,
      substitutes: false,
      author: 'Editor Author',
    });
    expect((await live.load(plainDeck())).ok).toBe(true);
    const runtime = BrowserEditor.createBrowser(live);
    await runtime.run(async (context) => {
      const comments = context.presentation.slides.getItemAt(0).comments;
      comments.add('One');
      comments.add('Two', { shape: '2' });
      await context.sync();
    });
    const all = await readAll(runtime);
    expect(all.map((c) => [c.content, c.authorName])).toEqual([
      ['One', 'Editor Author'],
      ['Two', 'Editor Author'],
    ]);
    const pkg = live.session!.store.package;
    const part = presentationInfo(pkg).slides[0]!.partName;
    expect(readCommentThreads(pkg, part).map((t) => t.text)).toEqual(['One', 'Two']);
    expect(live.exec({ type: 'undo' }).ok).toBe(true);
    expect(readCommentThreads(live.session!.store.package, part)).toEqual([]);
    expect(live.snapshot().canUndo).toBe(false);
  });

  test('the runtime author option wins over the editor author', async () => {
    const live = createPptxEditor({
      fixedMeasurer: true,
      substitutes: false,
      author: 'Editor Author',
    });
    expect((await live.load(plainDeck())).ok).toBe(true);
    const runtime = BrowserEditor.createBrowser(live, { author: 'Agent' });
    await runtime.run(async (context) => {
      context.presentation.slides.getItemAt(0).comments.add('Hi');
      await context.sync();
    });
    expect((await readAll(runtime))[0]!.authorName).toBe('Agent');
  });

  test('an add sends the lookups and one addComment operation', async () => {
    // The collaboration test (`packages/pro/src/__tests__/yjs-automation-comments.test.ts`)
    // sends these operations to two linked editors; this test keeps them equal.
    const live = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    expect((await live.load(plainDeck())).ok).toBe(true);
    const inner = createBrowserAutomationHost(live, { author: 'Agent' });
    const batches: AutomationOperation[][] = [];
    const host: AutomationHost = {
      ...inner,
      execute: (request) => {
        batches.push([...request.operations]);
        return inner.execute(request);
      },
    };
    await createRuntime({ host, save: false }).run(async (context) => {
      context.presentation.slides.getItemAt(0).comments.add('Hi');
      await context.sync();
    });
    expect(batches).toEqual([
      [
        { op: 'getPresentation' },
        { op: 'getSlide', key: { index: 0 } },
        { op: 'addComment', slide: { result: 1 }, text: 'Hi' },
      ],
    ]);
  });
});

describe('Slide.comments: browser runtime author', () => {
  test('a browser runtime with no author refuses as a server runtime does', async () => {
    const live = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    expect((await live.load(plainDeck())).ok).toBe(true);
    const runtime = BrowserEditor.createBrowser(live);
    expect(
      await codeOf(runtime, (context) =>
        context.presentation.slides.getItemAt(0).comments.add('Hi')
      )
    ).toEqual(['NotSupported', 'edit.no-comment-author']);
  });

  test('createBrowser refuses an author option that is not a name', async () => {
    const live = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    expect((await live.load(plainDeck())).ok).toBe(true);
    for (const author of ['', '\u0007\u0007', '\uD800', 'x'.repeat(256), 7]) {
      let error: unknown;
      try {
        BrowserEditor.createBrowser(live, { author: author as string });
      } catch (e) {
        error = e;
      }
      expect(isPptxEditorError(error) && error.code).toBe('InvalidArgument');
    }
    let error: unknown;
    try {
      BrowserEditor.createBrowser(live, null as unknown as { author?: string });
    } catch (e) {
      error = e;
    }
    expect(isPptxEditorError(error) && error.code).toBe('InvalidArgument');
  });
});

describe('Slide.comments: documentation', () => {
  test('the example of the comments page runs as written', async () => {
    // docs/site/content/editor-api/comments.mdx, "Review the comments of a slide".
    async function answerOpenComments(runtime: PptxEditorRuntime) {
      return runtime.run(async (context) => {
        const comments = context.presentation.slides.getItemAt(0).comments;
        comments.load('items/authorName,items/content,items/resolved');
        await context.sync();
        const open: Comment[] = comments.items.filter((comment) => !comment.resolved);
        for (const comment of open) {
          comment.reply('Fixed in the next draft.', { authorName: 'Review Agent' });
          comment.resolved = true;
        }
        await context.sync();
        return open.map((comment) => `${comment.authorName}: ${comment.content}`);
      });
    }
    const runtime = await server();
    await runtime.run(async (context) => {
      context.presentation.slides.getItemAt(0).comments.add('Shorten the title.', {
        authorName: 'Grace Hopper',
      });
      await context.sync();
    });
    expect(await answerOpenComments(runtime)).toEqual(['Grace Hopper: Shorten the title.']);
    expect((await readAll(runtime))[0]).toMatchObject({
      resolved: true,
      replies: ['Fixed in the next draft.'],
    });
    runtime.dispose();
  });
});
