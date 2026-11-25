// The comment operations of the automation host, in both comment formats. The legacy format is
// ECMA-376 Part 1, 19.4 (`p:cm`, `CT_Comment`, 19.4.1, in a `p:cmLst`, 19.4.4); the threaded
// format is the `p188` comment part that a slide names in its `p:extLst` (19.3.1.20).

import { describe, expect, test } from 'bun:test';
import { buildPptx, para, run, shape } from '../../../../../scripts/fixtures/builder.ts';
import { buildSlideStructureDeck } from '../../../../../scripts/fixtures/slide-structure.ts';
import { readPptx } from '../../store/index.ts';
import {
  createServerAutomationHost,
  type AutomationBatchResponse,
  type AutomationCommentThreadRead,
  type AutomationHandle,
  type AutomationHost,
  type AutomationOperation,
  type AutomationValue,
  type ServerAutomationOptions,
} from '../index.ts';

const IN = 914_400;
const SLIDE = { result: 1 } as const;

function deck(): Uint8Array {
  return buildPptx({
    slides: [
      {
        layout: 'blank',
        shapes:
          shape({
            id: 2,
            name: 'Box',
            x: IN,
            y: IN,
            cx: 2 * IN,
            cy: IN,
            text: para(run('Hello')) + para(run('World')),
          }) +
          '<p:grpSp><p:nvGrpSpPr><p:cNvPr id="4" name="Group"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
          '<p:grpSpPr/>' +
          shape({ id: 5, name: 'Inner', x: 0, y: 0, cx: IN, cy: IN, text: para(run('In')) }) +
          '</p:grpSp>',
      },
      { layout: 'blank', shapes: shape({ id: 2, name: 'Other', x: 0, y: 0, cx: IN, cy: IN }) },
    ],
  });
}

function open(bytes = deck(), options: ServerAutomationOptions = { author: 'Ada Lovelace' }) {
  const opened = createServerAutomationHost(bytes, options);
  if (!opened.ok) throw new Error(JSON.stringify(opened.error));
  return opened.host;
}

/** A batch on slide `index`: the presentation, the slide, then `ops` (which name it as SLIDE). */
function onSlide(
  host: AutomationHost,
  ops: AutomationOperation[],
  index = 0
): AutomationBatchResponse {
  return host.execute({
    operations: [{ op: 'getPresentation' }, { op: 'getSlide', key: { index } }, ...ops],
  });
}

function ok(host: AutomationHost, ops: AutomationOperation[], index = 0): AutomationValue[] {
  const response = onSlide(host, ops, index);
  if (!response.ok) throw new Error(JSON.stringify(response.results));
  return response.results
    .slice(2)
    .map((r) => (r.status === 'ok' ? r.value : { kind: 'null' as const }));
}

function errorOf(host: AutomationHost, ops: AutomationOperation[], index = 0) {
  const response = onSlide(host, ops, index);
  expect(response.ok).toBe(false);
  return response.results.find((r) => r.status === 'error')!;
}

function threads(host: AutomationHost, index = 0): readonly AutomationCommentThreadRead[] {
  const [value] = ok(host, [{ op: 'readComments', slide: SLIDE }], index);
  if (value?.kind !== 'comments') throw new Error(JSON.stringify(value));
  return value.comments;
}

function keyOf(value: AutomationValue | undefined): string {
  if (value?.kind !== 'text') throw new Error(JSON.stringify(value));
  return value.text;
}

const shapeOp = (id: string): AutomationOperation => ({
  op: 'getShape',
  container: SLIDE,
  key: { id },
});

function reopened(host: AutomationHost): AutomationHost {
  const saved = host.save();
  if (!saved.ok) throw new Error('save');
  expect(readPptx(saved.bytes).ok).toBe(true);
  return open(saved.bytes);
}

describe('automation comments: threaded format', () => {
  test('a slide with no comments reads an empty list', () => {
    expect(threads(open())).toEqual([]);
  });

  test('add on the slide, a shape, and a text range; read their anchors back', () => {
    const host = open();
    const [slideKey, , shapeKey, , textKey] = ok(host, [
      { op: 'addComment', slide: SLIDE, text: 'On the slide' },
      shapeOp('2'),
      { op: 'addComment', slide: SLIDE, text: 'On the box', shape: { result: 3 } },
      shapeOp('2'),
      {
        op: 'addComment',
        slide: SLIDE,
        text: 'On World',
        shape: { result: 5 },
        span: { start: 6, length: 5 },
        author: 'Grace Hopper',
        initials: 'GMH',
      },
    ]);
    const all = threads(host);
    expect(all.map((t) => t.id)).toEqual([keyOf(slideKey), keyOf(shapeKey), keyOf(textKey)]);
    expect(all[0]).toMatchObject({
      format: 'modern',
      author: 'Ada Lovelace',
      initials: 'AL',
      text: 'On the slide',
      resolved: false,
      anchor: { kind: 'slide' },
      replies: [],
    });
    expect(all[0]!.date).toMatch(/^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}/);
    expect(all[1]!.anchor).toEqual({ kind: 'shape', shapeId: '2' });
    expect(all[2]).toMatchObject({
      author: 'Grace Hopper',
      initials: 'GMH',
      anchor: { kind: 'text', shapeId: '2', start: 6, length: 5 },
    });
    // The id key is the comment's GUID, braced, as the file stores it.
    expect(keyOf(slideKey)).toMatch(/^\{[0-9A-F-]{36}\}$/);
    // The anchors and the author survive save and reopen.
    expect(threads(reopened(host))).toEqual(all);
  });

  test('reply, edit, resolve, reopen, and delete', () => {
    const host = open();
    const id = keyOf(ok(host, [{ op: 'addComment', slide: SLIDE, text: 'Check this' }])[0]);
    const reply = keyOf(
      ok(host, [{ op: 'replyToComment', slide: SLIDE, comment: id, text: 'Done.\nTwo lines' }])[0]
    );
    ok(host, [
      { op: 'setCommentText', slide: SLIDE, comment: id, text: 'Check this, please' },
      { op: 'setCommentResolved', slide: SLIDE, comment: id, resolved: true },
    ]);
    let [thread] = threads(host);
    expect(thread).toMatchObject({ text: 'Check this, please', resolved: true });
    expect(thread!.replies).toMatchObject([{ id: reply, text: 'Done.\nTwo lines' }]);
    ok(host, [{ op: 'setCommentResolved', slide: SLIDE, comment: id, resolved: false }]);
    expect(threads(host)[0]!.resolved).toBe(false);
    ok(host, [{ op: 'deleteComment', slide: SLIDE, comment: reply }]);
    [thread] = threads(host);
    expect(thread!.replies).toEqual([]);
    ok(host, [{ op: 'deleteComment', slide: SLIDE, comment: id }]);
    expect(threads(host)).toEqual([]);
    expect(threads(reopened(host))).toEqual([]);
  });

  test('a read later in the batch sees the writes before it', () => {
    const host = open();
    const [key, read] = ok(host, [
      { op: 'addComment', slide: SLIDE, text: 'New' },
      { op: 'readComments', slide: SLIDE },
    ]);
    expect(read).toMatchObject({ kind: 'comments', comments: [{ id: keyOf(key), text: 'New' }] });
  });

  test('one batch is one transaction: a refused write commits nothing', () => {
    const host = open();
    const before = host.revision();
    const failure = errorOf(host, [
      { op: 'addComment', slide: SLIDE, text: 'Kept?' },
      { op: 'addComment', slide: SLIDE, text: '   ' },
    ]);
    expect(failure).toMatchObject({
      error: { code: 'invalid-argument', detail: 'errors.commentText' },
    });
    expect(host.revision()).toBe(before);
    expect(threads(host)).toEqual([]);
  });
});

describe('automation comments: author', () => {
  test('without an author option or a per-call author, an add refuses', () => {
    const host = open(deck(), {});
    expect(errorOf(host, [{ op: 'addComment', slide: SLIDE, text: 'Hi' }])).toMatchObject({
      error: { code: 'unsupported-content', detail: 'disabled.noCommentAuthor' },
    });
    ok(host, [{ op: 'addComment', slide: SLIDE, text: 'Hi', author: '  Lin   Yu ' }]);
    expect(threads(host)[0]).toMatchObject({ author: 'Lin Yu', initials: 'LY' });
  });

  test('an author name over 255 characters, or not a string, refuses', () => {
    const host = open();
    expect(
      errorOf(host, [{ op: 'addComment', slide: SLIDE, text: 'Hi', author: 'x'.repeat(256) }])
    ).toMatchObject({ error: { code: 'invalid-argument', detail: 'errors.commentAuthor' } });
    expect(
      errorOf(host, [
        { op: 'addComment', slide: SLIDE, text: 'Hi', author: 7 } as unknown as AutomationOperation,
      ])
    ).toMatchObject({ error: { code: 'invalid-argument', detail: 'comment-author' } });
  });

  test('text and author from the caller are stored as text, not markup', () => {
    const host = open();
    ok(host, [{ op: 'addComment', slide: SLIDE, text: '<b>&amp;</b>', author: 'A <script> "B"' }]);
    const again = threads(reopened(host));
    expect(again[0]).toMatchObject({ text: '<b>&amp;</b>', author: 'A <script> "B"' });
  });
});

describe('automation comments: refusals', () => {
  test('an unknown comment id refuses with item-not-found', () => {
    const host = open();
    for (const op of [
      {
        op: 'replyToComment',
        slide: SLIDE,
        comment: '{00000000-0000-4000-8000-000000000000}',
        text: 'x',
      },
      { op: 'setCommentText', slide: SLIDE, comment: 'cm-0-9', text: 'x' },
      { op: 'deleteComment', slide: SLIDE, comment: 'nope' },
    ] as AutomationOperation[])
      expect(errorOf(host, [op])).toMatchObject({
        error: { code: 'item-not-found', detail: 'errors.noComment' },
      });
    expect(errorOf(host, [{ op: 'deleteComment', slide: SLIDE, comment: '' }])).toMatchObject({
      error: { code: 'invalid-argument', detail: 'comment' },
    });
  });

  test('a comment of another slide is not found on this one', () => {
    const host = open();
    const id = keyOf(ok(host, [{ op: 'addComment', slide: SLIDE, text: 'First' }])[0]);
    expect(errorOf(host, [{ op: 'deleteComment', slide: SLIDE, comment: id }], 1)).toMatchObject({
      error: { code: 'item-not-found' },
    });
  });

  test('a shape of another slide, a shape in a group, or a bad span refuses', () => {
    const host = open();
    const other = ok(host, [shapeOp('2')], 1)[0] as { handle: AutomationHandle };
    expect(
      errorOf(host, [{ op: 'addComment', slide: SLIDE, text: 'x', shape: other.handle }])
    ).toMatchObject({ error: { code: 'invalid-argument', detail: 'comment-shape' } });
    expect(
      errorOf(host, [
        shapeOp('4'),
        { op: 'getShape', container: { result: 2 }, key: { id: '5' } },
        { op: 'addComment', slide: SLIDE, text: 'x', shape: { result: 3 } },
      ])
    ).toMatchObject({ error: { code: 'invalid-argument', detail: 'comment-shape' } });
    for (const span of [
      { start: 0, length: 0 },
      { start: 10, length: 5 },
      { start: -1, length: 2 },
    ])
      expect(
        errorOf(host, [
          shapeOp('2'),
          { op: 'addComment', slide: SLIDE, text: 'x', shape: { result: 2 }, span },
        ])
      ).toMatchObject({ error: { code: 'invalid-argument', detail: 'comment-span' } });
    expect(
      errorOf(host, [{ op: 'addComment', slide: SLIDE, text: 'x', span: { start: 0, length: 1 } }])
    ).toMatchObject({ error: { code: 'invalid-argument', detail: 'comment-span' } });
    expect(threads(host)).toEqual([]);
  });

  test('a write whose expect no longer matches the comment refuses', () => {
    const host = open();
    const id = keyOf(ok(host, [{ op: 'addComment', slide: SLIDE, text: 'Mine' }])[0]);
    const [thread] = threads(host);
    const expect_ = { author: thread!.author, text: thread!.text, date: thread!.date! };
    ok(host, [
      { op: 'setCommentResolved', slide: SLIDE, comment: id, resolved: true, expect: expect_ },
    ]);
    expect(
      errorOf(host, [
        {
          op: 'deleteComment',
          slide: SLIDE,
          comment: id,
          expect: { ...expect_, author: 'Someone Else' },
        },
      ])
    ).toMatchObject({ error: { code: 'transaction-refused', detail: 'errors.commentChanged' } });
  });
});

describe('automation comments: scope', () => {
  test('a key outside the scope is item-not-found and writes nothing', () => {
    const host = open();
    const [a, b] = ok(host, [
      { op: 'addComment', slide: SLIDE, text: 'A' },
      { op: 'addComment', slide: SLIDE, text: 'B' },
    ]).map(keyOf);
    const reply = keyOf(
      ok(host, [{ op: 'replyToComment', slide: SLIDE, comment: a!, text: 'R' }])[0]
    );
    const before = host.revision();
    for (const op of [
      { op: 'deleteComment', slide: SLIDE, comment: b, scope: { replyOf: a } },
      { op: 'setCommentText', slide: SLIDE, comment: a, scope: { replyOf: a }, text: 'x' },
      { op: 'setCommentResolved', slide: SLIDE, comment: reply, scope: 'thread', resolved: true },
      { op: 'replyToComment', slide: SLIDE, comment: reply, scope: 'thread', text: 'x' },
      { op: 'setCommentText', slide: SLIDE, comment: reply, scope: { replyOf: b }, text: 'x' },
    ] as AutomationOperation[])
      expect(errorOf(host, [op])).toMatchObject({
        error: { code: 'item-not-found', detail: 'errors.noComment' },
      });
    expect(
      errorOf(host, [
        {
          op: 'deleteComment',
          slide: SLIDE,
          comment: a,
          scope: 'reply',
        } as unknown as AutomationOperation,
      ])
    ).toMatchObject({ error: { code: 'invalid-argument', detail: 'comment-scope' } });
    expect(host.revision()).toBe(before);
    // In scope: the reply of its own thread, and the thread itself.
    ok(host, [
      { op: 'setCommentText', slide: SLIDE, comment: reply, scope: { replyOf: a! }, text: 'R2' },
      { op: 'setCommentResolved', slide: SLIDE, comment: a!, scope: 'thread', resolved: true },
    ]);
    expect(threads(host)[0]).toMatchObject({ resolved: true, replies: [{ text: 'R2' }] });
  });
});

describe('automation comments: legacy format', () => {
  test('read a legacy thread with its reply, at its p:pos', () => {
    const host = open(buildSlideStructureDeck());
    const [thread] = threads(host);
    expect(thread).toMatchObject({
      id: 'cm-0-1',
      format: 'legacy',
      author: 'Fixture Author',
      initials: 'FA',
      date: '2026-01-01T09:00:00.000',
      text: 'Check the chart numbers.',
      resolved: false,
      // `p:pos x="10" y="10"` counts in 1/576 inch: 10 / 576 * 72 = 1.25 points.
      anchor: { kind: 'point', left: 1.25, top: 1.25 },
      replies: [{ id: 'cm-0-2', text: 'Checked.', author: 'Fixture Author' }],
    });
  });

  test('a new thread and a reply keep the legacy format; resolve refuses', () => {
    const host = open(buildSlideStructureDeck());
    const [added] = ok(host, [
      shapeOp('2'),
      { op: 'addComment', slide: SLIDE, text: 'New', shape: { result: 2 } },
    ]).slice(1);
    const key = keyOf(added);
    expect(key).toMatch(/^cm-\d+-\d+$/);
    ok(host, [{ op: 'replyToComment', slide: SLIDE, comment: 'cm-0-1', text: 'Again' }]);
    const all = threads(host);
    expect(all.map((t) => t.format)).toEqual(['legacy', 'legacy']);
    expect(all[0]!.replies.map((r) => r.text)).toEqual(['Checked.', 'Again']);
    expect(all[1]).toMatchObject({ id: key, author: 'Ada Lovelace', text: 'New' });
    expect(
      errorOf(host, [{ op: 'setCommentResolved', slide: SLIDE, comment: 'cm-0-1', resolved: true }])
    ).toMatchObject({ error: { code: 'unsupported-content', detail: 'disabled.commentNoStatus' } });
    // A text range needs the threaded format.
    expect(
      errorOf(host, [
        shapeOp('2'),
        {
          op: 'addComment',
          slide: SLIDE,
          text: 'x',
          shape: { result: 2 },
          span: { start: 0, length: 2 },
        },
      ])
    ).toMatchObject({
      error: { code: 'unsupported-content', detail: 'disabled.commentNoTextRange' },
    });
    expect(threads(reopened(host))).toEqual(all);
  });
});
