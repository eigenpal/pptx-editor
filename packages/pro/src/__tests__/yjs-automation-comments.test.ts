// The automation comment operations, which `Slide.comments` of the object model sends, through
// two editors on two linked Y.Docs (`AGENTS.md` "Collaboration verification"). The writes go to
// the threaded comment part and the authors part, or to the legacy `p:cmLst` (ECMA-376 Part 1,
// 19.4.4) and `p:cmAuthorLst` (19.4.3); both replicas must converge.

import { describe, expect, test } from 'bun:test';
import * as Y from 'yjs';
import type { AutomationHost, AutomationOperation } from '@pptx-editor.dev/core/automation';
import {
  createBrowserAutomationHost,
  createPptxEditor,
  readCommentThreads,
  type PptxEditorInstance,
} from '@pptx-editor.dev/core/editor';
import {
  canonicalPackageFingerprint,
  presentationInfo,
  readPptx,
} from '@pptx-editor.dev/core/store';
import { buildPptx, para, run, shape } from '../../../../scripts/fixtures/builder.ts';
import { buildSlideStructureDeck } from '../../../../scripts/fixtures/slide-structure.ts';
import { collaborationModule, createYjsCollaborationSession } from '../index.ts';

/** Two docs and a sync link that can pause, to make edits concurrent. */
function link(a: Y.Doc, b: Y.Doc) {
  const queue: [Y.Doc, Uint8Array][] = [];
  let paused = false;
  const send = (to: Y.Doc) => (update: Uint8Array, origin: unknown) => {
    if (origin === 'link') return;
    if (paused) queue.push([to, update]);
    else Y.applyUpdate(to, update, 'link');
  };
  a.on('update', send(b));
  b.on('update', send(a));
  return {
    pause: () => (paused = true),
    flush: () => {
      paused = false;
      for (const [to, update] of queue.splice(0)) Y.applyUpdate(to, update, 'link');
    },
  };
}

/** Two slides; the first has a text box. */
const deck = (): Uint8Array =>
  buildPptx({
    slides: [
      {
        layout: 'blank',
        shapes: shape({
          id: 2,
          name: 'Box',
          x: 0,
          y: 0,
          cx: 914_400,
          cy: 914_400,
          text: para(run('Hello')),
        }),
      },
      { layout: 'blank', shapes: '' },
    ],
  });

async function replica(doc: Y.Doc, actorId: string, bytes: Uint8Array) {
  const session = createYjsCollaborationSession({
    doc,
    documentId: 'deck',
    identity: { actorId, name: `Peer ${actorId.toUpperCase()}` },
  });
  const editor = createPptxEditor({
    fixedMeasurer: true,
    substitutes: false,
    modules: [collaborationModule({ session })],
  });
  expect((await editor.load(bytes)).ok).toBe(true);
  return editor;
}

/**
 * Run comment operations on slide `index` as `Slide.comments` sends them: the lookups, then the
 * writes. `packages/editor-api/src/__tests__/comments.test.ts` checks that it sends these.
 */
function send(host: AutomationHost, index: number, ops: AutomationOperation[]): string[] {
  const operations: AutomationOperation[] = [
    { op: 'getPresentation' },
    { op: 'getSlide', key: { index } },
    ...ops,
  ];
  const response = host.execute({ operations });
  if (!response.ok) throw new Error(JSON.stringify(response.results));
  return response.results.slice(2).map((r) => {
    const v = r.status === 'ok' ? r.value : undefined;
    return v?.kind === 'text' ? v.text : '';
  });
}

const SLIDE = { result: 1 } as const;
const fingerprint = (e: PptxEditorInstance): string =>
  canonicalPackageFingerprint(e.session!.store.package);

function threadsOf(e: PptxEditorInstance, index = 0) {
  const pkg = e.session!.store.package;
  const part = presentationInfo(pkg).slides[index]!.partName;
  return readCommentThreads(pkg, part).map((t) => ({
    text: t.text,
    author: t.author,
    resolved: t.resolved,
    replies: t.replies.map((r) => r.text),
  }));
}

function savedFingerprint(e: PptxEditorInstance): string {
  const r = readPptx(e.save()!);
  if (!r.ok) throw new Error('reopen');
  return canonicalPackageFingerprint(r.pkg);
}

async function pair(bytes = deck()) {
  const docA = new Y.Doc();
  const docB = new Y.Doc();
  const sync = link(docA, docB);
  return { sync, a: await replica(docA, 'a', bytes), b: await replica(docB, 'b', bytes) };
}

describe('Slide.comments on two peers', () => {
  test('a comment added on one peer reaches the other, signed by its peer', async () => {
    const { a, b } = await pair();
    send(createBrowserAutomationHost(b), 0, [{ op: 'addComment', slide: SLIDE, text: 'From B' }]);
    expect(threadsOf(a)).toEqual([
      { text: 'From B', author: 'Peer B', resolved: false, replies: [] },
    ]);
    expect(fingerprint(a)).toBe(fingerprint(b));
    expect(savedFingerprint(a)).toBe(savedFingerprint(b));
  });

  test('concurrent new threads on one slide keep both, while a peer edits the slide', async () => {
    const { sync, a, b } = await pair();
    sync.pause();
    send(createBrowserAutomationHost(a), 0, [{ op: 'addComment', slide: SLIDE, text: 'From A' }]);
    send(createBrowserAutomationHost(b), 0, [
      { op: 'addComment', slide: SLIDE, text: 'From B' },
      { op: 'getShape', container: SLIDE, key: { id: '2' } },
      { op: 'setText', shape: { result: 3 }, text: 'Changed' },
    ]);
    sync.flush();
    expect(fingerprint(a)).toBe(fingerprint(b));
    expect(
      threadsOf(a)
        .map((t) => t.text)
        .sort()
    ).toEqual(['From A', 'From B']);
    expect(savedFingerprint(a)).toBe(savedFingerprint(b));
  });

  test('a concurrent reply and resolve of one thread keep both', async () => {
    const { sync, a, b } = await pair();
    const [id] = send(createBrowserAutomationHost(a), 0, [
      { op: 'addComment', slide: SLIDE, text: 'Thread' },
    ]);
    sync.pause();
    send(createBrowserAutomationHost(a), 0, [
      { op: 'replyToComment', slide: SLIDE, comment: id!, text: 'Reply A' },
    ]);
    send(createBrowserAutomationHost(b), 0, [
      { op: 'setCommentResolved', slide: SLIDE, comment: id!, resolved: true },
    ]);
    sync.flush();
    expect(fingerprint(a)).toBe(fingerprint(b));
    expect(threadsOf(a)).toEqual([
      { text: 'Thread', author: 'Peer A', resolved: true, replies: ['Reply A'] },
    ]);
  });

  test('a concurrent delete and reply of one thread converge', async () => {
    const { sync, a, b } = await pair();
    const [id] = send(createBrowserAutomationHost(a), 0, [
      { op: 'addComment', slide: SLIDE, text: 'Thread' },
    ]);
    sync.pause();
    send(createBrowserAutomationHost(a), 0, [{ op: 'deleteComment', slide: SLIDE, comment: id! }]);
    send(createBrowserAutomationHost(b), 0, [
      { op: 'replyToComment', slide: SLIDE, comment: id!, text: 'Late reply' },
    ]);
    sync.flush();
    expect(fingerprint(a)).toBe(fingerprint(b));
    expect(threadsOf(a)).toEqual(threadsOf(b));
    expect(savedFingerprint(a)).toBe(savedFingerprint(b));
  });

  test('concurrent legacy threads on one slide keep both', async () => {
    const { sync, a, b } = await pair(buildSlideStructureDeck());
    sync.pause();
    send(createBrowserAutomationHost(a), 0, [{ op: 'addComment', slide: SLIDE, text: 'From A' }]);
    send(createBrowserAutomationHost(b), 0, [{ op: 'addComment', slide: SLIDE, text: 'From B' }]);
    sync.flush();
    expect(fingerprint(a)).toBe(fingerprint(b));
    expect(threadsOf(a).map((t) => t.text)).toEqual(
      expect.arrayContaining(['Check the chart numbers.', 'From A', 'From B'])
    );
    expect(savedFingerprint(a)).toBe(savedFingerprint(b));
  });

  test('undo of an add on one peer removes it on both', async () => {
    const { a, b } = await pair();
    send(createBrowserAutomationHost(b), 0, [{ op: 'addComment', slide: SLIDE, text: 'Undo me' }]);
    expect(b.exec({ type: 'undo' }).ok).toBe(true);
    expect(threadsOf(a)).toEqual([]);
    expect(fingerprint(a)).toBe(fingerprint(b));
  });
});
