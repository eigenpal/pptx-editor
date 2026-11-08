/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/

// Comments that two peers write at the same time: new threads, replies, edits, status, and
// deletes in both comment formats. Each case pauses the link, lets both peers edit, flushes,
// and checks that both replicas hold one deck, that its saves agree, and what the threads are.

import { describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import * as Y from 'yjs';
import {
  createPptxEditor,
  readCommentThreads,
  type EditorCommand,
  type PptxEditorInstance,
} from '@pptx-editor.dev/core/editor';
import {
  presentationInfo,
  readPptx,
  relationshipsOf,
  serializeXmlPart,
  xmlPartOf,
  type PptxPackage,
} from '@pptx-editor.dev/core/store';
import { collaborationModule, createYjsCollaborationSession } from '../index.ts';
import { coverageDeck, differingParts } from './convergence-harness.ts';

const FIXTURES = join(import.meta.dir, '../../../../fixtures');
const DATE = '2026-11-07T10:00:00.000';

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
    pause: () => void (paused = true),
    flush: () => {
      paused = false;
      for (const [to, update] of queue.splice(0)) Y.applyUpdate(to, update, 'link');
    },
  };
}

async function peers(bytes: Uint8Array) {
  const docs = [new Y.Doc(), new Y.Doc()] as const;
  const sync = link(docs[0], docs[1]);
  const open = async (doc: Y.Doc, actorId: string) => {
    const session = createYjsCollaborationSession({
      doc,
      documentId: 'deck',
      identity: { actorId, name: actorId === 'a' ? 'Ada' : 'Ben' },
    });
    const editor = createPptxEditor({
      fixedMeasurer: true,
      substitutes: false,
      modules: [collaborationModule({ session })],
    });
    expect((await editor.load(bytes)).ok).toBe(true);
    return editor;
  };
  const a = await open(docs[0], 'a');
  const b = await open(docs[1], 'b');
  return { a, b, sync };
}

const run = (e: PptxEditorInstance, command: EditorCommand): void =>
  expect(e.exec(command)).toEqual({ ok: true });

const threads = (e: PptxEditorInstance) =>
  readCommentThreads(e.session!.store.package, e.session!.info.slides[0]!.partName);

/** Both replicas hold one deck, and their saves reopen to one package. */
function converged(a: PptxEditorInstance, b: PptxEditorInstance): void {
  expect(differingParts(a.session!.store.package, b.session!.store.package)).toEqual([]);
  const saved = [a, b].map((e) => {
    const r = readPptx(e.save()!);
    if (!r.ok) throw new Error('a saved deck does not reopen');
    return r.pkg;
  });
  expect(differingParts(saved[0]!, saved[1]!)).toEqual([]);
  wellFormed(saved[0]!);
  // The save joins copies, and the reopened deck reads the threads the editor shows.
  const slide = presentationInfo(saved[0]!).slides[0]!.partName;
  expect(readCommentThreads(saved[0]!, slide)).toEqual(threads(a));
}

/**
 * A saved deck holds one author list and one comment list per slide in each format, one reply
 * list per thread, an author for each comment, and slide monikers that name their slide.
 */
function wellFormed(pkg: PptxPackage): void {
  const of = (source: string, type: string) =>
    relationshipsOf(pkg, source).filter((r) => r.type.endsWith(type));
  expect(of('/ppt/presentation.xml', '/commentAuthors').length).toBeLessThanOrEqual(1);
  expect(of('/ppt/presentation.xml', '2018/10/relationships/authors').length).toBeLessThanOrEqual(
    1
  );
  const authors = of('/ppt/presentation.xml', '2018/10/relationships/authors')[0];
  const authorXml = authors ? serializeXmlPart(xmlPartOf(pkg, authors.target!)!) : '';
  for (const s of presentationInfo(pkg).slides) {
    const slideXml = serializeXmlPart(xmlPartOf(pkg, s.partName)!);
    expect(of(s.partName, 'officeDocument/2006/relationships/comments').length).toBeLessThanOrEqual(
      1
    );
    const modern = of(s.partName, '2018/10/relationships/comments');
    expect(modern.length).toBeLessThanOrEqual(1);
    if (!modern[0]) continue;
    expect(slideXml.match(/<p188:commentRel /g)).toHaveLength(1);
    expect(slideXml).toContain(`r:id="${modern[0].id}"`);
    const cId = /<p14:creationId [^>]*val="(\d+)"/.exec(slideXml)![1];
    const list = serializeXmlPart(xmlPartOf(pkg, modern[0].target!)!);
    for (const cm of list.split('<p188:cm ').slice(1)) {
      expect(cm.match(/<p188:replyLst>/g)?.length ?? 0).toBeLessThanOrEqual(1);
      expect(cm).toContain(`<pc:sldMk cId="${cId}" sldId="${s.sldId}"/>`);
    }
    for (const [, id] of list.matchAll(/authorId="(\{[^"]+\})"/g))
      expect(authorXml).toContain(`id="${id}"`);
  }
}

/** The comment relationships of slide 1. */
const commentRels = (e: PptxEditorInstance) =>
  relationshipsOf(e.session!.store.package, e.session!.info.slides[0]!.partName).filter((r) =>
    r.type.endsWith('/comments')
  );

describe('threaded comments in a room', () => {
  test('two peers add threads and replies at once, and both stay', async () => {
    const { a, b, sync } = await peers(coverageDeck());
    try {
      run(a, { type: 'addComment', text: 'First', date: DATE });
      const first = threads(a)[0]!.id;
      sync.pause();
      run(a, { type: 'addComment', text: 'From Ada', date: DATE });
      run(b, { type: 'addComment', text: 'From Ben', date: DATE });
      run(a, { type: 'replyToComment', commentId: first, text: 'Ada replies', date: DATE });
      run(b, { type: 'replyToComment', commentId: first, text: 'Ben replies', date: DATE });
      sync.flush();
      converged(a, b);
      const texts = threads(a).map((t) => t.text);
      expect(texts[0]).toBe('First');
      expect(texts.slice(1).sort()).toEqual(['From Ada', 'From Ben']);
      expect(
        threads(a)[0]!
          .replies.map((r) => `${r.author}: ${r.text}`)
          .sort()
      ).toEqual(['Ada: Ada replies', 'Ben: Ben replies']);
      expect(commentRels(a)).toHaveLength(1);
    } finally {
      a.destroy();
      b.destroy();
    }
  });

  test('an edit, a resolve, and a reply to one thread at once all apply', async () => {
    const { a, b, sync } = await peers(coverageDeck());
    try {
      run(a, { type: 'addComment', text: 'Draft', date: DATE });
      const id = threads(a)[0]!.id;
      sync.pause();
      run(a, { type: 'editComment', commentId: id, text: 'Final' });
      run(b, { type: 'resolveComment', commentId: id });
      run(b, { type: 'replyToComment', commentId: id, text: 'Ok', date: DATE });
      sync.flush();
      converged(a, b);
      expect(threads(b)[0]).toMatchObject({
        text: 'Final',
        resolved: true,
        replies: [{ text: 'Ok' }],
      });
    } finally {
      a.destroy();
      b.destroy();
    }
  });

  test('a delete and a reply to the deleted thread at once converge', async () => {
    const { a, b, sync } = await peers(coverageDeck());
    try {
      run(a, { type: 'addComment', text: 'One', date: DATE });
      run(a, { type: 'addComment', text: 'Two', date: DATE });
      const id = threads(a)[0]!.id;
      sync.pause();
      run(a, { type: 'deleteComment', commentId: id });
      run(b, { type: 'replyToComment', commentId: id, text: 'Late reply', date: DATE });
      sync.flush();
      converged(a, b);
      expect(threads(a).map((t) => t.text)).toEqual(['Two']);
    } finally {
      a.destroy();
      b.destroy();
    }
  });

  test('a delete of the only thread and a new comment at once keep the new comment', async () => {
    const { a, b, sync } = await peers(coverageDeck());
    try {
      run(a, { type: 'addComment', text: 'Only', date: DATE });
      const id = threads(a)[0]!.id;
      sync.pause();
      run(a, { type: 'deleteComment', commentId: id });
      run(b, { type: 'addComment', text: 'From Ben', date: DATE });
      sync.flush();
      converged(a, b);
      expect(threads(a).map((t) => t.text)).toEqual(['From Ben']);
      expect(threads(b).map((t) => t.text)).toEqual(['From Ben']);
      // A save after the last delete holds no comment list.
      run(b, { type: 'deleteComment', commentId: threads(b)[0]!.id });
      expect(threads(a)).toEqual([]);
      const saved = readPptx(a.save()!);
      if (!saved.ok) throw new Error('the saved deck does not reopen');
      const slide = presentationInfo(saved.pkg).slides[0]!.partName;
      expect(relationshipsOf(saved.pkg, slide).filter((r) => r.type.endsWith('/comments'))).toEqual(
        []
      );
      expect([...saved.pkg.parts.keys()].filter((n) => /comment/i.test(n))).toEqual([]);
      expect(serializeXmlPart(xmlPartOf(saved.pkg, slide)!)).not.toContain('commentRel');
    } finally {
      a.destroy();
      b.destroy();
    }
  });

  test('two peers who add the first comment of a slide at once converge', async () => {
    const { a, b, sync } = await peers(coverageDeck());
    try {
      sync.pause();
      run(a, { type: 'addComment', text: 'From Ada', date: DATE });
      run(b, { type: 'addComment', text: 'From Ben', date: DATE });
      sync.flush();
      converged(a, b);
      expect(
        threads(a)
          .map((t) => t.text)
          .sort()
      ).toEqual(['From Ada', 'From Ben']);
      expect(
        threads(a)
          .map((t) => t.author)
          .sort()
      ).toEqual(['Ada', 'Ben']);
    } finally {
      a.destroy();
      b.destroy();
    }
  });

  test('two peers who add the first comment of a legacy deck slide at once converge', async () => {
    const bytes = new Uint8Array(readFileSync(join(FIXTURES, 'slide-structure.pptx')));
    const { a, b, sync } = await peers(bytes);
    try {
      a.goToSlide(1);
      b.goToSlide(1);
      sync.pause();
      run(a, { type: 'addComment', text: 'From Ada', date: DATE });
      run(b, { type: 'addComment', text: 'From Ben', date: DATE });
      sync.flush();
      converged(a, b);
      const part = a.session!.info.slides[1]!.partName;
      const saved = readPptx(a.save()!);
      if (!saved.ok) throw new Error('the saved deck does not reopen');
      expect(
        readCommentThreads(saved.pkg, part)
          .map((t) => `${t.author}: ${t.text}`)
          .sort()
      ).toEqual(['Ada: From Ada', 'Ben: From Ben']);
    } finally {
      a.destroy();
      b.destroy();
    }
  });
});

describe('legacy comments in a room', () => {
  test('a delete of the only thread and a new comment at once keep the new comment', async () => {
    const bytes = new Uint8Array(readFileSync(join(FIXTURES, 'slide-structure.pptx')));
    const { a, b, sync } = await peers(bytes);
    try {
      const id = threads(a)[0]!.id;
      sync.pause();
      run(a, { type: 'deleteComment', commentId: id });
      run(b, { type: 'addComment', text: 'From Ben', date: DATE });
      sync.flush();
      converged(a, b);
      expect(threads(a).map((t) => `${t.format}: ${t.text}`)).toEqual(['legacy: From Ben']);
    } finally {
      a.destroy();
      b.destroy();
    }
  });

  test('two new authors who comment at once get their own ids and indexes', async () => {
    const bytes = new Uint8Array(readFileSync(join(FIXTURES, 'slide-structure.pptx')));
    const { a, b, sync } = await peers(bytes);
    try {
      const id = threads(a)[0]!.id;
      sync.pause();
      run(a, { type: 'replyToComment', commentId: id, text: 'Ada replies', date: DATE });
      run(b, { type: 'replyToComment', commentId: id, text: 'Ben replies', date: DATE });
      run(b, { type: 'addComment', text: 'Ben adds', date: DATE });
      sync.flush();
      converged(a, b);
      const [thread, added] = threads(a);
      expect(thread!.replies.map((r) => `${r.author}: ${r.text}`).sort()).toEqual([
        'Ada: Ada replies',
        'Ben: Ben replies',
        'Fixture Author: Checked.',
      ]);
      expect(added).toMatchObject({ author: 'Ben', text: 'Ben adds', format: 'legacy' });
      const ids = [...thread!.replies, added!].map((c) => c.id);
      expect(new Set(ids).size).toBe(ids.length);
    } finally {
      a.destroy();
      b.destroy();
    }
  });
});
