// The comment commands and the comment threads of the snapshot, in both comment formats.

import { describe, expect, test } from 'bun:test';
import { buildPptx, para, run, shape } from '../../../../../scripts/fixtures/builder.ts';
import { buildSlideStructureDeck } from '../../../../../scripts/fixtures/slide-structure.ts';
import {
  contentTypeOf,
  presentationInfo,
  readPptx,
  relationshipsOf,
  serializeXmlPart,
  xmlPartOf,
  type PptxPackage,
} from '../../store/index.ts';
import { commentTree } from '../../automation/commands/comment-read.ts';
import { createPptxEditor, type EditorModule, type PptxEditorInstance } from '../index.ts';

const review: EditorModule = {
  id: 'review',
  review: { displayModes: ['final', 'markup'], collectReviewItems: () => [] },
};

async function open(bytes: Uint8Array, author = 'Ada Lovelace'): Promise<PptxEditorInstance> {
  const e = createPptxEditor({
    fixedMeasurer: true,
    substitutes: false,
    modules: [review],
    author,
  });
  expect((await e.load(bytes)).ok).toBe(true);
  return e;
}

const plainDeck = (): Uint8Array =>
  buildPptx({
    slides: [
      {
        shapes:
          shape({
            id: 2,
            name: 'Title',
            x: 0,
            y: 0,
            cx: 4_000_000,
            cy: 800_000,
            text: para(run('Plan')),
          }) +
          shape({
            id: 3,
            name: 'Body',
            x: 0,
            y: 1_000_000,
            cx: 4_000_000,
            cy: 800_000,
            text: para(run('Body')),
          }),
      },
      {
        shapes: shape({
          id: 2,
          name: 'Title',
          x: 0,
          y: 0,
          cx: 4_000_000,
          cy: 800_000,
          text: para(run('Two')),
        }),
      },
    ],
  });

const xmlOf = (pkg: PptxPackage, part: string): string => {
  const tree = commentTree(pkg, part);
  if (!tree) throw new Error(`no part ${part}`);
  return serializeXmlPart(tree);
};

const relOf = (pkg: PptxPackage, source: string, suffix: string) =>
  relationshipsOf(pkg, source).find((r) => r.type.endsWith(suffix));

const reopen = (e: PptxEditorInstance): PptxPackage => {
  const r = readPptx(e.save()!);
  if (!r.ok) throw new Error('the saved deck does not read');
  return r.pkg;
};

const DATE = '2026-11-07T10:00:00.000';

describe('threaded comments', () => {
  test('add, reply, edit, resolve, and delete write the threaded parts', async () => {
    const e = await open(plainDeck());
    expect(e.snapshot().comments).toEqual([]);
    expect(e.exec({ type: 'addComment', text: 'Check <this> & that', date: DATE }).ok).toBe(true);
    let [thread] = e.snapshot().comments!;
    expect(thread).toMatchObject({
      format: 'modern',
      author: 'Ada Lovelace',
      initials: 'AL',
      date: DATE,
      text: 'Check <this> & that',
      resolved: false,
      anchor: { kind: 'slide' },
      replies: [],
    });
    expect(
      e.exec({
        type: 'replyToComment',
        commentId: thread!.id,
        text: 'Done.\nTwo lines',
        author: 'Grace Hopper',
        date: DATE,
      }).ok
    ).toBe(true);
    [thread] = e.snapshot().comments!;
    expect(thread!.replies).toMatchObject([
      { author: 'Grace Hopper', initials: 'GH', text: 'Done.\nTwo lines' },
    ]);
    expect(e.exec({ type: 'editComment', commentId: thread!.id, text: 'Edited' }).ok).toBe(true);
    expect(e.exec({ type: 'resolveComment', commentId: thread!.id }).ok).toBe(true);
    [thread] = e.snapshot().comments!;
    expect(thread).toMatchObject({ text: 'Edited', resolved: true });

    const pkg = reopen(e);
    const slide = presentationInfo(pkg).slides[0]!.partName;
    const list = relOf(pkg, slide, '2018/10/relationships/comments')!.target!;
    expect(list).toMatch(/^\/ppt\/comments\/modernComment_100_[0-9A-F]+\.xml$/);
    expect(contentTypeOf(pkg, list)).toBe('application/vnd.ms-powerpoint.comments+xml');
    const authors = relOf(pkg, '/ppt/presentation.xml', '2018/10/relationships/authors')!.target!;
    expect(authors).toBe('/ppt/authors.xml');
    expect(contentTypeOf(pkg, authors)).toBe('application/vnd.ms-powerpoint.authors+xml');
    expect(xmlOf(pkg, authors)).toContain(
      'name="Grace Hopper" providerId="None" userId="Grace Hopper"'
    );
    const cm = xmlOf(pkg, list);
    expect(cm).toContain('status="resolved"');
    expect(cm).toContain('<pc:sldMk cId="');
    expect(cm).toContain('<a:t>Edited</a:t>');
    expect(cm).toContain(
      '<a:p><a:r><a:t>Done.</a:t></a:r></a:p><a:p><a:r><a:t>Two lines</a:t></a:r></a:p>'
    );
    const slideXml = serializeXmlPart(xmlPartOf(pkg, slide)!);
    expect(slideXml).toContain('{6950BFC3-D8DA-4A85-94F7-54DA5524770B}');
    expect(slideXml).toContain('p14:creationId');

    // The reopened deck reads the same threads.
    const again = await open(e.save()!);
    expect(again.snapshot().comments).toEqual(e.snapshot().comments);

    expect(e.exec({ type: 'reopenComment', commentId: thread!.id }).ok).toBe(true);
    expect(e.snapshot().comments![0]!.resolved).toBe(false);
    expect(e.exec({ type: 'deleteComment', commentId: thread!.replies[0]!.id }).ok).toBe(true);
    expect(e.snapshot().comments![0]!.replies).toEqual([]);
    expect(e.exec({ type: 'deleteComment', commentId: thread!.id }).ok).toBe(true);
    expect(e.snapshot().comments).toEqual([]);
    // The empty list goes with its relationship, its type, and the slide's reference.
    const after = reopen(e);
    expect(relOf(after, slide, '2018/10/relationships/comments')).toBeUndefined();
    expect(after.parts.has(list)).toBe(false);
    expect(serializeXmlPart(xmlPartOf(after, slide)!)).not.toContain('commentRel');
  });

  test('a comment on a selected element names it, and undo removes the comment', async () => {
    const e = await open(plainDeck());
    const body = e.snapshot().shapes!.find((s) => s.name === 'Body')!.nodeId;
    e.setSelection({ kind: 'shapes', nodeIds: [body] });
    expect(e.exec({ type: 'addComment', text: 'On the body', date: DATE }).ok).toBe(true);
    expect(e.snapshot().comments![0]!.anchor).toEqual({ kind: 'shape', nodeId: body });
    const pkg = e.session!.store.package;
    const slide = presentationInfo(pkg).slides[0]!.partName;
    expect(serializeXmlPart(xmlPartOf(pkg, slide)!)).toContain('a16:creationId');
    const list = relOf(pkg, slide, '2018/10/relationships/comments')!.target!;
    expect(xmlOf(pkg, list)).toMatch(/<ac:spMk creationId="\{[0-9A-F-]{36}\}" id="3"\/>/);
    // The anchor reads back after a save.
    const again = await open(e.save()!);
    const reread = again.snapshot().comments![0]!.anchor;
    expect(reread.kind).toBe('shape');
    expect(e.exec({ type: 'undo' }).ok).toBe(true);
    expect(e.snapshot().comments).toEqual([]);
  });

  test('refuses empty text, no author, and unknown ids', async () => {
    const e = createPptxEditor({ fixedMeasurer: true, substitutes: false, modules: [review] });
    await e.load(plainDeck());
    expect(e.exec({ type: 'addComment', text: 'Hi' })).toMatchObject({
      ok: false,
      reason: 'disabled.noCommentAuthor',
    });
    expect(e.exec({ type: 'addComment', text: '  ', author: 'A' })).toMatchObject({
      ok: false,
      reason: 'errors.commentText',
    });
    expect(e.exec({ type: 'addComment', text: 'x', author: 'A', date: 'yesterday' })).toMatchObject(
      {
        ok: false,
        reason: 'errors.commentDate',
      }
    );
    expect(
      e.exec({ type: 'deleteComment', commentId: '/ppt/slides/slide1.xml#{0}' })
    ).toMatchObject({
      ok: false,
      reason: 'errors.noComment',
    });
    e.setAuthor('Lin');
    expect(e.exec({ type: 'addComment', text: 'Hi' }).ok).toBe(true);
    expect(e.snapshot().comments![0]!.author).toBe('Lin');
  });

  test('without a review module the snapshot has no threads, and the commands work', async () => {
    const e = createPptxEditor({ fixedMeasurer: true, substitutes: false, author: 'A' });
    await e.load(plainDeck());
    expect(e.snapshot().comments).toBeNull();
    expect(e.exec({ type: 'addComment', text: 'Hi' }).ok).toBe(true);
  });
});

describe('legacy comments', () => {
  test('a deck with legacy comments keeps their format', async () => {
    const e = await open(buildSlideStructureDeck());
    const threads = e.snapshot().comments!;
    expect(threads).toHaveLength(1);
    expect(threads[0]).toMatchObject({
      format: 'legacy',
      author: 'Fixture Author',
      text: 'Check the chart numbers.',
      replies: [{ text: 'Checked.' }],
    });
    const id = threads[0]!.id;
    expect(e.exec({ type: 'resolveComment', commentId: id })).toMatchObject({
      ok: false,
      reason: 'disabled.commentNoStatus',
    });
    expect(e.exec({ type: 'replyToComment', commentId: id, text: 'Thanks', date: DATE }).ok).toBe(
      true
    );
    expect(e.exec({ type: 'editComment', commentId: id, text: 'Check the numbers' }).ok).toBe(true);
    e.selectSlide(1);
    expect(e.exec({ type: 'addComment', text: 'New one', date: DATE }).ok).toBe(true);
    const pkg = reopen(e);
    const authors = xmlOf(pkg, '/ppt/commentAuthors.xml');
    expect(authors).toContain('name="Ada Lovelace"');
    expect(authors).toMatch(/id="1" initials="AL" lastIdx="2"/);
    const first = xmlOf(pkg, '/ppt/comments/comment1.xml');
    expect(first).toContain('<p:text>Check the numbers</p:text>');
    expect(first).toContain('<p15:parentCm authorId="0" idx="1"/>');
    expect(first).toContain('authorId="1" dt="2026-11-07T10:00:00.000" idx="1"');
    const second = relOf(pkg, presentationInfo(pkg).slides[1]!.partName, 'relationships/comments')!;
    expect(xmlOf(pkg, second.target!)).toContain('idx="2"');
    // Deleting the thread takes its replies.
    e.selectSlide(0);
    expect(e.exec({ type: 'deleteComment', commentId: id }).ok).toBe(true);
    expect(e.snapshot().comments).toEqual([]);
  });
});
