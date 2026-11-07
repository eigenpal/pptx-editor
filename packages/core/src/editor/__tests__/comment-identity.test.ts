// Comment identities: the token that names a comment in the snapshot while the deck is open. A
// file can repeat a comment's GUID, author, and date (twins); the editor keeps their identities
// where it can tell them apart, and gives them new ones where it cannot.

import { describe, expect, test } from 'bun:test';
import { buildPptx, shape } from '../../../../../scripts/fixtures/builder.ts';
import { buildSlideStructureDeck } from '../../../../../scripts/fixtures/slide-structure.ts';
import {
  parseXmlPart,
  readPptx,
  serializeXmlPart,
  writePptx,
  type PackagePart,
} from '../../store/index.ts';
import { createPptxEditor, type EditorModule, type PptxEditorInstance } from '../index.ts';

const review: EditorModule = {
  id: 'review',
  review: { displayModes: ['final', 'markup'], collectReviewItems: () => [] },
};
const DATE = '2026-11-07T10:00:00.000';

async function open(bytes: Uint8Array): Promise<PptxEditorInstance> {
  const e = createPptxEditor({ fixedMeasurer: true, substitutes: false, modules: [review] });
  expect((await e.load(bytes)).ok).toBe(true);
  return e;
}

/** `bytes` with each comment list part that `pattern` names passed through `edit`. */
function patch(bytes: Uint8Array, pattern: RegExp, edit: (xml: string) => string): Uint8Array {
  const r = readPptx(bytes);
  if (!r.ok) throw new Error('the deck does not read');
  const parts = r.pkg.parts as Map<string, PackagePart>;
  for (const [name, part] of [...parts]) {
    if (!pattern.test(name)) continue;
    const xml =
      part.kind === 'xml' ? serializeXmlPart(part.xml) : new TextDecoder().decode(part.bytes);
    const parsed = parseXmlPart(name, edit(xml));
    if (!parsed.ok) throw new Error('the patched list does not parse');
    parts.set(name, { kind: 'xml', partName: name, xml: parsed.part });
  }
  return writePptx(r.pkg);
}

/** A deck whose one threaded comment "A" (with a reply "R" when asked) `edit` repeats. */
async function threads(edit: (thread: string) => string, reply = false): Promise<Uint8Array> {
  const e = createPptxEditor({
    fixedMeasurer: true,
    substitutes: false,
    modules: [review],
    author: 'Ada Lovelace',
  });
  const deck = buildPptx({
    slides: [{ shapes: shape({ id: 2, name: 'Title', x: 0, y: 0, cx: 914_400, cy: 914_400 }) }],
  });
  expect((await e.load(deck)).ok).toBe(true);
  expect(e.exec({ type: 'addComment', text: 'A', date: DATE }).ok).toBe(true);
  if (reply) {
    const [t] = e.snapshot().comments!;
    expect(e.exec({ type: 'replyToComment', commentId: t!.id, text: 'R', date: DATE }).ok).toBe(
      true
    );
  }
  const saved = e.save()!;
  e.destroy();
  return patch(saved, /\/comments\/modernComment/, (xml) => {
    const thread = /<p188:cm [^>]*>.*?<\/p188:cm>/s.exec(xml)![0];
    return xml.replace(thread, edit(thread));
  });
}

const as = (thread: string, text: string): string => thread.replace('>A<', `>${text}<`);
const identities = (e: PptxEditorInstance): string[] =>
  e.snapshot().comments!.map((t) => t.identity);

describe('comment identity', () => {
  test('a delete of one twin keeps the identities of the others', async () => {
    const e = await open(await threads((t) => as(t, 'A') + as(t, 'B') + as(t, 'C')));
    const [a, b, c] = e.snapshot().comments!;
    expect(e.exec({ type: 'deleteComment', commentId: a!.id }).ok).toBe(true);
    expect(identities(e)).toEqual([b!.identity, c!.identity]);
  });

  test('a delete of one of two twins keeps the identity of the other', async () => {
    const e = await open(await threads((t) => as(t, 'A') + as(t, 'B')));
    const [a, b] = e.snapshot().comments!;
    e.selectComment(b!.id);
    expect(e.exec({ type: 'deleteComment', commentId: a!.id }).ok).toBe(true);
    expect(identities(e)).toEqual([b!.identity]);
    expect(e.snapshot().activeComment).toBe(e.snapshot().comments![0]!.id);
  });

  test('an edit of a picked twin keeps its identity and the pick', async () => {
    const e = await open(await threads((t) => as(t, 'A') + as(t, 'B')));
    const [a, b] = e.snapshot().comments!;
    e.selectComment(a!.id);
    expect(e.exec({ type: 'editComment', commentId: a!.id, text: 'A2' }).ok).toBe(true);
    expect(identities(e)).toEqual([a!.identity, b!.identity]);
    expect(e.snapshot().activeComment).toBe(a!.id);
  });

  test('when twins with one text shrink, no identity or pick moves to another twin', async () => {
    const e = await open(await threads((t) => t + t + t));
    const before = identities(e);
    const [first, second] = e.snapshot().comments!;
    e.selectComment(second!.id);
    expect(e.exec({ type: 'deleteComment', commentId: first!.id }).ok).toBe(true);
    const after = identities(e);
    expect(after).toHaveLength(2);
    for (const id of after) expect(before).not.toContain(id);
    expect(new Set(after).size).toBe(2);
    expect(e.snapshot().activeComment).toBeNull();
  });

  test('the comment that takes the id of a deleted pick can be picked', async () => {
    const e = await open(await threads((t) => as(t, 'A') + as(t, 'B') + as(t, 'C')));
    const [a] = e.snapshot().comments!;
    e.selectComment(a!.id);
    expect(e.exec({ type: 'deleteComment', commentId: a!.id }).ok).toBe(true);
    expect(e.snapshot().activeComment).toBeNull();
    const [b] = e.snapshot().comments!;
    expect(b!.id).toBe(a!.id);
    e.selectComment(b!.id);
    expect(e.snapshot().activeComment).toBe(b!.id);
  });

  test('replies of repeated threads and a GUID that holds ~ get their own identities', async () => {
    const e = await open(
      await threads((t) => {
        const guid = / id="([^"]+)"/.exec(t)![1]!;
        return t + as(t, 'B') + as(t, 'C').replace(guid, `${guid}~1`);
      }, true)
    );
    const all = e.snapshot().comments!.flatMap((t) => [t, ...t.replies]);
    expect(all).toHaveLength(6);
    expect(new Set(all.map((c) => c.identity)).size).toBe(6);
    expect(all[4]!.identity).toContain('%7E1~');
  });

  test('legacy comments that repeat an author and index get their own ids and identities', async () => {
    const deck = patch(buildSlideStructureDeck(), /\/comments\/comment\d+\.xml$/, (xml) => {
      const cm = /<p:cm [^>]*>(?:(?!<\/p:cm>).)*<p:text>Check[^<]*<\/p:text>.*?<\/p:cm>/s.exec(
        xml
      )![0];
      return xml.replace(cm, cm + cm.replace(/<p:text>[^<]*/, '<p:text>Second'));
    });
    const e = await open(deck);
    const [first, second] = e.snapshot().comments!;
    expect(first!.id).not.toBe(second!.id);
    expect(first!.identity).not.toBe(second!.identity);
    e.selectComment(second!.id);
    expect(e.snapshot().comments!.find((t) => t.id === e.snapshot().activeComment)!.text).toBe(
      'Second'
    );
    expect(e.exec({ type: 'editComment', commentId: second!.id, text: 'Edited' }).ok).toBe(true);
    expect(e.snapshot().comments!.map((t) => t.text)).toEqual([
      'Check the chart numbers.',
      'Edited',
    ]);
    // The plain key names no comment while the author and index repeat.
    const plain = second!.id.replace(/-\d+$/, '');
    expect(e.exec({ type: 'editComment', commentId: plain, text: 'x' }).ok).toBe(false);
  });

  test('a delete and an undo in one read give the twins new identities', async () => {
    const e = await open(await threads((t) => as(t, 'A') + as(t, 'B') + as(t, 'C')));
    const c = e.snapshot().comments![2]!;
    expect(e.exec({ type: 'deleteComment', commentId: c.id }).ok).toBe(true);
    const s1 = e.snapshot().comments!;
    expect(e.exec({ type: 'undo' }).ok).toBe(true);
    expect(e.exec({ type: 'deleteComment', commentId: s1[0]!.id }).ok).toBe(true);
    const s2 = e.snapshot().comments!;
    expect(s2.map((t) => t.text)).toEqual(['B', 'C']);
    for (const t of s2) expect(s1.map((x) => x.identity)).not.toContain(t.identity);
  });

  test('an undo of a history group closes the fields of the twins, and never moves them', async () => {
    const e = await open(await threads((t) => as(t, 'x') + as(t, 'y')));
    const [a, b] = e.snapshot().comments!;
    const g = e.beginHistoryGroup();
    expect(e.exec({ type: 'deleteComment', commentId: a!.id }, { historyGroup: g }).ok).toBe(true);
    const b1 = e.snapshot().comments![0]!;
    expect(b1.identity).toBe(b!.identity);
    expect(
      e.exec({ type: 'editComment', commentId: b1.id, text: 'x' }, { historyGroup: g }).ok
    ).toBe(true);
    g.end();
    const [picked] = e.snapshot().comments!;
    e.selectComment(picked!.id);
    expect(e.snapshot().activeComment).toBe(picked!.id);
    expect(e.exec({ type: 'undo' }).ok).toBe(true);
    const s2 = e.snapshot();
    expect(s2.comments!.map((t) => t.text)).toEqual(['x', 'y']);
    expect(s2.comments!.some((t) => t.identity === picked!.identity)).toBe(false);
    expect(s2.activeComment).toBeNull();
  });

  test('a slide that comes back keeps the identities it showed', async () => {
    const e0 = createPptxEditor({
      fixedMeasurer: true,
      substitutes: false,
      modules: [review],
      author: 'Ada Lovelace',
    });
    const sh = shape({ id: 2, name: 'Title', x: 0, y: 0, cx: 914_400, cy: 914_400 });
    expect((await e0.load(buildPptx({ slides: [{ shapes: sh }, { shapes: sh }] }))).ok).toBe(true);
    expect(e0.exec({ type: 'addComment', text: 'A', date: DATE }).ok).toBe(true);
    const bytes = patch(e0.save()!, /\/comments\/modernComment/, (xml) => {
      const t = /<p188:cm [^>]*>.*?<\/p188:cm>/s.exec(xml)![0];
      return xml.replace(t, as(t, 'x') + as(t, 'y'));
    });
    e0.destroy();
    const e = await open(bytes);
    const [a, b] = e.snapshot().comments!;
    expect(e.exec({ type: 'editComment', commentId: a!.id, text: 'y' }).ok).toBe(true);
    e.snapshot();
    expect(e.exec({ type: 'editComment', commentId: b!.id, text: 'x' }).ok).toBe(true);
    const before = identities(e);
    expect(before).toEqual([a!.identity, b!.identity]);
    e.goToSlide(1);
    e.snapshot();
    e.goToSlide(0);
    expect(identities(e)).toEqual(before);
  });

  test('expect: a twin whose text changed refuses the write, and can() agrees', async () => {
    const e = await open(await threads((t) => as(t, 'A') + as(t, 'B')));
    const [a] = e.snapshot().comments!;
    const seen = { author: a!.author, date: a!.date!, text: a!.text };
    expect(e.exec({ type: 'editComment', commentId: a!.id, text: 'Peer' }).ok).toBe(true);
    const mine = { type: 'editComment', commentId: a!.id, text: 'Mine', expect: seen } as const;
    expect(e.can(mine)).toMatchObject({ ok: false, reason: 'errors.commentChanged' });
    expect(e.exec(mine)).toMatchObject({ ok: false, reason: 'errors.commentChanged' });
    for (const type of ['deleteComment', 'resolveComment'] as const)
      expect(e.exec({ type, commentId: a!.id, expect: seen }).ok).toBe(false);
    expect(e.exec({ type: 'replyToComment', commentId: a!.id, text: 'R', expect: seen }).ok).toBe(
      false
    );
    expect(e.snapshot().comments!.map((t) => t.text)).toEqual(['Peer', 'B']);
    const fresh = { ...seen, text: 'Peer' };
    expect(e.exec({ ...mine, expect: fresh }).ok).toBe(true);
  });

  test('expect: a peer edit of a comment without a twin never refuses your edit', async () => {
    const e = await open(await threads((t) => t));
    const [a] = e.snapshot().comments!;
    const seen = { author: a!.author, date: a!.date!, text: a!.text };
    expect(e.exec({ type: 'editComment', commentId: a!.id, text: 'Peer' }).ok).toBe(true);
    expect(e.exec({ type: 'editComment', commentId: a!.id, text: 'Mine', expect: seen }).ok).toBe(
      true
    );
    expect(identities(e)).toEqual([a!.identity]);
    // Another author or date is another comment.
    for (const other of [
      { ...seen, author: 'Grace' },
      { author: seen.author, text: seen.text },
    ])
      expect(
        e.exec({ type: 'editComment', commentId: a!.id, text: 'x', expect: other })
      ).toMatchObject({ ok: false, reason: 'errors.commentChanged' });
  });
});
