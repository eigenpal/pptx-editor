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

  test('legacy comments that repeat an author and index get their own identities', async () => {
    const deck = patch(buildSlideStructureDeck(), /\/comments\/comment\d+\.xml$/, (xml) => {
      const cm = /<p:cm [^>]*>(?:(?!<\/p:cm>).)*<p:text>Check[^<]*<\/p:text>.*?<\/p:cm>/s.exec(
        xml
      )![0];
      return xml.replace(cm, cm + cm);
    });
    const e = await open(deck);
    const threads = e.snapshot().comments!;
    expect(threads).toHaveLength(2);
    expect(threads[0]!.id).toBe(threads[1]!.id);
    expect(threads[0]!.identity).not.toBe(threads[1]!.identity);
  });
});
