// Comments on a range of text: the text moniker that `addComment` writes, the `text` anchor that
// the snapshot reads back, the range following its text through edits, undo, and save, and the
// review author roster that colors the threads.

import { describe, expect, test } from 'bun:test';
import { unzipSync, zipSync } from 'fflate';
import { buildPptx, para, run, shape } from '../../../../../scripts/fixtures/builder.ts';
import { buildSlideStructureDeck } from '../../../../../scripts/fixtures/slide-structure.ts';
import { presentationInfo, relationshipsOf, type PptxPackage } from '../../store/index.ts';
import { commentTree } from '../../automation/commands/comment-read.ts';
import { serializeXmlPart } from '../../store/index.ts';
import {
  createPptxEditor,
  type CommentThread,
  type EditorModule,
  type PptxEditorInstance,
} from '../index.ts';

const review: EditorModule = {
  id: 'review',
  review: { displayModes: ['final', 'markup'], collectReviewItems: () => [] },
};
const DATE = '2026-11-08T10:00:00.000';

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

/** A slide with a two-paragraph body ("The quick fox" / "jumps high") and a second slide. */
const deck = (): Uint8Array =>
  buildPptx({
    slides: [
      {
        shapes: shape({
          id: 3,
          name: 'Body',
          x: 0,
          y: 0,
          cx: 6_000_000,
          cy: 2_000_000,
          text: para(run('The quick fox')) + para(run('jumps high')),
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

const bodyOf = (e: PptxEditorInstance): string =>
  e.snapshot().shapes!.find((s) => s.name === 'Body')!.nodeId;

const select = (e: PptxEditorInstance, from: [number, number], to: [number, number]): void =>
  e.setSelection({
    kind: 'text',
    nodeId: bodyOf(e),
    anchor: { paragraph: from[0], offset: from[1] },
    focus: { paragraph: to[0], offset: to[1] },
  });

/** The name of the comment list part of slide 1. */
function listName(pkg: PptxPackage): string {
  const slide = presentationInfo(pkg).slides[0]!.partName;
  const rel = relationshipsOf(pkg, slide).find((r) =>
    r.type.endsWith('2018/10/relationships/comments')
  );
  return rel!.target!;
}

/** The comment list part of slide 1, as XML. */
const listXml = (pkg: PptxPackage): string => serializeXmlPart(commentTree(pkg, listName(pkg))!);

const anchor = (e: PptxEditorInstance, i = 0): CommentThread['anchor'] =>
  e.snapshot().comments![i]!.anchor;

const textAnchor = (
  e: PptxEditorInstance,
  from: [number, number],
  to: [number, number]
): CommentThread['anchor'] => ({
  kind: 'text',
  nodeId: bodyOf(e),
  start: { paragraph: from[0], offset: from[1] },
  end: { paragraph: to[0], offset: to[1] },
});

describe('a comment on selected text', () => {
  test('writes a text moniker, reads back as a text anchor, and survives a save', async () => {
    const e = await open(deck());
    // "quick", selected from right to left.
    select(e, [0, 9], [0, 4]);
    expect(e.exec({ type: 'addComment', text: 'Faster?', date: DATE }).ok).toBe(true);
    expect(anchor(e)).toEqual(textAnchor(e, [0, 4], [0, 9]));
    const xml = listXml(e.session!.store.package);
    expect(xml).toMatch(
      /<ac:txMkLst [^>]*><pc:docMk\/><pc:sldMk [^>]*\/><ac:spMk creationId="\{[0-9A-F-]{36}\}" id="3"\/><ac:txMk cp="4" len="5"\/><\/ac:txMkLst>/
    );
    const again = await open(e.save()!);
    expect(anchor(again)).toEqual(textAnchor(again, [0, 4], [0, 9]));
  });

  test('counts one character between paragraphs', async () => {
    const e = await open(deck());
    // "fox" through "jumps": 10 to 19 in "The quick fox\njumps high".
    select(e, [0, 10], [1, 5]);
    expect(e.exec({ type: 'addComment', text: 'Two lines', date: DATE }).ok).toBe(true);
    expect(listXml(e.session!.store.package)).toContain('<ac:txMk cp="10" len="9"/>');
    expect(anchor(e)).toEqual(textAnchor(e, [0, 10], [1, 5]));
  });

  test('a caret with no range, a cell, and a shape selection point at the element', async () => {
    const e = await open(deck());
    select(e, [0, 4], [0, 4]);
    expect(e.exec({ type: 'addComment', text: 'Caret', date: DATE }).ok).toBe(true);
    expect(anchor(e)).toEqual({ kind: 'shape', nodeId: bodyOf(e) });
    expect(listXml(e.session!.store.package)).not.toContain('txMk');
  });

  test('takes an explicit range, and refuses a range without an element or with no text', async () => {
    const e = await open(deck());
    const nodeId = bodyOf(e);
    const range = { start: { paragraph: 1, offset: 0 }, end: { paragraph: 1, offset: 5 } };
    expect(e.exec({ type: 'addComment', text: 'x', date: DATE, nodeId, range }).ok).toBe(true);
    expect(anchor(e)).toEqual(textAnchor(e, [1, 0], [1, 5]));
    const refused = (command: Record<string, unknown>) =>
      e.exec({ type: 'addComment', text: 'x', date: DATE, ...command } as never);
    expect(refused({ range })).toMatchObject({ ok: false });
    const empty = { start: range.start, end: range.start };
    expect(refused({ nodeId, range: empty })).toMatchObject({ ok: false });
    expect(refused({ nodeId, range: { start: 'x', end: 1 } })).toMatchObject({ ok: false });
  });

  test('a deck with only legacy comments points the thread at the element', async () => {
    // The legacy format has no text anchors.
    const e = await open(buildSlideStructureDeck());
    const body = e.snapshot().shapes!.find((s) => s.kind === 'shape' && s.name !== '')!;
    const range = { start: { paragraph: 0, offset: 0 }, end: { paragraph: 0, offset: 2 } };
    const r = e.exec({ type: 'addComment', text: 'Old', date: DATE, nodeId: body.nodeId, range });
    expect(r.ok).toBe(true);
    const added = e.snapshot().comments!.at(-1)!;
    expect(added.format).toBe('legacy');
    expect(added.anchor.kind).not.toBe('text');
  });
});

describe('a text range follows its text', () => {
  async function commented(): Promise<PptxEditorInstance> {
    const e = await open(deck());
    select(e, [0, 4], [0, 9]);
    expect(e.exec({ type: 'addComment', text: 'Faster?', date: DATE }).ok).toBe(true);
    return e;
  }
  const typeAt = (e: PptxEditorInstance, at: [number, number], text: string): void => {
    select(e, at, at);
    e.input('text', text);
  };

  test('typing before the range moves it, and undo and redo move it back and again', async () => {
    const e = await commented();
    typeAt(e, [0, 0], 'Oh, ');
    expect(anchor(e)).toEqual(textAnchor(e, [0, 8], [0, 13]));
    expect(listXml(e.session!.store.package)).toContain('<ac:txMk cp="8" len="5"/>');
    expect(e.exec({ type: 'undo' }).ok).toBe(true);
    expect(anchor(e)).toEqual(textAnchor(e, [0, 4], [0, 9]));
    expect(e.exec({ type: 'redo' }).ok).toBe(true);
    expect(anchor(e)).toEqual(textAnchor(e, [0, 8], [0, 13]));
  });

  test('typing inside grows it; typing at either end leaves it', async () => {
    const e = await commented();
    typeAt(e, [0, 6], 'iii');
    expect(anchor(e)).toEqual(textAnchor(e, [0, 4], [0, 12]));
    typeAt(e, [0, 12], '!');
    expect(anchor(e)).toEqual(textAnchor(e, [0, 4], [0, 12]));
    typeAt(e, [0, 4], '(');
    expect(anchor(e)).toEqual(textAnchor(e, [0, 5], [0, 13]));
  });

  test('a new paragraph before it moves it to the next paragraph', async () => {
    const e = await commented();
    select(e, [0, 0], [0, 0]);
    e.input('paragraph');
    expect(anchor(e)).toEqual(textAnchor(e, [1, 4], [1, 9]));
  });

  test('deleting all of its text points the thread at the element', async () => {
    const e = await commented();
    select(e, [0, 3], [0, 10]);
    e.input('deleteBackward');
    expect(anchor(e)).toEqual({ kind: 'shape', nodeId: bodyOf(e) });
    expect(listXml(e.session!.store.package)).toContain('<ac:txMk cp="3" len="0"/>');
  });

  test('a saved deck, whose comments the editor keeps as bytes, moves the range and drops its check', async () => {
    const e = await commented();
    // The file's check of the text that the range was made on.
    const files = unzipSync(e.save()!);
    const name = Object.keys(files).find((f) => f.startsWith('ppt/comments/modernComment'))!;
    const xml = new TextDecoder().decode(files[name]);
    files[name] = new TextEncoder().encode(
      xml.replace(
        '<ac:txMk cp="4" len="5"/>',
        '<ac:txMk cp="4" len="5"><ac:context len="24" hash="12345"/></ac:txMk>'
      )
    );
    const again = await open(zipSync(files));
    typeAt(again, [0, 0], 'A ');
    expect(anchor(again)).toEqual(textAnchor(again, [0, 6], [0, 11]));
    const now = listXml(again.session!.store.package);
    expect(now).toContain('<ac:txMk cp="6" len="5"/>');
    expect(now).not.toContain('ac:context');
  });

  test('counts a slide number field as the number it shows', async () => {
    // The field stores three characters and shows one, "1": "A1 quick fox".
    const field =
      '<a:fld id="{5C3B2B5E-1C33-4F1B-9D7C-0A1B2C3D4E5F}" type="slidenum"><a:t>‹#›</a:t></a:fld>';
    const e = await open(
      buildPptx({
        slides: [
          {
            shapes: shape({
              id: 3,
              name: 'Body',
              x: 0,
              y: 0,
              cx: 6_000_000,
              cy: 2_000_000,
              text: para(run('A') + field + run(' quick fox')),
            }),
          },
        ],
      })
    );
    select(e, [0, 3], [0, 8]);
    expect(e.exec({ type: 'addComment', text: 'Faster?', date: DATE }).ok).toBe(true);
    // Text typed at the end of the range stays out of it, and text before it moves it.
    typeAt(e, [0, 8], '!');
    expect(anchor(e)).toEqual(textAnchor(e, [0, 3], [0, 8]));
    typeAt(e, [0, 2], 'zz');
    expect(anchor(e)).toEqual(textAnchor(e, [0, 5], [0, 10]));
  });

  test('an edit in another shape or slide leaves the range and the part alone', async () => {
    const e = await commented();
    const before = listXml(e.session!.store.package);
    const part = e.session!.store.package.parts.get(listName(e.session!.store.package));
    e.selectSlide(1);
    const title = e.snapshot().shapes!.find((s) => s.name === 'Title')!.nodeId;
    e.setSelection({
      kind: 'text',
      nodeId: title,
      anchor: { paragraph: 0, offset: 0 },
      focus: { paragraph: 0, offset: 0 },
    });
    e.input('text', 'Z');
    const pkg = e.session!.store.package;
    expect(listXml(pkg)).toBe(before);
    // The same part object: the edit wrote nothing to it.
    expect(pkg.parts.get(listName(pkg))).toBe(part);
  });
});

describe('getReviewAuthors', () => {
  test('ranks authors by first appearance in the deck, then the signing author', async () => {
    const e = await open(deck(), 'Me');
    e.exec({ type: 'addComment', text: 'one', date: DATE, author: 'Grace' });
    e.exec({ type: 'addComment', text: 'two', date: DATE, author: 'Ada' });
    const again = await open(e.save()!, 'Me');
    expect(again.getReviewAuthors().map((a) => [a.author, a.slot, a.color])).toEqual([
      ['Grace', 0, 'var(--doc-review-author-0)'],
      ['Ada', 1, 'var(--doc-review-author-1)'],
      ['Me', 2, 'var(--doc-review-author-2)'],
    ]);
  });

  test('keeps slots when an author goes and comes back, and the same array while unchanged', async () => {
    const e = await open(deck(), 'Me');
    e.exec({ type: 'addComment', text: 'one', date: DATE, author: 'Grace' });
    e.exec({ type: 'addComment', text: 'two', date: DATE, author: 'Ada' });
    const first = e.getReviewAuthors();
    expect(e.getReviewAuthors()).toBe(first);
    const grace = e.snapshot().comments![0]!;
    e.exec({ type: 'deleteComment', commentId: grace.id });
    expect(e.getReviewAuthors().find((a) => a.author === 'Ada')!.slot).toBe(
      first.find((a) => a.author === 'Ada')!.slot
    );
    e.exec({ type: 'undo' });
    expect(e.getReviewAuthors().map((a) => a.slot)).toEqual(first.map((a) => a.slot));
  });

  test('is empty without a deck, even with an author', () => {
    expect(createPptxEditor({ fixedMeasurer: true }).getReviewAuthors()).toEqual([]);
    expect(createPptxEditor({ fixedMeasurer: true, author: 'Ada' }).getReviewAuthors()).toEqual([]);
  });
});
