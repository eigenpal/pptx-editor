/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Concurrent text structure edits (openspec change `collaboration-text-lineage`, "Character
// identity"): a split moves text while a peer edits that text, two Enters cut one paragraph,
// and a delete removes a run a peer types into. Two editors on two linked Y.Docs, the link
// paused, the edits on both, then a flush, in both client-id orders. Each case compares the
// text, because replicas also converge on wrong text.

import { describe, expect, test } from 'bun:test';
import * as Y from 'yjs';
import type { PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import { canonicalPackageFingerprint, readPptx, semanticDigest } from '@pptx-editor.dev/core/store';
import { buildPptx, para, run, shape } from '../../../../scripts/fixtures/builder.ts';
import { F, ROOM_LIMITS } from '../collaboration/yjs/schema.ts';
import { COPIES_KEY, IDENTITY } from '../collaboration/yjs/text-identity.ts';
import {
  IN,
  converged,
  formatOf,
  inlineOf,
  oneRun,
  peers,
  replicas,
  selectIn,
  sessionsOf,
  slide1,
  twoParagraphs,
} from './lineage-support.ts';

const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const ok = (r: { ok: boolean }) => expect(r.ok).toBe(true);
/** Selects `from`..`to` in the first paragraph of the slide's first shape. */
const sel = (e: PptxEditorInstance, from: number, to = from) => selectIn(e, 0, from, to);
const order = (swap: boolean) => `client order ${swap ? 'B, A' : 'A, B'}`;
const textBox = (body: string) =>
  buildPptx({
    slides: [
      { layout: 'blank', shapes: shape({ id: 2, x: IN, y: IN, cx: 4 * IN, cy: IN, text: body }) },
    ],
  });
/** "ABCDEFGH" with "DE" bold: three runs. */
const boldMiddle = () => textBox(para(run('ABC') + run('DE', 'b="1"') + run('FGH')));
/** One run and a paragraph end (`a:endParaRPr`), as most decks write a paragraph. */
const withEnd = (text: string) => textBox(para(run(text) + '<a:endParaRPr lang="en-US"/>'));

/**
 * Both replicas converged (`converged`: status, fingerprint, paragraphs, reopen), and each one's
 * save reopens to the same package, which a second save and reopen keeps.
 */
function savedAlike(...editors: PptxEditorInstance[]): string[] {
  const text = converged(...editors);
  const saved = editors.map((e) => {
    const r = readPptx(e.save()!);
    if (!r.ok) throw new Error('reopen failed');
    return r.pkg;
  });
  for (const pkg of saved) {
    expect(canonicalPackageFingerprint(pkg)).toBe(canonicalPackageFingerprint(saved[0]!));
    expect(semanticDigest(pkg)).toBe(canonicalPackageFingerprint(pkg));
  }
  return text;
}

/** Peer A bolds characters 3 to 5 of "Hello" while peer B runs `edit`; then the link flushes. */
async function boldWhile(swap: boolean, edit: (b: PptxEditorInstance) => void) {
  const pair = await replicas(oneRun('Hello'), swap);
  pair.sync.pause();
  sel(pair.a, 3, 5);
  ok(pair.a.exec({ type: 'toggleBold' }));
  edit(pair.b);
  pair.sync.flush();
  return pair;
}

describe('an edit at a run a peer splits at the same time (#17)', () => {
  for (const swap of [false, true]) {
    test(`text typed after the moved text goes with it, not "Hel!lo" (${order(swap)})`, async () => {
      const { a, b } = await boldWhile(swap, (b) => (sel(b, 5), ok(b.input('text', '!'))));
      expect(savedAlike(a, b)).toEqual(['Hello!']);
      expect(formatOf(a, 'lo!')).toEqual({ b: '1' });
      // The typist's undo takes its text back, and its redo puts it back where it shows.
      ok(b.exec({ type: 'undo' }));
      expect(savedAlike(a, b)).toEqual(['Hello']);
      ok(b.exec({ type: 'redo' }));
      expect(savedAlike(a, b)).toEqual(['Hello!']);
    });

    test(`text typed inside the moved text stays inside it, not "HelXlo" (${order(swap)})`, async () => {
      const { a, b } = await boldWhile(swap, (b) => (sel(b, 4), ok(b.input('text', 'X'))));
      expect(savedAlike(a, b)).toEqual(['HellXo']);
      expect(formatOf(a, 'lXo')).toEqual({ b: '1' });
    });

    test(`text typed at the cut stays before it (${order(swap)})`, async () => {
      const { a, b } = await boldWhile(swap, (b) => (sel(b, 3), ok(b.input('text', 'X'))));
      expect(savedAlike(a, b)).toEqual(['HelXlo']);
      expect(formatOf(a, 'lo')).toEqual({ b: '1' });
    });

    test(`an Enter at the cut shows the moved text once, not "Hello" and "lo" (${order(swap)})`, async () => {
      const { a, b } = await boldWhile(swap, (b) => (sel(b, 3), ok(b.input('paragraph'))));
      expect(savedAlike(a, b)).toEqual(['Hel', 'lo']);
      ok(b.exec({ type: 'undo' }));
      expect(savedAlike(a, b)).toEqual(['Hello']);
      ok(b.exec({ type: 'redo' }));
      expect(savedAlike(a, b)).toEqual(['Hel', 'lo']);
    });

    test(`text typed after the moved text goes with an Enter in a paragraph with an end (${order(swap)})`, async () => {
      const { a, b, sync } = await replicas(withEnd('Hello'), swap);
      sync.pause();
      sel(a, 3);
      ok(a.input('paragraph'));
      sel(b, 5);
      ok(b.input('text', '!'));
      sync.flush();
      expect(savedAlike(a, b)).toEqual(['Hel', 'lo!']);
    });

    test(`a delete of the moved text removes it, not "Hello" (${order(swap)})`, async () => {
      const { a, b } = await boldWhile(swap, (b) => (sel(b, 3, 5), ok(b.input('deleteBackward'))));
      expect(savedAlike(a, b)).toEqual(['Hel']);
      ok(b.exec({ type: 'undo' }));
      expect(savedAlike(a, b)).toEqual(['Hello']);
      ok(b.exec({ type: 'redo' }));
      expect(savedAlike(a, b)).toEqual(['Hel']);
    });

    for (const [name, edit, text] of [
      ['deletes', (e: PptxEditorInstance) => ok(e.input('deleteBackward')), 'Hello '],
      ['retypes', (e: PptxEditorInstance) => ok(e.input('text', 'there')), 'Hello there'],
    ] as const)
      test(`the peer whose format loses ${name} its word before the merge (${order(swap)})`, async () => {
        // Client 1001 wins a format tie (`dropLosingSplits`): the loser is the larger id.
        const { a, b, sync } = await replicas(oneRun('Hello world'), swap);
        const loser = swap ? a : b;
        sync.pause();
        sel(a, 6, 11);
        ok(a.exec({ type: 'toggleBold' }));
        sel(b, 6, 11);
        ok(b.exec({ type: 'toggleItalic' }));
        sel(loser, 6, 11);
        edit(loser);
        sync.flush();
        expect(savedAlike(a, b)).toEqual([text]);
      });

    test(`two formats and an Enter before the cut keep the word in its paragraph (${order(swap)})`, async () => {
      const { a, b, sync } = await replicas(twoParagraphs(), swap);
      selectIn(a, 1, 6, 11);
      selectIn(b, 1, 6, 11);
      sync.pause();
      ok(a.exec({ type: 'toggleBold' }));
      ok(b.exec({ type: 'toggleItalic' }));
      selectIn(b, 1, 3, 3);
      ok(b.input('paragraph'));
      sync.flush();
      expect(savedAlike(a, b)).toEqual(['First', 'Hel', 'lo world']);
      // The format tie keeps one peer's format (client 1001).
      expect(formatOf(a, 'world')).toEqual(swap ? { i: '1' } : { b: '1' });
    });
  }

  for (const ids of [
    [1_001, 2_002, 3_003],
    [3_003, 2_002, 1_001],
  ])
    test(`a third peer's delete reaches both formats and shows the rest once (client ids ${ids.join(', ')})`, async () => {
      const { editors, sync } = await peers(withEnd('ABC'), ids);
      const [a, b, c] = editors as [PptxEditorInstance, PptxEditorInstance, PptxEditorInstance];
      sync.pause();
      sel(a, 1, 3);
      ok(a.exec({ type: 'toggleBold' }));
      sel(b, 1, 3);
      ok(b.exec({ type: 'toggleItalic' }));
      sel(c, 2, 3);
      ok(c.input('deleteBackward'));
      sync.flush();
      expect(savedAlike(...editors)).toEqual(['AB']);
    });

  // An undo puts characters back as new characters at their old place. After a peer split the
  // same run, the undo shows each character once, in the run that now holds its neighbours
  // (`undo-restore.ts`).
  for (const swap of [false, true]) {
    for (const [bold, italic, runs] of [
      [
        [6, 11],
        [3, 8],
        ['Hel', 'lo wo', 'rld'],
      ],
      [
        [6, 11],
        [0, 3],
        ['Hel', 'lo world'],
      ],
      [
        [6, 11],
        [7, 9],
        ['Hello w', 'or', 'ld'],
      ],
    ] as const)
      test(`an undo of a format after a peer formatted ${italic.join('..')} again keeps "Hello world" (${order(swap)})`, async () => {
        const { a, b } = await replicas(oneRun('Hello world'), swap);
        sel(a, bold[0], bold[1]);
        ok(a.exec({ type: 'toggleBold' }));
        sel(b, italic[0], italic[1]);
        ok(b.exec({ type: 'toggleItalic' }));
        ok(a.exec({ type: 'undo' }));
        expect(savedAlike(a, b)).toEqual(['Hello world']);
        expect(inlineOf(a, 'H')).toEqual([...runs]);
        // The peer's format stays. A run the peer's split made keeps the bold it copied.
        expect(formatOf(a, 'Hello world'.slice(italic[0], italic[1]))).toMatchObject({ i: '1' });
        for (const [e, step] of [
          [a, 'redo'],
          [b, 'undo'],
          [a, 'undo'],
          [b, 'redo'],
          [a, 'redo'],
        ] as const) {
          ok(e.exec({ type: step }));
          expect(savedAlike(a, b)).toEqual(['Hello world']);
        }
      });

    test(`an undo of a format the peer formatted at the same time keeps the text (${order(swap)})`, async () => {
      const { a, b, sync } = await replicas(oneRun('Hello world'), swap);
      sync.pause();
      sel(a, 6, 11);
      ok(a.exec({ type: 'toggleBold' }));
      sel(b, 3, 8);
      ok(b.exec({ type: 'toggleItalic' }));
      sync.flush();
      expect(savedAlike(a, b)).toEqual(['Hello world']);
      for (const [e, step] of [
        [a, 'undo'],
        [b, 'undo'],
        [a, 'redo'],
        [b, 'redo'],
        [b, 'undo'],
        [a, 'undo'],
      ] as const) {
        ok(e.exec({ type: step }));
        expect(savedAlike(a, b)).toEqual(['Hello world']);
      }
    });

    test(`an undo of the format keeps the text a peer typed after the moved text (${order(swap)})`, async () => {
      const { a, b } = await boldWhile(swap, (b) => (sel(b, 5), ok(b.input('text', '!'))));
      ok(a.exec({ type: 'undo' }));
      expect(savedAlike(a, b)).toEqual(['Hello!']);
      expect(formatOf(a, 'Hello')).toEqual({});
      ok(a.exec({ type: 'redo' }));
      expect(savedAlike(a, b)).toEqual(['Hello!']);
      expect(formatOf(a, 'lo!')).toEqual({ b: '1' });
      ok(b.exec({ type: 'undo' }));
      expect(savedAlike(a, b)).toEqual(['Hello']);
      ok(b.exec({ type: 'redo' }));
      expect(savedAlike(a, b)).toEqual(['Hello!']);
    });
  }
});

describe('two Enters in one paragraph at once', () => {
  for (const swap of [false, true]) {
    test(`each paragraph shows its text once, not ["A", "BCDE", "CDE"] (${order(swap)})`, async () => {
      const { a, b, sync } = await replicas(oneRun('ABCDE'), swap);
      sync.pause();
      sel(a, 1);
      ok(a.input('paragraph'));
      sel(b, 2);
      ok(b.input('paragraph'));
      sync.flush();
      expect(savedAlike(a, b)).toEqual(['A', 'B', 'CDE']);
      // A later edit by either peer keeps the paragraphs.
      selectIn(b, 1, 1);
      ok(b.input('text', '!'));
      selectIn(a, 2, 3);
      ok(a.input('text', '?'));
      expect(savedAlike(a, b)).toEqual(['A', 'B!', 'CDE?']);
    });

    test(`undo and redo keep every character (${order(swap)})`, async () => {
      const { a, b, sync } = await replicas(oneRun('ABCDE'), swap);
      sync.pause();
      sel(a, 1);
      ok(a.input('paragraph'));
      sel(b, 2);
      ok(b.input('paragraph'));
      sync.flush();
      for (const [e, step] of [
        [b, 'undo'],
        [b, 'redo'],
        [a, 'undo'],
        [a, 'redo'],
      ] as const) {
        ok(e.exec({ type: step }));
        const text = savedAlike(a, b).join('');
        for (const ch of 'ABCDE') expect(text).toContain(ch);
      }
      expect(savedAlike(a, b)).toEqual(['A', 'B', 'CDE']);
    });

    test(`each paragraph shows its text once in a paragraph with an end (${order(swap)})`, async () => {
      const { a, b, sync } = await replicas(withEnd('ABCDE'), swap);
      sync.pause();
      sel(a, 1);
      ok(a.input('paragraph'));
      sel(b, 2);
      ok(b.input('paragraph'));
      sync.flush();
      expect(savedAlike(a, b)).toEqual(['A', 'B', 'CDE']);
    });

    test(`two Enters over selections delete both, not ["A", "CDE", "DE"] (${order(swap)})`, async () => {
      const { a, b, sync } = await replicas(withEnd('ABCDE'), swap);
      sync.pause();
      sel(a, 1, 2);
      ok(a.input('paragraph'));
      sel(b, 2, 3);
      ok(b.input('paragraph'));
      sync.flush();
      expect(savedAlike(a, b)).toEqual(['A', '', 'DE']);
    });

    test(`two Enters at one place leave an empty paragraph between (${order(swap)})`, async () => {
      const { a, b, sync } = await replicas(oneRun('ABCDE'), swap);
      sync.pause();
      sel(a, 2);
      ok(a.input('paragraph'));
      sel(b, 2);
      ok(b.input('paragraph'));
      sync.flush();
      expect(savedAlike(a, b)).toEqual(['AB', '', 'CDE']);
    });

    test(`an Enter while a peer deletes across it (${order(swap)})`, async () => {
      const { a, b, sync } = await replicas(oneRun('ABCDE'), swap);
      sync.pause();
      sel(a, 2);
      ok(a.input('paragraph'));
      sel(b, 1, 4);
      ok(b.input('deleteBackward'));
      sync.flush();
      expect(savedAlike(a, b)).toEqual(['A', 'E']);
    });
  }

  for (const ids of [
    [1_001, 2_002, 3_003],
    [3_003, 2_002, 1_001],
  ])
    test(`three Enters at once (client ids ${ids.join(', ')})`, async () => {
      const { editors, sync } = await peers(oneRun('ABCDE'), ids);
      sync.pause();
      editors.forEach((e, i) => {
        sel(e, i + 1);
        ok(e.input('paragraph'));
      });
      sync.flush();
      expect(savedAlike(...editors)).toEqual(['A', 'B', 'C', 'DE']);
    });

  for (const ids of [
    [1_001, 2_002, 3_003],
    [3_003, 2_002, 1_001],
  ])
    test(`a format of the text an Enter moved keeps it once (client ids ${ids.join(', ')})`, async () => {
      // A deletes "L" and bolds "IJ"; B presses Enter after "I" and italicizes "J".
      const { editors, sync } = await peers(withEnd('IJKL'), ids);
      const [a, b] = editors as [PptxEditorInstance, PptxEditorInstance];
      sync.pause();
      sel(a, 3, 4);
      ok(a.input('deleteBackward'));
      sel(a, 0, 2);
      ok(a.exec({ type: 'toggleBold' }));
      sel(b, 1);
      ok(b.input('paragraph'));
      selectIn(b, 1, 0, 1);
      ok(b.exec({ type: 'toggleItalic' }));
      sync.flush();
      expect(savedAlike(...editors)).toEqual(['I', 'JK']);
    });

  for (const swap of [false, true])
    for (const steps of [
      [
        ['b', 'undo', ['A', 'BCDE']],
        ['b', 'redo', ['A', 'B', 'CDE']],
        ['a', 'undo', ['AB', 'CDE']],
        ['a', 'redo', ['A', 'B', 'CDE']],
      ],
      [
        ['a', 'undo', ['AB', 'CDE']],
        ['a', 'redo', ['A', 'B', 'CDE']],
        ['b', 'undo', ['A', 'BCDE']],
        ['a', 'undo', ['AB', 'CDE']],
        ['a', 'redo', ['A', 'BCDE']],
        ['b', 'redo', ['A', 'B', 'CDE']],
      ],
    ] as const)
      test(`an undo of one Enter after the merge gives its text back once, ${steps[0]![0]} first (${order(swap)})`, async () => {
        const { a, b, sync } = await replicas(oneRun('ABCDE'), swap);
        sync.pause();
        sel(a, 1);
        ok(a.input('paragraph'));
        sel(b, 2);
        ok(b.input('paragraph'));
        sync.flush();
        for (const [who, step, text] of steps) {
          ok((who === 'a' ? a : b).exec({ type: step }));
          expect(savedAlike(a, b)).toEqual([...text]);
        }
      });
});

describe('a delete of a run a peer types into', () => {
  for (const swap of [false, true]) {
    test(`a range delete keeps the typed character, not "ABGH" (${order(swap)})`, async () => {
      const { a, b, sync } = await replicas(boldMiddle(), swap);
      sync.pause();
      sel(a, 2, 6);
      ok(a.input('deleteBackward'));
      sel(b, 4);
      ok(b.input('text', '!'));
      sync.flush();
      expect(savedAlike(a, b)).toEqual(['AB!GH']);
      expect(formatOf(a, '!')).toEqual({ b: '1' });
      ok(b.exec({ type: 'undo' }));
      expect(savedAlike(a, b)).toEqual(['ABGH']);
      ok(b.exec({ type: 'redo' }));
      expect(savedAlike(a, b)).toEqual(['AB!GH']);
      ok(a.exec({ type: 'undo' }));
      expect(savedAlike(a, b)).toEqual(['ABCD!EFGH']);
      ok(a.exec({ type: 'redo' }));
      expect(savedAlike(a, b)).toEqual(['AB!GH']);
      // An edit of the paragraph writes the run back into the room, and a delete removes it.
      sel(a, 2, 3);
      ok(a.input('deleteBackward'));
      expect(savedAlike(a, b)).toEqual(['ABGH']);
    });

    test(`a delete of the whole run keeps the typed character (${order(swap)})`, async () => {
      const { a, b, sync } = await replicas(boldMiddle(), swap);
      sync.pause();
      sel(a, 3, 5);
      ok(a.input('deleteBackward'));
      sel(b, 4);
      ok(b.input('text', '!'));
      sync.flush();
      expect(savedAlike(a, b)).toEqual(['ABC!FGH']);
    });

    test(`a delete of the run alone stays a delete (${order(swap)})`, async () => {
      const { a, b, sync } = await replicas(boldMiddle(), swap);
      sync.pause();
      sel(a, 3, 5);
      ok(a.input('deleteBackward'));
      sel(b, 0);
      ok(b.input('text', '>'));
      sync.flush();
      expect(savedAlike(a, b)).toEqual(['>ABCFGH']);
    });

    test(`a run that a peer's Enter moved and then deleted keeps the typed character (${order(swap)})`, async () => {
      const deck = textBox(
        para(run('ABCD') + run('EF', 'b="1"') + run('GH') + '<a:endParaRPr lang="en-US"/>')
      );
      const { a, b, sync } = await replicas(deck, swap);
      sync.pause();
      // A's Enter moves the run "GH" into a new paragraph, then an Enter over "GH" removes it.
      sel(a, 5);
      ok(a.input('paragraph'));
      selectIn(a, 1, 1, 3);
      ok(a.input('paragraph'));
      sel(b, 7);
      ok(b.input('text', '!'));
      sync.flush();
      expect(savedAlike(a, b)).toEqual(['ABCDE', 'F!', '']);
    });

    test(`a format of the run while a peer deletes it stays deleted (${order(swap)})`, async () => {
      const { a, b, sync } = await replicas(boldMiddle(), swap);
      sync.pause();
      sel(a, 3, 5);
      ok(a.input('deleteBackward'));
      sel(b, 4, 5);
      ok(b.exec({ type: 'toggleItalic' }));
      sync.flush();
      expect(savedAlike(a, b)).toEqual(['ABCFGH']);
    });
  }
});

// A run that came back for a peer's text is listed again by two writes at once: a peer types
// into it, and another peer's Enter moves it to a new paragraph. The read shows it in the
// paragraph it left, though its move mark names that paragraph. A later delete of its text
// while a peer types into it again keeps the typed text there: a move mark counts only while
// another record lists the run (`movedAway` in `text-identity.ts`).
describe("a run kept for a peer's text, then listed by two peers at once", () => {
  for (const swap of [false, true])
    test(`a delete of its text keeps the next typed character (${order(swap)})`, async () => {
      const deck = textBox(
        para(run('ABCD') + run('EF', 'b="1"') + run('GH') + '<a:endParaRPr lang="en-US"/>')
      );
      const { a, b, sync } = await replicas(deck, swap);
      sync.pause();
      sel(a, 0, 4);
      ok(a.input('deleteBackward'));
      sel(b, 0);
      ok(b.input('text', 'X'));
      sync.flush();
      expect(savedAlike(a, b)).toEqual(['XEFGH']);
      sync.pause();
      sel(b, 0);
      ok(b.input('paragraph'));
      sel(a, 1);
      ok(a.input('text', 'Y'));
      sync.flush();
      expect(savedAlike(a, b)).toEqual(['XY', 'EFGH']);
      sync.pause();
      sel(a, 0, 2);
      ok(a.input('deleteBackward'));
      sel(b, 0);
      ok(b.input('text', 'Z'));
      sync.flush();
      expect(savedAlike(a, b)).toEqual(['Z', 'EFGH']);
    });
});

describe('hostile identity values', () => {
  for (const [name, value] of [
    ['a copy mark of the wrong shape', 'not a mark'],
    ['a copy mark with a negative clock', [5, -1, 1, 0, 1]],
    ['a copy mark of characters no replica holds', [999, 0, 999, 10_000, 3]],
  ] as const)
    test(`${name} is ignored`, async () => {
      const { a, b, sync, docs } = await replicas(oneRun('Hello'));
      const nodes = docs[1].getMap('pptx-nodes-v1') as Y.Map<Y.Map<unknown>>;
      const textKey = [...nodes.keys()].find((k) => nodes.get(k)!.get(F.text) instanceof Y.Text)!;
      sync.pause();
      docs[1].transact(() => {
        const key = '999.x1';
        const rec = new Y.Map<unknown>();
        rec.set(F.part, slide1(b));
        rec.set(F.key, key);
        rec.set(F.text, new Y.Text('lo'));
        rec.set(IDENTITY.copyOf, value);
        rec.set(IDENTITY.newParagraph, true);
        nodes.set(key, rec);
        docs[1].getMap(COPIES_KEY).set(key, slide1(b));
        nodes.get(textKey)!.set(IDENTITY.copyOf, value);
      }, 'hostile');
      sel(a, 3, 5);
      ok(a.exec({ type: 'toggleBold' }));
      sel(b, 5);
      ok(b.input('text', '!'));
      sync.flush();
      expect(savedAlike(a, b)).toEqual(['Hello!']);
    });

  test('a copy mark that names its own characters (a cycle) is ignored', async () => {
    const { a, b, sync, docs } = await replicas(oneRun('Hello'));
    sync.pause();
    sel(a, 3, 5);
    ok(a.exec({ type: 'toggleBold' }));
    sel(b, 5);
    ok(b.input('text', '!'));
    sync.deliverTo(docs[1]);
    // Peer B turns the mark of A's copy into a loop through its own characters.
    const nodes = docs[1].getMap('pptx-nodes-v1') as Y.Map<Y.Map<unknown>>;
    docs[1].transact(() => {
      for (const key of docs[1].getMap(COPIES_KEY).keys()) {
        const mark = nodes.get(key)!.get(IDENTITY.copyOf) as number[];
        nodes.get(key)!.set(IDENTITY.copyOf, [mark[0]!, mark[1]!, mark[0]!, mark[1]!, 2]);
      }
    }, 'hostile');
    sync.flush();
    // Every replica reads the same room: the text stays, and the replicas agree.
    expect(savedAlike(a, b).join('')).toContain('Hello');
  });

  test('removed-run records with prototype keys, unknown keys, and keys of another part are ignored', async () => {
    const { a, b, sync, docs } = await replicas(boldMiddle());
    const nodes = docs[1].getMap('pptx-nodes-v1') as Y.Map<Y.Map<unknown>>;
    const paragraph = [...nodes.keys()].find(
      (k) => (nodes.get(k)!.get(F.shell) as { local?: string } | undefined)?.local === 'p'
    )!;
    const otherPart = [...nodes.keys()].find((k) => nodes.get(k)!.get(F.part) !== slide1(b))!;
    sync.pause();
    docs[1].transact(() => {
      const rec = nodes.get(paragraph)!;
      for (const key of ['__proto__', 'constructor', 'toString', 'missing', otherPart, paragraph])
        rec.set(`${IDENTITY.removedRun}${key}`, key);
    }, 'hostile');
    sel(a, 3, 5);
    ok(a.input('deleteBackward'));
    sync.flush();
    expect(savedAlike(a, b)).toEqual(['ABCFGH']);
  });

  test('a removed run that the room refuses to read reaches each session and heals', async () => {
    const { a, b, sync, docs } = await replicas(boldMiddle());
    sync.pause();
    sel(a, 2, 6);
    ok(a.input('deleteBackward'));
    sel(b, 4);
    ok(b.input('text', '!'));
    sync.flush();
    expect(savedAlike(a, b)).toEqual(['AB!GH']);
    // A peer gives the removed run more copies of its text element than a read takes.
    const doc = docs[1];
    const nodes = doc.getMap('pptx-nodes-v1') as Y.Map<Y.Map<unknown>>;
    const listing = (child: string) =>
      [...nodes.keys()].find((k) => {
        const list = nodes.get(k)!.get(F.children);
        return list instanceof Y.Array && (list.toArray() as string[]).includes(child);
      })!;
    const bang = [...nodes.keys()].find((k) => nodes.get(k)!.get(F.text)?.toString() === '!')!;
    const runKey = listing(listing(bang));
    const hostile = Array.from({ length: ROOM_LIMITS.maxCopiesPerGroup + 2 }, (_, i) => `99.h${i}`);
    doc.transact(() => {
      for (const k of hostile) {
        const rec = new Y.Map<unknown>();
        rec.set(F.part, slide1(a));
        rec.set(F.key, k);
        rec.set(F.shell, { ns: A_NS, local: 't', prefix: 'a', nsDecls: [] });
        rec.set(F.attrs, new Y.Map());
        rec.set(F.children, new Y.Array());
        nodes.set(k, rec);
      }
      (nodes.get(runKey)!.get(F.children) as Y.Array<string>).push(hostile);
    }, 'hostile');
    for (const s of sessionsOf([a, b]))
      expect(s.statusSnapshot()).toMatchObject({
        status: 'error',
        reason: { code: 'remote-apply-failed' },
      });
    // The stores keep their last good state, and the next good update heals them.
    doc.transact(() => {
      (nodes.get(runKey)!.get(F.children) as Y.Array<string>).delete(2, hostile.length);
    }, 'hostile');
    expect(savedAlike(a, b)).toEqual(['AB!GH']);
  });
});
