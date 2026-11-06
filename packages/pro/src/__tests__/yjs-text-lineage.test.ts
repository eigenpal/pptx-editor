/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Text that concurrent structure edits duplicated or lost (openspec change
// `collaboration-text-lineage`): two editors on two linked Y.Docs, the link paused, the edits
// on both, then a flush. Each case compares the text, not only the fingerprints, because the
// replicas converge on the wrong text too.

import { describe, expect, test } from 'bun:test';
import type { PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import { PIECES_KEY, SPLITS_KEY } from '../collaboration/yjs/text-lineage.ts';
import {
  all,
  converged,
  deleteMergedFirstRowOps,
  firstRow,
  formatOf,
  kids,
  mergeOps,
  oneRun,
  paragraphs,
  replicas,
  root,
  selectIn,
  slide1,
  tableDeck,
  transact,
  twoParagraphs,
} from './lineage-support.ts';

const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';

/** Selects `from`..`to` in the first paragraph of the slide's first shape. */
const selectText = (e: PptxEditorInstance, from: number, to: number) => selectIn(e, 0, from, to);

describe('splits one after the other never drop text', () => {
  for (const swap of [false, true])
    test(`a split, its undo, then a peer's split of the same run (client order ${swap ? 'B, A' : 'A, B'})`, async () => {
      const { a, b } = await replicas(oneRun('Hello world'), swap);
      selectText(a, 6, 11);
      expect(a.exec({ type: 'toggleBold' }).ok).toBe(true);
      expect(a.exec({ type: 'undo' }).ok).toBe(true);
      selectText(b, 6, 11);
      expect(b.exec({ type: 'toggleItalic' }).ok).toBe(true);
      expect(converged(a, b)).toEqual(['Hello world']);
      // A write on the other peer keeps it.
      selectText(a, 0, 0);
      expect(a.input('text', '>').ok).toBe(true);
      expect(converged(a, b)).toEqual(['>Hello world']);
    });
});

describe('a hostile split index entry', () => {
  for (const swap of [false, true])
    for (const [name, cut, text] of [
      ['a negative cut wins nothing', -5, 'Hello world'],
      ['a NaN cut wins nothing', Number.NaN, 'Hello world'],
      // A valid earlier cut by a client that never split drops the pieces (design.md, hostile
      // only); both replicas read it at once, though the entry alone touches no record.
      ['an entry alone is read by every replica at once', 0, 'Hello '],
    ] as const)
      test(`${name} (client order ${swap ? 'B, A' : 'A, B'})`, async () => {
        const { a, b, docs } = await replicas(oneRun('Hello world'), swap);
        selectText(a, 6, 11);
        expect(a.exec({ type: 'toggleBold' }).ok).toBe(true);
        const splits = docs[1].getMap(SPLITS_KEY);
        const [entry] = [...splits.keys()];
        const group = entry!.split('\u0000').slice(0, 3).join('\u0000');
        docs[1].transact(() => splits.set(`${group}\u00000`, cut), 'hostile');
        expect(converged(a, b)).toEqual([text]);
        for (const e of [a, b]) {
          selectText(e, 0, 0);
          expect(e.input('text', '>').ok).toBe(true);
        }
        expect(converged(a, b)).toEqual([`>>${text}`]);
      });
});

describe('one peer alone: a split never drops text', () => {
  for (const [start, typed, from, to] of [
    ['Hello', 'lo', 3, 5],
    ['1000', '0', 3, 4],
  ] as const)
    test(`format the end of "${start}", then type "${typed}" before it`, async () => {
      const { a, b } = await replicas(oneRun(start));
      selectText(a, from, to);
      expect(a.exec({ type: 'toggleBold' }).ok).toBe(true);
      selectText(a, from, from);
      expect(a.input('text', typed).ok).toBe(true);
      const text = start.slice(0, from) + typed + start.slice(from);
      expect(paragraphs(a)).toEqual([text]);
      // A peer edit elsewhere writes the paragraph back; the text stays.
      selectText(b, 0, 0);
      expect(b.input('text', '>').ok).toBe(true);
      expect(converged(a, b)).toEqual([`>${text}`]);
    });
});

describe('two peers that format one run at once', () => {
  for (const swap of [false, true])
    test(`the same word: the text shows once, not "worldworld" (client order ${swap ? 'B, A' : 'A, B'})`, async () => {
      const { a, b, sync } = await replicas(oneRun('Hello world'), swap);
      selectText(a, 6, 11);
      selectText(b, 6, 11);
      sync.pause();
      expect(a.exec({ type: 'toggleBold' }).ok).toBe(true);
      expect(b.exec({ type: 'toggleItalic' }).ok).toBe(true);
      sync.flush();
      expect(converged(a, b)).toEqual(['Hello world']);
      // The winner (client 1001) keeps its format: bold from a, or italic from b.
      expect(formatOf(a, 'world')).toEqual(swap ? { i: '1' } : { b: '1' });
    });

  for (const swap of [false, true])
    test(`overlapping ranges: the text stays "Hello", not "Hellllo" (client order ${swap ? 'B, A' : 'A, B'})`, async () => {
      const { a, b, sync } = await replicas(oneRun('Hello'), swap);
      selectText(a, 1, 3);
      selectText(b, 2, 4);
      sync.pause();
      expect(a.exec({ type: 'toggleBold' }).ok).toBe(true);
      expect(b.exec({ type: 'toggleItalic' }).ok).toBe(true);
      sync.flush();
      expect(converged(a, b)).toEqual(['Hello']);
      // Peer a cut first (at 1), so it wins in both orders and keeps its bold.
      expect(formatOf(a, 'el')).toEqual({ b: '1' });
    });

  for (const swap of [false, true])
    test(`disjoint ranges: the text shows once, with one peer's format (#16, client order ${swap ? 'B, A' : 'A, B'})`, async () => {
      const { a, b, sync } = await replicas(oneRun('Hello big world'), swap);
      selectText(a, 6, 9);
      selectText(b, 10, 15);
      sync.pause();
      expect(a.exec({ type: 'toggleBold' }).ok).toBe(true);
      expect(b.exec({ type: 'toggleItalic' }).ok).toBe(true);
      sync.flush();
      expect(converged(a, b)).toEqual(['Hello big world']);
      // Peer a cut first (at 6), so it wins in both orders; b's italic is lost (#16).
      expect(formatOf(a, 'big')).toEqual({ b: '1' });
    });

  test('a split, then typing in a piece: the typed text survives, once', async () => {
    const { a, b, sync } = await replicas(oneRun('Hello world'));
    selectText(a, 6, 11);
    selectText(b, 6, 11);
    sync.pause();
    expect(a.exec({ type: 'toggleBold' }).ok).toBe(true);
    expect(b.exec({ type: 'toggleUnderline' }).ok).toBe(true);
    sync.flush();
    // After the merge, a peer types at the end of the piece that is left.
    selectText(a, 11, 11);
    expect(a.input('text', '!').ok).toBe(true);
    expect(converged(a, b)).toEqual(['Hello world!']);
  });

  test('an undo by the winner of a split of one word: the word shows once', async () => {
    const { a, b, sync } = await replicas(oneRun('Hello world'));
    selectText(a, 6, 11);
    selectText(b, 6, 11);
    sync.pause();
    expect(a.exec({ type: 'toggleBold' }).ok).toBe(true);
    expect(b.exec({ type: 'toggleItalic' }).ok).toBe(true);
    sync.flush();
    // Client 1001 (peer a) wins the tie.
    expect(a.exec({ type: 'undo' }).ok).toBe(true);
    expect(converged(a, b)).toEqual(['Hello world']);
  });

  test('an undo and a redo by the loser keep the text, never lose it', async () => {
    const { a, b, sync } = await replicas(oneRun('Hello world'));
    selectText(a, 6, 11);
    selectText(b, 6, 11);
    sync.pause();
    expect(a.exec({ type: 'toggleBold' }).ok).toBe(true);
    expect(b.exec({ type: 'toggleItalic' }).ok).toBe(true);
    sync.flush();
    expect(b.exec({ type: 'undo' }).ok).toBe(true);
    // The winner's pieces hold the word, so the loser's undo puts none of it back.
    expect(converged(a, b)).toEqual(['Hello world']);
    expect(b.exec({ type: 'redo' }).ok).toBe(true);
    expect(converged(a, b)).toEqual(['Hello world']);
  });

  test('a second format by the loser of part of its piece shows the word once', async () => {
    const { a, b, sync } = await replicas(oneRun('Hello world'));
    selectText(a, 6, 11);
    selectText(b, 6, 11);
    sync.pause();
    expect(a.exec({ type: 'toggleBold' }).ok).toBe(true);
    expect(b.exec({ type: 'toggleItalic' }).ok).toBe(true);
    selectText(b, 7, 9);
    expect(b.exec({ type: 'toggleUnderline' }).ok).toBe(true);
    sync.flush();
    expect(converged(a, b)).toEqual(['Hello world']);
  });

  test('text typed into the origin run after a settled split is kept', async () => {
    const { a, b, sync } = await replicas(oneRun('Hello world'));
    selectText(a, 6, 11);
    selectText(b, 6, 11);
    sync.pause();
    expect(a.exec({ type: 'toggleBold' }).ok).toBe(true);
    expect(b.exec({ type: 'toggleItalic' }).ok).toBe(true);
    sync.flush();
    selectText(b, 6, 6);
    expect(b.input('text', 'world').ok).toBe(true);
    expect(converged(a, b)).toEqual(['Hello worldworld']);
  });

  test('an undo by the winner of overlapping splits: the text stays "Hello"', async () => {
    const { a, b, sync } = await replicas(oneRun('Hello'));
    selectText(a, 1, 3);
    selectText(b, 2, 4);
    sync.pause();
    expect(a.exec({ type: 'toggleBold' }).ok).toBe(true);
    expect(b.exec({ type: 'toggleItalic' }).ok).toBe(true);
    sync.flush();
    expect(a.exec({ type: 'undo' }).ok).toBe(true);
    expect(converged(a, b)).toEqual(['Hello']);
  });

  // An undo or a redo of concurrent splits is in `yjs-undo-lineage.test.ts`; a loser's delete of
  // its piece is in `yjs-text-structure.test.ts`.

  test('an undo and a redo by the winner keep the word once', async () => {
    const { a, b, sync } = await replicas(oneRun('Hello world'));
    selectText(a, 6, 11);
    selectText(b, 6, 11);
    sync.pause();
    expect(a.exec({ type: 'toggleBold' }).ok).toBe(true);
    expect(b.exec({ type: 'toggleItalic' }).ok).toBe(true);
    sync.flush();
    expect(a.exec({ type: 'undo' }).ok).toBe(true);
    expect(a.exec({ type: 'redo' }).ok).toBe(true);
    expect(converged(a, b)).toEqual(['Hello world']);
  });

  test('a peer that formats and then deletes the word at once with the other: the word stays deleted', async () => {
    const { a, b, sync } = await replicas(oneRun('Hello world'));
    selectText(a, 6, 11);
    selectText(b, 6, 11);
    sync.pause();
    expect(a.exec({ type: 'toggleBold' }).ok).toBe(true);
    selectText(a, 5, 11);
    expect(a.input('deleteBackward').ok).toBe(true);
    expect(b.exec({ type: 'toggleItalic' }).ok).toBe(true);
    sync.flush();
    // Peer a wins the tie (client 1001), and its piece is gone: peer b's piece does not show.
    expect(converged(a, b)).toEqual(['Hello']);
  });

  test('text that either peer types into its own piece is kept, never lost', async () => {
    const { a, b, sync } = await replicas(oneRun('Hello world'));
    selectText(a, 6, 11);
    selectText(b, 6, 11);
    sync.pause();
    expect(a.exec({ type: 'toggleBold' }).ok).toBe(true);
    expect(b.exec({ type: 'toggleItalic' }).ok).toBe(true);
    selectText(a, 11, 11);
    expect(a.input('text', '!').ok).toBe(true);
    selectText(b, 11, 11);
    expect(b.input('text', '?').ok).toBe(true);
    sync.flush();
    // Duplicated text is better than lost text: both edited pieces stay.
    expect(converged(a, b)).toEqual(['Hello world!world?']);
  });

  test('a line break typed at the cut against a format of the word shows the word once', async () => {
    const { a, b, sync } = await replicas(oneRun('Hello world'));
    selectText(a, 6, 11);
    selectText(b, 6, 6);
    sync.pause();
    expect(a.exec({ type: 'toggleBold' }).ok).toBe(true);
    expect(b.input('lineBreak').ok).toBe(true);
    sync.flush();
    const [text] = converged(a, b);
    expect(text).toBe('Hello world');
  });

  test('text the loser types while a peer writes the merged paragraph back is kept', async () => {
    const { a, b, sync, docs } = await replicas(oneRun('Hello world'));
    selectText(a, 6, 11);
    selectText(b, 6, 11);
    sync.pause();
    expect(a.exec({ type: 'toggleBold' }).ok).toBe(true);
    expect(b.exec({ type: 'toggleItalic' }).ok).toBe(true);
    // Peer a gets b's split first, drops b's piece, and writes the paragraph back.
    sync.deliverTo(docs[0]);
    selectText(a, 0, 0);
    expect(a.input('text', 'X').ok).toBe(true);
    // Before that reaches b, b types into its own piece.
    selectText(b, 11, 11);
    expect(b.input('text', ' there').ok).toBe(true);
    sync.flush();
    // Kept, never lost; the edited piece shows its word again (design.md, "Known gaps").
    expect(converged(a, b)).toEqual(['XHello worldworld there']);
  });

  test('text the loser types after a line break at the cut, while a peer writes back, is kept', async () => {
    const { a, b, sync, docs } = await replicas(oneRun('Hello world'));
    selectText(a, 6, 11);
    selectText(b, 6, 6);
    sync.pause();
    expect(a.exec({ type: 'toggleBold' }).ok).toBe(true);
    expect(b.input('lineBreak').ok).toBe(true);
    sync.deliverTo(docs[0]);
    selectText(a, 0, 0);
    expect(a.input('text', 'X').ok).toBe(true);
    selectText(b, 12, 12);
    expect(b.input('text', ' there').ok).toBe(true);
    sync.flush();
    // Line breaks are not text here; the edited piece shows its word again (a known gap).
    expect(converged(a, b)).toEqual(['XHello worldworld there']);
  });

  for (const swap of [false, true])
    test(`text the loser types is kept when a peer deletes the run before its piece (client order ${swap ? 'B, A' : 'A, B'})`, async () => {
      const { a, b, sync, docs } = await replicas(oneRun('Hello world'), swap);
      selectText(a, 4, 11);
      selectText(b, 6, 11);
      sync.pause();
      expect(a.exec({ type: 'toggleBold' }).ok).toBe(true);
      expect(b.exec({ type: 'toggleItalic' }).ok).toBe(true);
      // Peer a reads b's split, then deletes "Hell", the run before b's piece.
      sync.deliverTo(docs[0]);
      selectText(a, 0, 4);
      expect(a.input('deleteBackward').ok).toBe(true);
      selectText(b, 11, 11);
      expect(b.input('text', '!!').ok).toBe(true);
      sync.flush();
      expect(converged(a, b)).toEqual(['o worldworld!!']);
    });

  for (const swap of [false, true])
    for (const edit of ['join', 'enter'] as const)
      test(`text the loser types while a peer ${edit === 'join' ? 'joins the paragraphs' : 'presses Enter in the run'} is kept (client order ${swap ? 'B, A' : 'A, B'})`, async () => {
        const { a, b, sync, docs } = await replicas(twoParagraphs(), swap);
        selectIn(a, 1, 6, 11);
        selectIn(b, 1, 6, 11);
        sync.pause();
        expect(a.exec({ type: 'toggleBold' }).ok).toBe(true);
        expect(b.exec({ type: 'toggleItalic' }).ok).toBe(true);
        sync.deliverTo(docs[0]);
        if (edit === 'join') {
          selectIn(a, 1, 0, 0);
          expect(a.input('deleteBackward').ok).toBe(true);
        } else {
          selectIn(a, 1, 3, 3);
          expect(a.input('paragraph').ok).toBe(true);
        }
        selectIn(b, 1, 11, 11);
        expect(b.input('text', ' there').ok).toBe(true);
        sync.flush();
        // The kept piece follows the winner's piece, so the word order stays. When b wins (client
        // order B, A), its edited piece is the one that shows.
        const word = swap ? 'world there' : 'worldworld there';
        expect(converged(a, b).join('|')).toBe(
          edit === 'join' ? `FirstHello ${word}` : `First|Hel|lo ${word}`
        );
      });

  for (const swap of [false, true])
    for (const joiner of ['a', 'b'] as const)
      test(`the word stays when one formatting peer joins the paragraphs at once (${joiner} joins, client order ${swap ? 'B, A' : 'A, B'})`, async () => {
        const { a, b, sync } = await replicas(twoParagraphs(), swap);
        selectIn(a, 1, 6, 11);
        selectIn(b, 1, 6, 11);
        sync.pause();
        expect(a.exec({ type: 'toggleBold' }).ok).toBe(true);
        expect(b.exec({ type: 'toggleItalic' }).ok).toBe(true);
        const j = joiner === 'a' ? a : b;
        selectIn(j, 1, 0, 0);
        expect(j.input('deleteBackward').ok).toBe(true);
        sync.flush();
        // The winner's pieces went with the paragraph that the join removed without seeing
        // them, so the joiner's pieces hold the word.
        expect(converged(a, b)).toEqual(['FirstHello world']);
      });

  for (const swap of [false, true])
    test(`a format while a peer joins the paragraphs keeps the formatted word (client order ${swap ? 'B, A' : 'A, B'})`, async () => {
      const { a, b, sync } = await replicas(twoParagraphs(), swap);
      selectIn(a, 1, 6, 11);
      sync.pause();
      expect(a.exec({ type: 'toggleBold' }).ok).toBe(true);
      selectIn(b, 1, 0, 0);
      expect(b.input('deleteBackward').ok).toBe(true);
      sync.flush();
      expect(converged(a, b)).toEqual(['FirstHello world']);
      expect(formatOf(a, 'world').b).toBe('1');
      // A later edit by either peer keeps it.
      selectIn(b, 0, 0, 0);
      expect(b.input('text', 'Z').ok).toBe(true);
      expect(converged(a, b)).toEqual(['ZFirstHello world']);
    });

  for (const swap of [false, true])
    test(`undo and redo of the join and of the format keep the word (client order ${swap ? 'B, A' : 'A, B'})`, async () => {
      const { a, b, sync } = await replicas(twoParagraphs(), swap);
      selectIn(a, 1, 6, 11);
      sync.pause();
      expect(a.exec({ type: 'toggleBold' }).ok).toBe(true);
      selectIn(b, 1, 0, 0);
      expect(b.input('deleteBackward').ok).toBe(true);
      sync.flush();
      expect(converged(a, b)).toEqual(['FirstHello world']);
      expect(b.exec({ type: 'undo' }).ok).toBe(true);
      expect(converged(a, b)).toEqual(['First', 'Hello world']);
      expect(formatOf(a, 'world').b).toBe('1');
      expect(a.exec({ type: 'undo' }).ok).toBe(true);
      expect(converged(a, b)).toEqual(['First', 'Hello world']);
      expect(formatOf(a, 'world').b).toBeUndefined();
      expect(a.exec({ type: 'redo' }).ok).toBe(true);
      expect(b.exec({ type: 'redo' }).ok).toBe(true);
      expect(converged(a, b)).toEqual(['FirstHello world']);
      expect(formatOf(b, 'world').b).toBe('1');
    });

  for (const swap of [false, true])
    test(`a format of the middle of a word while a peer joins keeps the text in order (client order ${swap ? 'B, A' : 'A, B'})`, async () => {
      const { a, b, sync } = await replicas(twoParagraphs(), swap);
      selectIn(a, 1, 7, 9);
      sync.pause();
      expect(a.exec({ type: 'toggleBold' }).ok).toBe(true);
      selectIn(b, 1, 0, 0);
      expect(b.input('deleteBackward').ok).toBe(true);
      sync.flush();
      expect(converged(a, b)).toEqual(['FirstHello world']);
    });

  for (const swap of [false, true])
    test(`a format while a peer deletes the whole paragraph: the text stays deleted (client order ${swap ? 'B, A' : 'A, B'})`, async () => {
      const { a, b, sync } = await replicas(twoParagraphs(), swap);
      selectIn(a, 1, 6, 11);
      sync.pause();
      expect(a.exec({ type: 'toggleBold' }).ok).toBe(true);
      const nodeId = b.session!.layout(0)!.items[0]!.element.nodeId;
      b.setSelection({
        kind: 'text',
        nodeId,
        anchor: { paragraph: 0, offset: 5 },
        focus: { paragraph: 1, offset: 11 },
      });
      expect(b.input('deleteBackward').ok).toBe(true);
      sync.flush();
      expect(converged(a, b)).toEqual(['First']);
    });

  test('dead entries of the piece index do not count against its cap', async () => {
    const { a, b, sync, docs } = await replicas(twoParagraphs());
    // Entries of pieces whose records are gone, as an undo of a split leaves them.
    docs[0].transact(() => {
      const index = docs[0].getMap(PIECES_KEY);
      for (let i = 0; i < 5_000; i += 1) index.set(`zz.${i}`, slide1(a));
    });
    selectIn(a, 1, 6, 11);
    selectIn(b, 1, 6, 11);
    sync.pause();
    expect(a.exec({ type: 'toggleBold' }).ok).toBe(true);
    expect(b.exec({ type: 'toggleItalic' }).ok).toBe(true);
    sync.deliverTo(docs[0]);
    selectIn(a, 1, 0, 0);
    expect(a.input('deleteBackward').ok).toBe(true);
    selectIn(b, 1, 11, 11);
    expect(b.input('text', ' there').ok).toBe(true);
    sync.flush();
    expect(converged(a, b).join('|')).toContain(' there');
  });

  for (const swap of [false, true])
    test(`two formats and an Enter before the cut keep the text (#17, client order ${swap ? 'B, A' : 'A, B'})`, async () => {
      const { a, b, sync } = await replicas(twoParagraphs(), swap);
      selectIn(a, 1, 6, 11);
      selectIn(b, 1, 6, 11);
      sync.pause();
      expect(a.exec({ type: 'toggleBold' }).ok).toBe(true);
      expect(b.exec({ type: 'toggleItalic' }).ok).toBe(true);
      selectIn(b, 1, 3, 3);
      expect(b.input('paragraph').ok).toBe(true);
      sync.flush();
      // No text is lost, and the word goes with the Enter (yjs-text-structure.test.ts).
      expect(converged(a, b)).toEqual(['First', 'Hel', 'lo world']);
    });

  for (const swap of [false, true])
    test(`text the loser types while a peer deletes across both paragraphs is kept (client order ${swap ? 'B, A' : 'A, B'})`, async () => {
      const { a, b, sync, docs } = await replicas(twoParagraphs(), swap);
      selectIn(a, 1, 6, 11);
      selectIn(b, 1, 6, 11);
      sync.pause();
      expect(a.exec({ type: 'toggleBold' }).ok).toBe(true);
      expect(b.exec({ type: 'toggleItalic' }).ok).toBe(true);
      sync.deliverTo(docs[0]);
      const nodeId = a.session!.layout(0)!.items[0]!.element.nodeId;
      a.setSelection({
        kind: 'text',
        nodeId,
        anchor: { paragraph: 0, offset: 3 },
        focus: { paragraph: 1, offset: 3 },
      });
      expect(a.input('deleteBackward').ok).toBe(true);
      selectIn(b, 1, 11, 11);
      expect(b.input('text', ' there').ok).toBe(true);
      sync.flush();
      expect(converged(a, b)).toEqual([swap ? 'Firlo world there' : 'Firlo worldworld there']);
    });

  test('the loser types into its piece and takes it back: the replicas stay equal', async () => {
    const { a, b, sync } = await replicas(oneRun('Hello world'));
    selectText(a, 6, 11);
    selectText(b, 6, 11);
    sync.pause();
    expect(a.exec({ type: 'toggleBold' }).ok).toBe(true);
    expect(b.exec({ type: 'toggleItalic' }).ok).toBe(true);
    selectText(b, 11, 11);
    expect(b.input('text', '!!').ok).toBe(true);
    sync.flush();
    expect(converged(a, b)).toEqual(['Hello worldworld!!']);
    // An edit marks the piece for good, so taking the text back keeps it on both replicas.
    const at = paragraphs(b)[0]!.length;
    selectText(b, at - 2, at);
    expect(b.input('deleteBackward').ok).toBe(true);
    expect(converged(a, b)).toEqual(['Hello worldworld']);
  });

  test('a shown piece of the loser that a user then deletes stays deleted', async () => {
    const { a, b, sync } = await replicas(oneRun('Hello world'));
    selectText(a, 6, 11);
    selectText(b, 6, 11);
    sync.pause();
    expect(a.exec({ type: 'toggleBold' }).ok).toBe(true);
    expect(b.exec({ type: 'toggleItalic' }).ok).toBe(true);
    selectText(b, 11, 11);
    expect(b.input('text', '?').ok).toBe(true);
    sync.flush();
    expect(converged(a, b)).toEqual(['Hello worldworld?']);
    selectText(a, 11, 17);
    expect(a.input('deleteBackward').ok).toBe(true);
    expect(converged(a, b)).toEqual(['Hello world']);
  });

  for (const swap of [false, true])
    test(`text the loser types into a piece of a second concurrent split is kept (client order ${swap ? 'B, A' : 'A, B'})`, async () => {
      const { a, b, sync } = await replicas(oneRun('Hello world'), swap);
      selectText(a, 6, 11);
      selectText(b, 6, 11);
      sync.pause();
      expect(a.exec({ type: 'toggleBold' }).ok).toBe(true);
      expect(b.exec({ type: 'toggleItalic' }).ok).toBe(true);
      selectText(b, 11, 11);
      expect(b.input('text', '?').ok).toBe(true);
      sync.flush();
      const len = paragraphs(a)[0]!.length;
      // A second round: both peers format the end of the paragraph at once, b types into its
      // piece, and a deletes its own piece.
      selectText(a, len - 4, len);
      selectText(b, len - 4, len);
      sync.pause();
      expect(a.exec({ type: 'toggleUnderline' }).ok).toBe(true);
      expect(b.exec({ type: 'toggleBold' }).ok).toBe(true);
      selectText(b, len - 2, len - 2);
      expect(b.input('text', 'Q').ok).toBe(true);
      selectText(a, len - 4, len);
      expect(a.input('deleteBackward').ok).toBe(true);
      sync.flush();
      const [text] = converged(a, b);
      expect(text).toContain('Q');
      // A later edit keeps it.
      selectText(a, 0, 0);
      expect(a.input('text', 'Z').ok).toBe(true);
      expect(converged(a, b)[0]).toContain('Q');
    });

  for (const swap of [false, true])
    test(`a delete across paragraphs removes the loser's edited piece for good (client order ${swap ? 'B, A' : 'A, B'})`, async () => {
      const { a, b, sync } = await replicas(oneRun('Hello world'), swap);
      selectText(a, 6, 11);
      selectText(b, 6, 11);
      sync.pause();
      expect(a.exec({ type: 'toggleBold' }).ok).toBe(true);
      expect(b.exec({ type: 'toggleItalic' }).ok).toBe(true);
      selectText(b, 11, 11);
      expect(b.input('text', '?').ok).toBe(true);
      sync.flush();
      selectText(a, 5, 5);
      expect(a.input('paragraph').ok).toBe(true);
      const nodeId = a.session!.layout(0)!.items[0]!.element.nodeId;
      a.setSelection({
        kind: 'text',
        nodeId,
        anchor: { paragraph: 0, offset: 5 },
        focus: { paragraph: 1, offset: paragraphs(a)[1]!.length },
      });
      expect(a.input('deleteBackward').ok).toBe(true);
      expect(converged(a, b)).toEqual(['Hello']);
      selectText(b, 0, 0);
      expect(b.input('text', 'Z').ok).toBe(true);
      expect(converged(a, b)).toEqual(['ZHello']);
    });

  test('two splits one after the other (not at once) keep every piece', async () => {
    const { a, b } = await replicas(oneRun('Hello world'));
    selectText(a, 6, 11);
    expect(a.exec({ type: 'toggleBold' }).ok).toBe(true);
    selectText(b, 0, 2);
    expect(b.exec({ type: 'toggleItalic' }).ok).toBe(true);
    expect(converged(a, b)).toEqual(['Hello world']);
  });
});

describe('a merge while a peer deletes the row of the merged cell', () => {
  for (const swap of [false, true])
    test(`the moved text goes back to the cell it came from (client order ${swap ? 'B, A' : 'A, B'})`, async () => {
      const { a, b, sync } = await replicas(tableDeck(), swap);
      sync.pause();
      transact(a, mergeOps(a));
      transact(b, [{ op: 'removeNode', part: slide1(b), node: firstRow(b).id }]);
      sync.flush();
      // Row 1 is gone; "B" goes back to its old place in the lower cell, and the empty
      // paragraph the merge left there goes.
      expect(converged(a, b)).toEqual(['B']);
      // The rescued cell keeps an orphan `vMerge`, which the store refuses only when a local
      // edit introduces it: an edit of the same slide still commits, and reaches the peer.
      const t = all(root(b), 't')[0]!;
      transact(b, [{ op: 'setText', part: slide1(b), node: t.id, text: 'B!' }]);
      expect(converged(a, b)).toEqual(['B!']);
    });

  // The rescued cell keeps `vMerge="1"` with no cell above it. Its reference PDF draws it and
  // its text; the editor's cascade hides such a cell (older code).
  test.todo('the rescued text in a cell that keeps vMerge="1" is in the layout');

  test('a delete that saw the merge removes the merged text with the row', async () => {
    const { a, b } = await replicas(tableDeck());
    transact(a, mergeOps(a));
    transact(b, deleteMergedFirstRowOps(b));
    expect(converged(a, b)).toEqual(['']);
  });

  for (const swap of [false, true])
    test(`two deletes of the row, one that saw the merge, remove the merged text (client order ${swap ? 'B, A' : 'A, B'})`, async () => {
      const { a, b, sync } = await replicas(tableDeck(), swap);
      sync.pause();
      transact(a, mergeOps(a));
      transact(a, deleteMergedFirstRowOps(a));
      transact(b, [{ op: 'removeNode', part: slide1(b), node: firstRow(b).id }]);
      sync.flush();
      // Each removal keeps its own snapshot, so a's, which saw "B", still counts.
      expect(converged(a, b)).toEqual(['']);
    });

  for (const swap of [false, true])
    test(`a node moved into another moved node is rescued once, inside it (client order ${swap ? 'B, A' : 'A, B'})`, async () => {
      const { a, b, sync } = await replicas(tableDeck('B\nZ'), swap);
      const part = slide1(a);
      const [top, bottom] = all(root(a), 'tc');
      const topBody = kids(top!).find((c) => c.local === 'txBody')!;
      const [first, second] = all(bottom!, 'p');
      const run = kids(second!).find((c) => c.local === 'r')!;
      // The run moved once before, so the room lists it before the paragraph it goes into.
      transact(a, [{ op: 'moveNode', part, node: run.id, parent: first!.id, index: 1 }]);
      transact(a, [{ op: 'moveNode', part, node: run.id, parent: second!.id, index: 0 }]);
      sync.pause();
      transact(a, [
        {
          op: 'moveNode',
          part,
          node: first!.id,
          parent: topBody.id,
          index: topBody.children.length,
        },
      ]);
      transact(a, [
        { op: 'moveNode', part, node: run.id, parent: first!.id, index: first!.children.length },
      ]);
      transact(b, [{ op: 'removeNode', part: slide1(b), node: firstRow(b).id }]);
      sync.flush();
      const text = converged(a, b);
      expect(text).toEqual(['BZ', '']);
    });

  test('paragraphs moved one by one come back in their order', async () => {
    const { a, b, sync } = await replicas(tableDeck('B1\nB2\nB3'));
    const part = slide1(a);
    sync.pause();
    for (let i = 0; i < 3; i += 1) {
      const [top, bottom] = all(root(a), 'tc');
      const topBody = kids(top!).find((c) => c.local === 'txBody')!;
      const first = all(bottom!, 'p')[0]!;
      transact(a, [
        {
          op: 'moveNode',
          part,
          node: first.id,
          parent: topBody.id,
          index: topBody.children.length,
        },
      ]);
    }
    transact(b, [{ op: 'removeNode', part: slide1(b), node: firstRow(b).id }]);
    sync.flush();
    expect(converged(a, b)).toEqual(['B1', 'B2', 'B3']);
  });

  test('paragraphs that one write moves come back in their order, in place of the new empty paragraph of the cell', async () => {
    const { a, b, sync } = await replicas(tableDeck('B1\nB2'));
    const part = slide1(a);
    const [top, bottom] = all(root(a), 'tc');
    const topBody = kids(top!).find((c) => c.local === 'txBody')!;
    const bottomBody = kids(bottom!).find((c) => c.local === 'txBody')!;
    const [p1, p2] = all(bottom!, 'p');
    sync.pause();
    // One batch, as the table merge writes it: both paragraphs move up, and the cell keeps an
    // empty paragraph.
    transact(a, [
      { op: 'moveNode', part, node: p1!.id, parent: topBody.id, index: topBody.children.length },
      {
        op: 'moveNode',
        part,
        node: p2!.id,
        parent: topBody.id,
        index: topBody.children.length + 1,
      },
      {
        op: 'insertNode',
        part,
        parent: bottomBody.id,
        index: bottomBody.children.length - 2,
        spec: { ns: A_NS, local: 'p', prefix: 'a' },
      },
    ]);
    transact(b, [{ op: 'removeNode', part: slide1(b), node: firstRow(b).id }]);
    sync.flush();
    expect(converged(a, b)).toEqual(['B1', 'B2']);
  });

  // A merge, a delete of the source row, and a peer delete of the target row: the moved text
  // goes with the rows (`yjs-table-row-deletes.test.ts`).

  test('a node that its mover deleted after the move stays deleted', async () => {
    const { a, b, sync } = await replicas(tableDeck());
    sync.pause();
    transact(a, mergeOps(a));
    const moved = all(all(root(a), 'tc')[0]!, 'p').at(-1)!;
    transact(a, [{ op: 'removeNode', part: slide1(a), node: moved.id }]);
    transact(b, [{ op: 'removeNode', part: slide1(b), node: firstRow(b).id }]);
    sync.flush();
    expect(converged(a, b)).toEqual(['']);
  });

  test('undo of the row delete brings the row back with its text, and the rescue leaves it', async () => {
    const { a, b, sync } = await replicas(tableDeck());
    sync.pause();
    transact(a, mergeOps(a));
    transact(b, [{ op: 'removeNode', part: slide1(b), node: firstRow(b).id }]);
    sync.flush();
    expect(paragraphs(a)).toEqual(['B']);
    expect(b.exec({ type: 'undo' }).ok).toBe(true);
    // The row is back with "A" and the merged "B", and the lower cell keeps its empty paragraph.
    expect(converged(a, b)).toEqual(['A', 'B', '']);
  });
});
