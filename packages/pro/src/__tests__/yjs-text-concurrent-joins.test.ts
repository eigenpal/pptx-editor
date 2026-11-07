/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Two peers press Backspace at the start of two neighboring paragraphs at the same time.
// One join moves the runs of the last paragraph into the middle
// one; the other join moves the runs of the middle paragraph into the first one and removes
// the middle paragraph without having seen the runs that came in. The move rescue (openspec
// change `collaboration-text-lineage`, "Follow a join") sends those runs after the runs that
// the second join kept, so every replica shows all of the text, in every client order and
// every delivery order.

import { describe, expect, test } from 'bun:test';
import type { PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import { canonicalPackageFingerprint, readPptx } from '@pptx-editor.dev/core/store';
import { packageViolations } from '../../../core/src/store/store/part-checks.ts';
import { buildPptx, para, run, shape } from '../../../../scripts/fixtures/builder.ts';
import { openFuzzPeers, type Relay } from './fuzz-support.ts';
import { IN, converged, paragraphs, selectIn } from './lineage-support.ts';

const deck = (texts = ['AB', 'CDEFGH', 'IJKL']) =>
  buildPptx({
    slides: [
      {
        layout: 'blank',
        shapes: shape({
          id: 2,
          x: IN,
          y: IN,
          cx: 6 * IN,
          cy: 2 * IN,
          text: texts.map((t) => para(run(t))).join(''),
        }),
      },
    ],
  });

/** Every order of three client ids, so the tie rules vary. */
const CLIENT_ORDERS = [
  [1_001, 2_002, 3_003],
  [1_001, 3_003, 2_002],
  [2_002, 1_001, 3_003],
  [2_002, 3_003, 1_001],
  [3_003, 1_001, 2_002],
  [3_003, 2_002, 1_001],
];

/** Who receives which join first, before every other update is delivered. */
const DELIVERIES: readonly { name: string; first: readonly [number, number][] }[] = [
  { name: 'all at once', first: [] },
  { name: 'the third peer gets the later join first', first: [[0, 2]] },
  { name: 'the third peer gets the earlier join first', first: [[1, 2]] },
  {
    name: 'each joiner gets the other join, then the third peer',
    first: [
      [1, 0],
      [0, 1],
    ],
  },
];

const order = (ids: readonly number[]) => `client ids ${ids.join(', ')}`;

/** Peer 0 joins "IJKL" into "CDEFGH" while peer 1 joins "CDEFGH" into "AB". */
async function concurrentJoins(ids: readonly number[]) {
  const { editors, relay } = await openFuzzPeers(deck(), ids);
  const [a, b, c] = editors as [PptxEditorInstance, PptxEditorInstance, PptxEditorInstance];
  selectIn(a, 2, 0);
  expect(a.input('deleteBackward').ok).toBe(true);
  expect(paragraphs(a)).toEqual(['AB', 'CDEFGHIJKL']);
  selectIn(b, 1, 0);
  expect(b.input('deleteBackward').ok).toBe(true);
  expect(paragraphs(b)).toEqual(['ABCDEFGH', 'IJKL']);
  return { a, b, c, relay };
}

/**
 * The replicas converged on `expected`, each saved copy reopens with the same canonical
 * fingerprint on every replica, and no saved copy breaks a package rule.
 */
function savedAlike(editors: readonly PptxEditorInstance[], relay: Relay, expected: string[]) {
  relay.flush();
  expect(converged(...editors)).toEqual(expected);
  const prints = editors.map((e) => {
    const r = readPptx(e.save()!);
    if (!r.ok) throw new Error('reopen failed');
    expect(packageViolations(r.pkg)).toEqual([]);
    return canonicalPackageFingerprint(r.pkg);
  });
  expect(new Set(prints).size).toBe(1);
}

describe('two concurrent joins of neighboring paragraphs keep all of the text', () => {
  for (const ids of CLIENT_ORDERS)
    for (const delivery of DELIVERIES)
      test(`${delivery.name} (${order(ids)})`, async () => {
        const { a, b, c, relay } = await concurrentJoins(ids);
        for (const [from, to] of delivery.first) relay.deliver(from, to, Infinity);
        savedAlike([a, b, c], relay, ['ABCDEFGHIJKL']);
      });

  for (const ids of CLIENT_ORDERS)
    test(`three joins of a chain of four paragraphs (${order(ids)})`, async () => {
      const { editors, relay } = await openFuzzPeers(deck(['AB', 'CD', 'EF', 'GH']), ids);
      const [a, b, c] = editors as [PptxEditorInstance, PptxEditorInstance, PptxEditorInstance];
      for (const [editor, paragraph] of [
        [a, 3],
        [b, 2],
        [c, 1],
      ] as const) {
        selectIn(editor, paragraph, 0);
        expect(editor.input('deleteBackward').ok).toBe(true);
      }
      savedAlike([a, b, c], relay, ['ABCDEFGH']);
    });
});

describe('an undo of either join after the merge', () => {
  for (const ids of CLIENT_ORDERS) {
    test(`undo and redo of the later join (${order(ids)})`, async () => {
      const { a, b, c, relay } = await concurrentJoins(ids);
      relay.flush();
      expect(a.exec({ type: 'undo' }).ok).toBe(true);
      savedAlike([a, b, c], relay, ['ABCDEFGH', 'IJKL']);
      expect(a.exec({ type: 'redo' }).ok).toBe(true);
      savedAlike([a, b, c], relay, ['ABCDEFGHIJKL']);
    });

    test(`undo and redo of the earlier join (${order(ids)})`, async () => {
      const { a, b, c, relay } = await concurrentJoins(ids);
      relay.flush();
      expect(b.exec({ type: 'undo' }).ok).toBe(true);
      savedAlike([a, b, c], relay, ['AB', 'CDEFGHIJKL']);
      expect(b.exec({ type: 'redo' }).ok).toBe(true);
      savedAlike([a, b, c], relay, ['ABCDEFGHIJKL']);
    });

    test(`a third peer types into the joined paragraph, then each joiner undoes (${order(ids)})`, async () => {
      const { a, b, c, relay } = await concurrentJoins(ids);
      relay.flush();
      // C's write lists the runs that followed the join in the first paragraph.
      selectIn(c, 0, 12);
      expect(c.input('text', 'x').ok).toBe(true);
      savedAlike([a, b, c], relay, ['ABCDEFGHIJKLx']);
      // A write never lists a node where the move rescue shows it, so B's undo shows "IJKL"
      // with the middle paragraph again.
      expect(b.exec({ type: 'undo' }).ok).toBe(true);
      savedAlike([a, b, c], relay, ['AB', 'CDEFGHIJKLx']);
      expect(a.exec({ type: 'undo' }).ok).toBe(true);
      savedAlike([a, b, c], relay, ['AB', 'CDEFGH', 'IJKLx']);
    });

    test(`both joiners undo at once, then both redo (${order(ids)})`, async () => {
      const { a, b, c, relay } = await concurrentJoins(ids);
      relay.flush();
      expect(a.exec({ type: 'undo' }).ok).toBe(true);
      expect(b.exec({ type: 'undo' }).ok).toBe(true);
      savedAlike([a, b, c], relay, ['AB', 'CDEFGH', 'IJKL']);
      expect(a.exec({ type: 'redo' }).ok).toBe(true);
      expect(b.exec({ type: 'redo' }).ok).toBe(true);
      savedAlike([a, b, c], relay, ['ABCDEFGHIJKL']);
    });
  }

  for (const ids of CLIENT_ORDERS)
    test(`a third peer presses Enter in the rescued text, then the earlier join is undone (${order(ids)})`, async () => {
      const { a, b, c, relay } = await concurrentJoins(ids);
      relay.flush();
      selectIn(c, 0, 10);
      expect(c.input('paragraph').ok).toBe(true);
      savedAlike([a, b, c], relay, ['ABCDEFGHIJ', 'KL']);
      // The undo lists the middle paragraph again at its old place, which the new paragraph
      // took: every character shows once, and "IJ" goes back with "CDEFGH".
      expect(b.exec({ type: 'undo' }).ok).toBe(true);
      savedAlike([a, b, c], relay, ['AB', 'KL', 'CDEFGHIJ']);
      expect(b.exec({ type: 'redo' }).ok).toBe(true);
      savedAlike([a, b, c], relay, ['ABCDEFGHIJ', 'KL']);
    });
});
