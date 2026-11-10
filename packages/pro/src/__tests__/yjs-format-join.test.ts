/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// A peer formats part of "ABCD" while another peer presses Enter inside it and joins the two
// paragraphs again, at the same time. The join lists the runs of the new paragraph again, at the
// place where the format lists its pieces. The pieces go right after the run they came from
// (`piece-place.ts`), and of two copies that start at one character, every replica keeps the
// same one ("Once" in `text-repair.ts`). So the text shows once and in order, in every client
// order and delivery order. A second peer's Enter at the same place keeps each character once.

import { afterEach, describe, expect, test } from 'bun:test';
import type { PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import { buildPptx, para, run, shape } from '../../../../scripts/fixtures/builder.ts';
import { openFuzzPeers, type Relay } from './fuzz-support.ts';
import { IN, converged, formatOf, selectIn } from './lineage-support.ts';

const active: PptxEditorInstance[] = [];
afterEach(() => {
  for (const editor of active.splice(0)) editor.destroy();
});
async function openPeers(ids: readonly number[]) {
  const peers = await openFuzzPeers(deck(), ids);
  active.push(...peers.editors);
  return peers;
}

const deck = () =>
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
          text:
            para(run('ABCD') + run('EF', 'b="1"') + run('GH') + '<a:endParaRPr lang="en-US"/>') +
            para(run('IJKL')),
        }),
      },
    ],
  });
const TEXT = ['ABCDEFGH', 'IJKL'];
const ok = (r: { ok: boolean }) => expect(r.ok).toBe(true);

/** Every order of three client ids, so the tie rules vary. */
const CLIENT_ORDERS = [
  [1_001, 2_002, 3_003],
  [1_001, 3_003, 2_002],
  [2_002, 1_001, 3_003],
  [2_002, 3_003, 1_001],
  [3_003, 1_001, 2_002],
  [3_003, 2_002, 1_001],
];

/** Who receives which peer's changes first (`[from, to]`), before the rest is delivered. */
const DELIVERIES: readonly { name: string; first: readonly [number, number][] }[] = [
  { name: 'all at once', first: [] },
  { name: 'the joiner gets the format first', first: [[2, 0]] },
  { name: 'the formatter gets the join first', first: [[0, 2]] },
  { name: 'the third peer gets the join first', first: [[0, 1]] },
  {
    name: 'the third peer gets the format, then the join',
    first: [
      [2, 1],
      [0, 1],
    ],
  },
];

function deliver(relay: Relay, first: readonly [number, number][]): void {
  for (const [from, to] of first) relay.deliver(from, to, 99);
  relay.flush();
}

/** Peer 2 sets italic on `from`..`to` of "ABCD"; peer 0 presses Enter at `cut` and joins. */
async function formatAndJoin(ids: readonly number[], from: number, to: number, cut: number) {
  const { editors, relay } = await openPeers(ids);
  const [joiner, , formatter] = editors as [
    PptxEditorInstance,
    PptxEditorInstance,
    PptxEditorInstance,
  ];
  selectIn(formatter, 0, from, to);
  ok(formatter.exec({ type: 'toggleItalic' }));
  selectIn(joiner, 0, cut);
  ok(joiner.input('paragraph'));
  selectIn(joiner, 1, 0);
  ok(joiner.input('deleteBackward'));
  return { editors, relay, joiner, formatter };
}

describe('a format and an Enter that the same peer joins again, at once', () => {
  for (const ids of CLIENT_ORDERS)
    for (const [from, to] of [
      [2, 3],
      [1, 3],
      [3, 4],
      [0, 2],
    ] as const)
      for (const cut of [1, 2, 3])
        for (const { name, first } of DELIVERIES)
          test(`italic ${from}..${to}, Enter at ${cut}, ${name} (client ids ${ids.join(', ')})`, async () => {
            const { editors, relay } = await formatAndJoin(ids, from, to, cut);
            deliver(relay, first);
            expect(converged(...editors)).toEqual(TEXT);
            // Text the Enter did not copy keeps the format.
            if (to <= cut)
              for (const e of editors) expect(formatOf(e, 'ABCD'.slice(from, to)).i).toBe('1');
          });
});

describe('an undo after a format and a joined Enter', () => {
  for (const ids of CLIENT_ORDERS) {
    test(`the formatter undoes the format (client ids ${ids.join(', ')})`, async () => {
      const { editors, relay, formatter } = await formatAndJoin(ids, 2, 3, 3);
      relay.flush();
      ok(formatter.exec({ type: 'undo' }));
      relay.flush();
      expect(converged(...editors)).toEqual(TEXT);
      for (const e of editors) expect(formatOf(e, 'ABCD').i).toBeUndefined();
    });
    test(`the joiner undoes the join, then the Enter (client ids ${ids.join(', ')})`, async () => {
      const { editors, relay, joiner } = await formatAndJoin(ids, 2, 3, 3);
      relay.flush();
      ok(joiner.exec({ type: 'undo' }));
      relay.flush();
      expect(converged(...editors)).toEqual(['ABC', 'DEFGH', 'IJKL']);
      ok(joiner.exec({ type: 'undo' }));
      relay.flush();
      expect(converged(...editors)).toEqual(TEXT);
    });
  }
});

// Peer 2 types "X" into "ABCD", presses Enter after it, and joins again. Peer 1, which saw only
// the typing, presses Enter at the same place. Both Enters copy "CD"; every replica keeps one
// copy, so each character shows once. Which copy keeps "CD" depends on the client ids, so the
// text after the cut can show in the second paragraph.
const ENTER_DELIVERIES: readonly { name: string; first: readonly [number, number][] }[] = [
  { name: 'all at once', first: [] },
  { name: 'the joiner gets the other Enter first', first: [[1, 2]] },
  { name: 'the other Enter gets the join first', first: [[2, 1]] },
  { name: 'the third peer gets the join first', first: [[2, 0]] },
  { name: 'the third peer gets the other Enter first', first: [[1, 0]] },
];

describe('two Enters at one place, one of them joined again', () => {
  for (const ids of CLIENT_ORDERS)
    for (const { name, first } of ENTER_DELIVERIES)
      test(`${name} (client ids ${ids.join(', ')})`, async () => {
        const { editors, relay } = await openPeers(ids);
        const [, other, joiner] = editors as [
          PptxEditorInstance,
          PptxEditorInstance,
          PptxEditorInstance,
        ];
        selectIn(joiner, 0, 2);
        ok(joiner.input('text', 'X'));
        selectIn(joiner, 0, 3);
        ok(joiner.input('paragraph'));
        selectIn(joiner, 1, 0);
        ok(joiner.input('deleteBackward'));
        relay.deliver(2, 1, 1);
        selectIn(other, 0, 3);
        ok(other.input('paragraph'));
        deliver(relay, first);
        const shown = converged(...editors);
        expect([...shown.join('')].sort()).toEqual([...'ABXCDEFGHIJKL'].sort());
        expect(shown[0]!.startsWith('ABX')).toBe(true);
      });
});
