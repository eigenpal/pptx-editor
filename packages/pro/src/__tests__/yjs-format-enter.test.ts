/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// A peer formats part of "Hello world" while another peer presses Enter inside it, at the same
// time. Each split copies the text after its cut, and the "Once" repair keeps each character
// in one copy (`text-repair.ts`). A format whose split starts after the cut follows the Enter
// into the new paragraph ("Follow an Enter"), so the text stays in order, and an undo of the
// Enter takes those runs back. Every cut against every format range, in both client orders,
// then undo by either peer or both.

import { describe, expect, test } from 'bun:test';
import type { PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import { buildPptx, para, run, shape } from '../../../../scripts/fixtures/builder.ts';
import { openFuzzPeers } from './fuzz-support.ts';
import { IN, converged, selectIn } from './lineage-support.ts';

const TEXT = 'Hello world';
const deck = (end: boolean) =>
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
          text: para(run(TEXT) + (end ? '<a:endParaRPr lang="en-US"/>' : '')),
        }),
      },
    ],
  });
const ok = (r: { ok: boolean }) => expect(r.ok).toBe(true);
const CLIENT_ORDERS = [
  [1_001, 2_002, 3_003],
  [2_002, 1_001, 3_003],
];
/** The bold ranges, `[from, to)` of "Hello world". */
const FORMATS = [
  [3, 8],
  [1, 4],
  [6, 10],
  [0, 11],
] as const;
/** Who undoes after the merge, in order: peer a formatted, peer b pressed Enter. */
const UNDOS = ['', 'a', 'b', 'ba', 'ab'] as const;

describe('a format and an Enter at once inside the same text', () => {
  for (const ids of CLIENT_ORDERS)
    for (const [from, to] of FORMATS)
      for (const cut of [2, 5, 7, 9])
        for (const end of [false, true])
          for (const undos of UNDOS)
            test(`bold ${from}..${to}, Enter at ${cut}${end ? ' before a paragraph end' : ''}, undo ${undos || 'none'} (client ids ${ids.join(', ')})`, async () => {
              const { editors, relay } = await openFuzzPeers(deck(end), ids);
              const [a, b] = editors as [PptxEditorInstance, PptxEditorInstance];
              selectIn(a, 0, from, to);
              ok(a.exec({ type: 'toggleBold' }));
              selectIn(b, 0, cut);
              ok(b.input('paragraph'));
              relay.flush();
              const split = [TEXT.slice(0, cut), TEXT.slice(cut)];
              expect(converged(...editors)).toEqual(split);
              for (const who of undos) {
                ok((who === 'a' ? a : b).exec({ type: 'undo' }));
                relay.flush();
              }
              const shown = converged(...editors);
              if (undos === 'ab') {
                // An undo of the format, then of the Enter, keeps each character once, but can
                // show the text out of order (a known limit).
                expect([...shown.join('')].sort()).toEqual([...TEXT].sort());
                return;
              }
              expect(shown).toEqual(undos.includes('b') ? [TEXT] : split);
            });
});

// The step that a redo of the Enter pushes keeps the runs that followed it (`carryFollowed` in
// `undo-followed.ts`), so a second undo takes them back too.
describe('an undo, a redo, and an undo again of an Enter that a format followed', () => {
  for (const ids of CLIENT_ORDERS)
    for (const [from, to] of FORMATS)
      for (const cut of [2, 5, 7])
        test(`bold ${from}..${to}, Enter at ${cut} (client ids ${ids.join(', ')})`, async () => {
          const { editors, relay } = await openFuzzPeers(deck(false), ids);
          const [a, b] = editors as [PptxEditorInstance, PptxEditorInstance];
          selectIn(a, 0, from, to);
          ok(a.exec({ type: 'toggleBold' }));
          selectIn(b, 0, cut);
          ok(b.input('paragraph'));
          relay.flush();
          const split = [TEXT.slice(0, cut), TEXT.slice(cut)];
          const steps: ['undo' | 'redo', string[]][] = [
            ['undo', [TEXT]],
            ['redo', split],
            ['undo', [TEXT]],
          ];
          for (const [type, shown] of steps) {
            ok(b.exec({ type }));
            relay.flush();
            expect(converged(...editors)).toEqual(shown);
          }
        });
});

// After the format followed the Enter, the peer that formatted presses Enter inside the
// formatted text: the run moves on. An undo of the first Enter leaves it where it went, so each
// character shows once (`checkFollowed` in `undo-followed.ts`), and a redo cuts the text again.
describe('a peer moves the followed run on before the Enter is undone', () => {
  for (const ids of CLIENT_ORDERS)
    for (const [from, to, cut] of [
      [6, 10, 2],
      [3, 8, 2],
      [6, 10, 5],
    ] as const)
      test(`bold ${from}..${to}, Enter at ${cut}, then Enter at ${to - 1} (client ids ${ids.join(', ')})`, async () => {
        const { editors, relay } = await openFuzzPeers(deck(false), ids);
        const [a, b] = editors as [PptxEditorInstance, PptxEditorInstance];
        selectIn(a, 0, from, to);
        ok(a.exec({ type: 'toggleBold' }));
        selectIn(b, 0, cut);
        ok(b.input('paragraph'));
        relay.flush();
        selectIn(a, 1, to - cut - 1);
        ok(a.input('paragraph'));
        relay.flush();
        const moved = [TEXT.slice(0, cut), TEXT.slice(cut, to - 1), TEXT.slice(to - 1)];
        expect(converged(...editors)).toEqual(moved);
        ok(b.exec({ type: 'undo' }));
        relay.flush();
        expect(converged(...editors)).toEqual([TEXT.slice(0, to - 1), TEXT.slice(to - 1)]);
        ok(b.exec({ type: 'redo' }));
        relay.flush();
        expect(converged(...editors).join('')).toBe(TEXT);
      });
});
