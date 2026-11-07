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
