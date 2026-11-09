/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// An undo or a redo changes only what its own edit wrote (`undo-moves.ts`, `unlistEmptied` in
// `undo-unlisted.ts`). Three editors on linked Y.Docs open one text box with the paragraphs
// "ABCDEFGH" (runs "ABCD", "EF", "GH") and "IJKL". Each case runs in several client-id orders,
// delivers the updates in stages, and checks the text after the undo and after the redo.

import { describe, expect, test } from 'bun:test';
import type { PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import { buildPptx, para, run, shape } from '../../../../scripts/fixtures/builder.ts';
import { converged, IN, paragraphs, peers, selectIn } from './lineage-support.ts';

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

const ORDERS: readonly (readonly number[])[] = [
  [1_001, 2_002, 3_003],
  [3_003, 2_002, 1_001],
  [2_002, 3_003, 1_001],
];
const ok = (result: { ok: boolean }) => expect(result.ok).toBe(true);
type Three = [PptxEditorInstance, PptxEditorInstance, PptxEditorInstance];

describe('an undo of Enter keeps the text a peer typed into the new paragraph', () => {
  // The undo takes the Enter's paragraph out, and the paragraph shows again with the peer's
  // text only (`undo-unlisted.ts`). No replica moves the peer's text: a move would not reach a
  // delete of it that is in flight.
  for (const ids of ORDERS)
    test(`typed at the start of the new paragraph (client ids ${ids.join(', ')})`, async () => {
      const { editors, sync, docs } = await peers(deck(), ids);
      const [a, b, c] = editors as Three;
      sync.pause();
      selectIn(c, 0, 5);
      ok(c.input('paragraph'));
      sync.deliverTo(docs[1]!);
      selectIn(b, 1, 0);
      ok(b.input('text', 'x'));
      sync.deliverTo(docs[2]!);
      expect(paragraphs(c)).toEqual(['ABCDE', 'xFGH', 'IJKL']);
      ok(c.exec({ type: 'undo' }));
      expect(paragraphs(c)).toEqual(['ABCDEFGH', 'x', 'IJKL']);
      sync.flush();
      expect(converged(a, b, c)).toEqual(['ABCDEFGH', 'x', 'IJKL']);
      ok(c.exec({ type: 'redo' }));
      expect(converged(a, b, c)).toEqual(['ABCDE', 'xFGH', 'IJKL']);
    });

  // The peer's join lists the new paragraph's run again, so the undo takes only the run's text
  // that the Enter wrote: the run stays for the peer's text.
  for (const ids of ORDERS)
    test(`typed into it, then joined back (client ids ${ids.join(', ')})`, async () => {
      const { editors, sync } = await peers(deck(), ids);
      const [a, b, c] = editors as Three;
      selectIn(c, 0, 5);
      ok(c.input('paragraph'));
      sync.pause();
      selectIn(b, 1, 1);
      ok(b.input('text', 'x'));
      selectIn(b, 1, 0);
      ok(b.input('deleteBackward'));
      sync.flush();
      expect(converged(a, b, c)).toEqual(['ABCDEFxGH', 'IJKL']);
      ok(c.exec({ type: 'undo' }));
      expect(paragraphs(c)).toEqual(['ABCDEFxGH', 'IJKL']);
      expect(converged(a, b, c)).toEqual(['ABCDEFxGH', 'IJKL']);
      ok(c.exec({ type: 'redo' }));
      expect(converged(a, b, c).join('')).toBe('ABCDEFxGHIJKL');
    });
});

describe('an undo of a format keeps text that a peer deleted deleted', () => {
  // The peer typed over the formatted text before the undo: its delete reached the format's
  // copy, and the undo does not put the old characters back.
  for (const ids of ORDERS)
    test(`typed over before the undo (client ids ${ids.join(', ')})`, async () => {
      const { editors, sync } = await peers(deck(), ids);
      const [a, b, c] = editors as Three;
      sync.pause();
      selectIn(a, 1, 2, 4);
      ok(a.exec({ type: 'toggleBold' }));
      selectIn(b, 1, 1, 4);
      ok(b.input('text', 'x'));
      sync.flush();
      expect(converged(a, b, c)).toEqual(['ABCDEFGH', 'Ix']);
      ok(a.exec({ type: 'undo' }));
      expect(paragraphs(a)).toEqual(['ABCDEFGH', 'Ix']);
      expect(converged(a, b, c)).toEqual(['ABCDEFGH', 'Ix']);
      // The peer took all of the formatted text, so the redo may have nothing left to change.
      a.exec({ type: 'redo' });
      expect(paragraphs(a)).toEqual(['ABCDEFGH', 'Ix']);
      expect(converged(a, b, c)).toEqual(['ABCDEFGH', 'Ix']);
    });

  // The peer deletes at the same time as the undo: the replica that deleted the text deletes
  // what the undo put back for it, when it arrives. Each stage order ends alike.
  for (const ids of ORDERS)
    for (const stage of ['undo first', 'delete first', 'at once'] as const)
      test(`deleted at the same time, ${stage} (client ids ${ids.join(', ')})`, async () => {
        const { editors, sync, docs } = await peers(deck(), ids);
        const [a, b, c] = editors as Three;
        sync.pause();
        selectIn(c, 0, 2, 5);
        ok(c.exec({ type: 'toggleItalic' }));
        ok(c.exec({ type: 'undo' }));
        if (stage === 'undo first') sync.deliverTo(docs[1]!);
        selectIn(b, 0, 2, 4);
        ok(b.input('deleteBackward'));
        if (stage === 'delete first') sync.deliverTo(docs[2]!);
        sync.flush();
        expect(converged(a, b, c)).toEqual(['ABEFGH', 'IJKL']);
        ok(c.exec({ type: 'redo' }));
        expect(converged(a, b, c)).toEqual(['ABEFGH', 'IJKL']);
        ok(c.exec({ type: 'undo' }));
        expect(converged(a, b, c)).toEqual(['ABEFGH', 'IJKL']);
      });

  // The redo puts the format's copies back as new characters: the replica that deleted the text
  // deletes those too.
  for (const ids of ORDERS)
    test(`deleted while the format, its undo, and its redo are in flight (client ids ${ids.join(', ')})`, async () => {
      const { editors, sync } = await peers(deck(), ids);
      const [a, b, c] = editors as Three;
      sync.pause();
      selectIn(c, 0, 2, 5);
      ok(c.exec({ type: 'toggleItalic' }));
      ok(c.exec({ type: 'undo' }));
      ok(c.exec({ type: 'redo' }));
      selectIn(b, 0, 2, 4);
      ok(b.input('deleteBackward'));
      sync.flush();
      expect(converged(a, b, c)).toEqual(['ABEFGH', 'IJKL']);
      ok(c.exec({ type: 'undo' }));
      expect(converged(a, b, c)).toEqual(['ABEFGH', 'IJKL']);
    });

  // A delete that the undoing peer had not seen yet, and a plain undo of a delete, still put
  // their text back.
  for (const ids of ORDERS)
    test(`the peer's own undo of its delete brings the text back (client ids ${ids.join(', ')})`, async () => {
      const { editors, sync } = await peers(deck(), ids);
      const [a, b, c] = editors as Three;
      sync.pause();
      selectIn(c, 0, 2, 5);
      ok(c.exec({ type: 'toggleItalic' }));
      selectIn(b, 0, 2, 4);
      ok(b.input('deleteBackward'));
      sync.flush();
      ok(c.exec({ type: 'undo' }));
      expect(converged(a, b, c)).toEqual(['ABEFGH', 'IJKL']);
      ok(b.exec({ type: 'undo' }));
      expect(converged(a, b, c)).toEqual(['ABCDEFGH', 'IJKL']);
    });
});

describe('an undo of Enter after an undo of a join keeps text a peer typed', () => {
  // The undo of the join lists the Enter's paragraph again with a new entry. The undo of the
  // Enter deletes that entry, and the paragraph shows again with the peer's text.
  for (const ids of ORDERS)
    test(`typed over part of the new paragraph (client ids ${ids.join(', ')})`, async () => {
      const { editors } = await peers(deck(), ids);
      const [a, , c] = editors as Three;
      selectIn(a, 0, 2);
      ok(a.input('paragraph'));
      selectIn(a, 1, 0);
      ok(a.input('deleteBackward'));
      ok(a.exec({ type: 'undo' }));
      expect(converged(a, c)).toEqual(['AB', 'CDEFGH', 'IJKL']);
      selectIn(c, 1, 1, 4);
      ok(c.input('text', 'z'));
      expect(converged(a, c)).toEqual(['AB', 'CzGH', 'IJKL']);
      ok(a.exec({ type: 'undo' }));
      expect(paragraphs(a)).toEqual(['ABCGH', 'z', 'IJKL']);
      expect(converged(a, c)).toEqual(['ABCGH', 'z', 'IJKL']);
    });
});

describe('a redo of a format keeps text that a peer deleted deleted', () => {
  // The peer deleted part of the formatted text before the undo. The undo puts back only the
  // rest, and the redo does not put the deleted part back.
  for (const ids of ORDERS)
    test(`deleted before the undo (client ids ${ids.join(', ')})`, async () => {
      const { editors, sync } = await peers(deck(), ids);
      const [a, , c] = editors as Three;
      sync.pause();
      selectIn(a, 1, 1, 3);
      ok(a.input('deleteBackward'));
      selectIn(c, 1, 2, 4);
      ok(c.exec({ type: 'toggleBold' }));
      sync.flush();
      expect(converged(...editors)).toEqual(['ABCDEFGH', 'IL']);
      ok(c.exec({ type: 'undo' }));
      expect(converged(...editors)).toEqual(['ABCDEFGH', 'IL']);
      c.exec({ type: 'redo' });
      expect(paragraphs(c)).toEqual(['ABCDEFGH', 'IL']);
      expect(converged(...editors)).toEqual(['ABCDEFGH', 'IL']);
    });
});

describe("an undo's delete reaches a peer's copy made at the same time", () => {
  // The peer's Enter copies the typed text into a new paragraph while the typist undoes it. The
  // typist's undo deletes the text, and its delete reaches the copy when the copy arrives.
  for (const ids of ORDERS)
    for (const stage of ['undo first', 'Enter first', 'at once'] as const)
      test(`an undo of typing and an Enter inside the text, ${stage} (client ids ${ids.join(', ')})`, async () => {
        const { editors, sync, docs } = await peers(deck(), ids);
        const [a, b, c] = editors as Three;
        selectIn(a, 1, 4);
        ok(a.input('text', 'x'));
        ok(a.input('text', 'y'));
        expect(converged(a, b, c)).toEqual(['ABCDEFGH', 'IJKLxy']);
        sync.pause();
        ok(a.exec({ type: 'undo' }));
        if (stage === 'undo first') sync.deliverTo(docs[1]!);
        selectIn(b, 1, 5);
        ok(b.input('paragraph'));
        if (stage === 'Enter first') sync.deliverTo(docs[0]!);
        sync.flush();
        expect(converged(a, b, c).join('')).toBe('ABCDEFGHIJKL');
        ok(a.exec({ type: 'redo' }));
        expect(converged(a, b, c).join('')).toBe('ABCDEFGHIJKLxy');
      });
});

describe('an undo that names a removed paragraph again keeps the text a peer typed', () => {
  // A redo of a join removes the second paragraph again, and a peer then deletes across the
  // paragraphs. The undo of an Enter names the removed paragraph as the place of its runs. The
  // rescue of removed runs does not count a removed paragraph as a place (`roomListings`).
  for (const ids of ORDERS)
    test(`an undo of Enter after a redo of a join (client ids ${ids.join(', ')})`, async () => {
      const { editors, sync } = await peers(deck(), ids);
      const [a, b, c] = editors as Three;
      sync.pause();
      selectIn(c, 1, 0);
      ok(c.input('deleteBackward'));
      sync.flush();
      sync.pause();
      selectIn(a, 0, 1, 4);
      ok(a.input('paragraph'));
      ok(c.exec({ type: 'undo' }));
      ok(c.exec({ type: 'redo' }));
      selectIn(a, 1, 0);
      ok(a.input('paragraph'));
      sync.flush();
      expect(converged(a, b, c)).toEqual(['AIJKL', '', 'EFGH']);
      sync.pause();
      selectIn(c, 0, 0);
      ok(c.input('text', 'x'));
      const nodeId = c.session!.layout(0)!.items[0]!.element.nodeId;
      c.setSelection({
        kind: 'text',
        nodeId,
        anchor: { paragraph: 0, offset: 2 },
        focus: { paragraph: 1, offset: 0 },
      });
      ok(c.input('deleteBackward'));
      selectIn(b, 0, 2, 3);
      ok(b.input('text', 'y'));
      sync.flush();
      expect(converged(a, b, c)).toEqual(['xAy', 'EFGH']);
      ok(a.exec({ type: 'undo' }));
      expect(paragraphs(a)).toEqual(['xAEFGHy']);
      expect(converged(a, b, c)).toEqual(['xAEFGHy']);
      ok(a.exec({ type: 'redo' }));
      expect(converged(a, b, c).join('')).toBe('xAyEFGH');
    });
});
