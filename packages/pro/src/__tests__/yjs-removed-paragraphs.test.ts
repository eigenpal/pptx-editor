/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// A peer removes a paragraph (a delete across paragraphs, a join) while another peer types
// into it. The removal deletes the characters its writer saw; what the peer typed at the same
// time stays, and the paragraph shows again at its old place with that text only
// (`withRemovedRuns` in `removed-content.ts`). Every client order, and with three peers every
// delivery order of the two edits.

import { describe, expect, test } from 'bun:test';
import type { PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import { canonicalPackageFingerprint, readPptx } from '@pptx-editor.dev/core/store';
import { packageViolations } from '../../../core/src/store/store/part-checks.ts';
import { buildPptx, para, run, shape } from '../../../../scripts/fixtures/builder.ts';
import { openFuzzPeers, type Relay } from './fuzz-support.ts';
import { IN, converged, paragraphs, selectIn } from './lineage-support.ts';

const deck = (body: string) =>
  buildPptx({
    slides: [
      {
        layout: 'blank',
        shapes: shape({ id: 2, x: IN, y: IN, cx: 6 * IN, cy: 2 * IN, text: body }),
      },
    ],
  });
const END = '<a:endParaRPr lang="en-US"/>';
const THREE = para(run('AB') + END) + para(run('CDEF') + END) + para(run('GH') + END);
const three = () => deck(THREE);
const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
const P14 = 'http://schemas.microsoft.com/office/powerpoint/2010/main';
const box = (text: string) => shape({ id: 3, name: 'Box', x: IN, y: IN, cx: 3 * IN, cy: IN, text });

const CLIENT_ORDERS = [
  [1_001, 2_002, 3_003],
  [2_002, 1_001, 3_003],
  [3_003, 2_002, 1_001],
];
const order = (ids: readonly number[]) => `client ids ${ids.join(', ')}`;
const ok = (r: { ok: boolean }) => expect(r.ok).toBe(true);

/** Who gets which edit first: peer 0 removes, peer 1 types, peer 2 watches. */
const DELIVERIES: readonly { name: string; first: readonly [number, number][] }[] = [
  { name: 'all at once', first: [] },
  { name: 'the watcher gets the removal first', first: [[0, 2]] },
  { name: 'the watcher gets the typing first', first: [[1, 2]] },
  {
    name: 'each writer gets the other edit first',
    first: [
      [1, 0],
      [0, 1],
    ],
  },
];

function savedAlike(editors: readonly PptxEditorInstance[], relay: Relay): string[] {
  relay.flush();
  const text = converged(...editors);
  const prints = editors.map((e) => {
    const r = readPptx(e.save()!);
    if (!r.ok) throw new Error('reopen failed');
    expect(packageViolations(r.pkg)).toEqual([]);
    return canonicalPackageFingerprint(r.pkg);
  });
  expect(new Set(prints).size).toBe(1);
  return text;
}

/** A selection from `from` in `p` to `to` in paragraph `q`. */
function selectAcross(e: PptxEditorInstance, p: number, from: number, q: number, to: number) {
  e.setSelection({
    kind: 'text',
    nodeId: e.session!.layout(0)!.items[0]!.element.nodeId,
    anchor: { paragraph: p, offset: from },
    focus: { paragraph: q, offset: to },
  });
}

interface Case {
  readonly name: string;
  readonly deck: () => Uint8Array;
  readonly remove: (e: PptxEditorInstance) => void;
  readonly type: (e: PptxEditorInstance) => void;
  readonly expected: readonly string[];
  /** After the removing peer's undo. */
  readonly undone?: readonly string[];
}

const CASES: readonly Case[] = [
  {
    name: 'a delete across three paragraphs keeps text typed into the middle one',
    deck: three,
    remove: (e) => {
      selectAcross(e, 0, 1, 2, 1);
      ok(e.input('deleteBackward'));
    },
    type: (e) => {
      selectIn(e, 1, 2);
      ok(e.input('text', 'x'));
    },
    expected: ['AH', 'x'],
    undone: ['AB', 'CDxEF', 'GH'],
  },
  {
    name: 'a delete of a whole paragraph keeps text typed at its end',
    deck: three,
    remove: (e) => {
      selectAcross(e, 0, 2, 2, 0);
      ok(e.input('deleteBackward'));
    },
    type: (e) => {
      selectIn(e, 1, 4);
      ok(e.input('text', 'x'));
    },
    expected: ['ABGH', 'x'],
    undone: ['AB', 'CDEFx', 'GH'],
  },
  {
    name: 'a delete that ends in a paragraph keeps text typed into its deleted run',
    deck: () => deck(para(run('AB') + END) + para(run('CD') + run('EF', 'b="1"') + END)),
    remove: (e) => {
      selectAcross(e, 0, 1, 1, 2);
      ok(e.input('deleteBackward'));
    },
    type: (e) => {
      selectIn(e, 1, 1);
      ok(e.input('text', 'x'));
    },
    expected: ['AEF', 'x'],
    undone: ['AB', 'CxDEF'],
  },
  {
    name: 'a join of an empty paragraph keeps text typed into it',
    deck: () => deck(para(run('AB') + END) + para(END) + para(run('GH') + END)),
    remove: (e) => {
      selectIn(e, 1, 0);
      ok(e.input('deleteBackward'));
    },
    type: (e) => {
      selectIn(e, 1, 0);
      ok(e.input('text', 'x'));
    },
    expected: ['AB', 'x', 'GH'],
    undone: ['AB', 'x', 'GH'],
  },
  {
    name: 'a delete of an empty paragraph keeps text typed into it',
    deck: () => deck(para(run('AB') + END) + para(END) + para(run('GH') + END)),
    remove: (e) => {
      selectAcross(e, 0, 2, 2, 0);
      ok(e.input('deleteBackward'));
    },
    type: (e) => {
      selectIn(e, 1, 0);
      ok(e.input('text', 'x'));
    },
    expected: ['ABGH', 'x'],
    undone: ['AB', 'x', 'GH'],
  },
  {
    name: 'a join keeps text typed into the joined paragraph in place',
    deck: three,
    remove: (e) => {
      selectIn(e, 1, 0);
      ok(e.input('deleteBackward'));
    },
    type: (e) => {
      selectIn(e, 1, 2);
      ok(e.input('text', 'x'));
    },
    expected: ['ABCDxEF', 'GH'],
    undone: ['AB', 'CDxEF', 'GH'],
  },
  {
    name: 'a delete across paragraphs keeps an Enter typed into the removed one',
    deck: three,
    remove: (e) => {
      selectAcross(e, 0, 1, 2, 1);
      ok(e.input('deleteBackward'));
    },
    type: (e) => {
      selectIn(e, 1, 2);
      ok(e.input('text', 'x'));
      ok(e.input('paragraph'));
      ok(e.input('text', 'y'));
    },
    expected: ['AH', 'x', 'y'],
  },
  {
    // The drawn copy of a shape in `mc:AlternateContent` (ECMA-376 Part 3, 10.2.1). An edit
    // removes the stale `mc:Fallback` copy, and an undo of the edit puts it back.
    name: 'a delete across paragraphs in a shape in mc:AlternateContent keeps typed text',
    deck: () =>
      buildPptx({
        slides: [
          {
            layout: 'blank',
            shapes:
              `<mc:AlternateContent xmlns:mc="${MC}">` +
              `<mc:Choice xmlns:p14="${P14}" Requires="p14">${box(THREE)}</mc:Choice>` +
              `<mc:Fallback>${box(para(run('stale')))}</mc:Fallback></mc:AlternateContent>`,
          },
        ],
      }),
    remove: (e) => {
      selectAcross(e, 0, 1, 2, 1);
      ok(e.input('deleteBackward'));
    },
    type: (e) => {
      selectIn(e, 1, 2);
      ok(e.input('text', 'x'));
    },
    expected: ['AH', 'x'],
    undone: ['AB', 'CDxEF', 'GH', 'stale'],
  },
];

describe('notes text that a peer removes at the same time stays', () => {
  const notes = (e: PptxEditorInstance) => e.snapshot().notes;
  for (const ids of CLIENT_ORDERS)
    test(`a notes edit that removes paragraphs keeps text typed into them (${order(ids)})`, async () => {
      const bytes = buildPptx({
        slides: [{ layout: 'blank', shapes: box(para(run('T'))), notesXml: THREE }],
      });
      const { editors, relay } = await openFuzzPeers(bytes, ids);
      const [a, b] = editors as [PptxEditorInstance, PptxEditorInstance];
      ok(a.exec({ type: 'setNotes', index: 0, text: 'AH' }));
      ok(b.exec({ type: 'setNotes', index: 0, text: 'AB\nCDxEF\nGH' }));
      savedAlike(editors, relay);
      expect(editors.map(notes)).toEqual(editors.map(() => 'AH\nx'));
      ok(a.exec({ type: 'undo' }));
      savedAlike(editors, relay);
      expect(editors.map(notes)).toEqual(editors.map(() => 'AB\nCDxEF\nGH'));
    });
});

describe('text typed into a paragraph that a peer removes at the same time stays', () => {
  for (const c of CASES)
    for (const ids of CLIENT_ORDERS)
      for (const delivery of DELIVERIES)
        test(`${c.name}: ${delivery.name} (${order(ids)})`, async () => {
          const { editors, relay } = await openFuzzPeers(c.deck(), ids);
          const [a, b] = editors as [PptxEditorInstance, PptxEditorInstance];
          c.remove(a);
          c.type(b);
          for (const [from, to] of delivery.first) relay.deliver(from, to, Infinity);
          expect(savedAlike(editors, relay)).toEqual([...c.expected]);
          if (!c.undone) return;
          ok(a.exec({ type: 'undo' }));
          expect(savedAlike(editors, relay)).toEqual([...c.undone]);
          ok(a.exec({ type: 'redo' }));
          expect(savedAlike(editors, relay)).toEqual([...c.expected]);
          ok(b.exec({ type: 'undo' }));
          const left = savedAlike(editors, relay);
          expect(left.join('')).not.toContain('x');
          expect(paragraphs(a)).toEqual(left);
        });
});
