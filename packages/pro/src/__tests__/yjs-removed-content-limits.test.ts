/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// The limits of the read that puts back removed runs and paragraphs (`removed-content.ts`):
// every replica puts back the same nodes, whatever order its room fields arrived in, past the
// count limits, and a hostile room cannot make one read take more than its step budget.

import { describe, expect, test } from 'bun:test';
import * as Y from 'yjs';
import type { PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import { buildPptx, para, run, shape } from '../../../../scripts/fixtures/builder.ts';
import { F, roomOf, type Room } from '../collaboration/yjs/schema.ts';
import { IDENTITY } from '../collaboration/yjs/text-identity.ts';
import { withRemovedRuns } from '../collaboration/yjs/removed-content.ts';
import { openFuzzPeers } from './fuzz-support.ts';
import { IN, converged, selectIn } from './lineage-support.ts';

const NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const PART = '/ppt/slides/slide1.xml';
const END = '<a:endParaRPr lang="en-US"/>';
const three = () =>
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
          text: para(run('AB') + END) + para(run('CDEF') + END) + para(run('GH') + END),
        }),
      },
    ],
  });
const ok = (r: { ok: boolean }) => expect(r.ok).toBe(true);

function selectAcross(e: PptxEditorInstance, p: number, from: number, q: number, to: number) {
  e.setSelection({
    kind: 'text',
    nodeId: e.session!.layout(0)!.items[0]!.element.nodeId,
    anchor: { paragraph: p, offset: from },
    focus: { paragraph: q, offset: to },
  });
}

/** Press Enter at the end of the last paragraph and join it again, `n` times. */
function enterAndJoin(e: PptxEditorInstance, n: number) {
  for (let i = 0; i < n; i += 1) {
    selectIn(e, 2, 2);
    ok(e.input('paragraph'));
    ok(e.input('deleteBackward'));
  }
}

describe('removed paragraphs past the count limits', () => {
  test('replicas whose room fields arrived in different orders show the same text', async () => {
    const { editors, relay } = await openFuzzPeers(three(), [1_001, 2_002, 3_003]);
    const [a, b] = editors as [PptxEditorInstance, PptxEditorInstance];
    // Peer a removes 1000 paragraphs before peer b sees any, and peer b removes 100, so the
    // body's fields arrive in another order at each peer.
    enterAndJoin(a, 1_000);
    selectAcross(a, 0, 1, 2, 1);
    ok(a.input('deleteBackward'));
    enterAndJoin(b, 100);
    selectIn(b, 1, 2);
    ok(b.input('text', 'x'));
    relay.flush();
    expect(converged(...editors)).toEqual(['AH', 'x']);
  }, 60_000);
});

/** A record in `room` with a shell, an optional child list, and optional text. */
function record(room: Room, key: string, local: string, children: string[] = [], text?: string) {
  const rec = new Y.Map<unknown>();
  room.nodes.set(key, rec);
  rec.set(F.part, PART);
  rec.set(F.key, key);
  rec.set(F.shell, { ns: NS_A, local });
  const list = new Y.Array<string>();
  list.push(children);
  rec.set(F.children, list);
  if (text !== undefined) rec.set(F.text, new Y.Text(text));
  return rec;
}

describe('a hostile room', () => {
  test('one read of removed paragraphs stays within its step budget', () => {
    const doc = new Y.Doc();
    const room = roomOf(doc);
    let body!: Y.Map<unknown>;
    doc.transact(() => {
      body = record(room, 'body', 'txBody');
      // 1024 empty runs, each with an a:t of 1024 empty text records.
      const xs = Array.from({ length: 1_024 }, (_, i) => `x${i}`);
      for (const x of xs) record(room, x, 't', [], '');
      const runs = Array.from({ length: 1_024 }, (_, i) => `r${i}`);
      for (const r of runs) {
        record(room, r, 'r', [`t${r}`]).set('rw.7', []);
        record(room, `t${r}`, 't', xs);
      }
      // 1024 removed paragraphs, each listing every run twice and holding every run's place.
      for (let i = 0; i < 1_024; i += 1) {
        const p = record(room, `p${i}`, 'p', [...runs, ...runs]);
        p.set('rw.7', []);
        for (const r of runs) p.set(`${IDENTITY.removedRun}${r}`, '');
        body.set(`${IDENTITY.removedRun}p${i}`, '');
      }
    });
    let reads = 0;
    const started = performance.now();
    const out = withRemovedRuns(room, PART, body, [], () => {
      reads += 1;
      return null;
    });
    expect(out).toEqual([]);
    expect(reads).toBe(0);
    expect(performance.now() - started).toBeLessThan(5_000);
  });

  test('the same fields in another order put back the same paragraphs', () => {
    const build = (reverse: boolean) => {
      const doc = new Y.Doc();
      const room = roomOf(doc);
      let body!: Y.Map<unknown>;
      doc.transact(() => {
        body = record(room, 'body', 'txBody');
        const keys = Array.from({ length: 1_100 }, (_, i) => `p${String(i).padStart(4, '0')}`);
        for (const k of reverse ? [...keys].reverse() : keys) {
          record(room, `${k}x`, 'x', [], 'z');
          record(room, `${k}t`, 't', [`${k}x`]);
          record(room, `${k}r`, 'r', [`${k}t`]);
          record(room, k, 'p', [`${k}r`]).set('rw.7', []);
          body.set(`${IDENTITY.removedRun}${k}`, '');
        }
      });
      const read = (key: string) => ({
        key,
        ns: NS_A,
        local: 'p',
        children: [
          {
            key: `${key}r`,
            ns: NS_A,
            local: 'r',
            children: [
              { key: `${key}t`, ns: NS_A, local: 't', children: [{ key: `${key}x`, text: 'z' }] },
            ],
          },
        ],
      });
      return withRemovedRuns(room, PART, body, [], read).map((s) => s.key);
    };
    const forward = build(false);
    expect(forward.length).toBeGreaterThan(0);
    expect(build(true)).toEqual(forward);
  });
});
