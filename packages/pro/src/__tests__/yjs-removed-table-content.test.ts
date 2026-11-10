/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
import { expect, test } from 'bun:test';
import * as Y from 'yjs';
import type { ReplicaElementSpec } from '@pptx-editor.dev/core/store';
import { roomOf, F } from '../collaboration/yjs/schema.ts';
import { IDENTITY, REMOVAL_CLOCK_PREFIX } from '../collaboration/yjs/text-identity.ts';
import { withRemovedTableText } from '../collaboration/yjs/removed-table-content.ts';

const ns = 'http://schemas.openxmlformats.org/drawingml/2006/main';
function fixture(known: boolean, bodiless = false) {
  const doc = new Y.Doc();
  const room = roomOf(doc);
  const node = (key: string, local: string, children: string[] = []) => {
    const record = new Y.Map<unknown>();
    room.nodes.set(key, record);
    record.set(F.key, key);
    record.set(F.part, 'slide');
    record.set(F.shell, { ns, local });
    const list = new Y.Array<string>();
    record.set(F.children, list);
    list.push(children);
    return record;
  };
  const table = node('table', 'tbl', ['row']);
  node('row', 'tr', ['cell']);
  node('cell', 'tc', bodiless ? [] : ['body']);
  node('body', 'txBody');
  const removed = node('removed', 'tr', ['oldCell']);
  node('oldCell', 'tc', ['oldBody']);
  node('oldBody', 'txBody', ['paragraph']);
  node('paragraph', 'p');
  const record = node('text', 't');
  const text = new Y.Text();
  record.set(F.text, text);
  if (known) text.insert(0, 'x');
  const clock = Y.encodeStateVector(doc);
  if (!known) text.insert(0, 'x');
  removed.set(`${REMOVAL_CLOCK_PREFIX}${doc.clientID}`, clock);
  table.set(`${IDENTITY.removedRun}removed`, '');
  const spec = (
    key: string,
    local: string,
    children: ReplicaElementSpec[] = []
  ): ReplicaElementSpec => ({ key, ns, local, children });
  const root = spec('table', 'tbl', [
    spec('row', 'tr', [spec('cell', 'tc', bodiless ? [] : [spec('body', 'txBody')])]),
  ]);
  const paragraph: ReplicaElementSpec = {
    key: 'paragraph',
    ns,
    local: 'p',
    children: [{ key: 'text', text: 'x' }],
  };
  const placed = new Set<string>();
  const shownIn = new Map<string, string>();
  const read = () =>
    withRemovedTableText(
      room,
      root,
      (key) =>
        key === 'oldBody' ? { key, ns, local: 'txBody', children: [paragraph] } : paragraph,
      shownIn,
      placed
    );
  return { doc, removed, root, read, placed, shownIn };
}

test('table recovery admits unseen typing but not text the deletion already observed', () => {
  for (const known of [false, true]) {
    const f = fixture(known);
    try {
      expect(f.read() === f.root).toBe(known);
      expect(f.placed.has('paragraph')).toBe(!known);
      expect(f.shownIn.get('paragraph')).toBe(known ? undefined : 'body');
    } finally {
      f.doc.destroy();
    }
  }
});

test('malformed and oversized deletion clocks leave the read and placement maps unchanged', () => {
  for (const value of ['bad', new Uint8Array([128]), new Uint8Array(1 << 17)]) {
    const f = fixture(false);
    try {
      f.removed.set(`${REMOVAL_CLOCK_PREFIX}${f.doc.clientID}`, value);
      expect(f.read()).toBe(f.root);
      expect(f.placed.size).toBe(0);
      expect(f.shownIn.size).toBe(0);
    } finally {
      f.doc.destroy();
    }
  }
});

test('a destination cell without a body receives the source body with the unseen typing', () => {
  const f = fixture(false, true);
  try {
    const result = f.read();
    expect(result).not.toBe(f.root);
    expect(f.shownIn.get('oldBody')).toBe('cell');
    expect(f.shownIn.get('paragraph')).toBe('oldBody');
    expect(JSON.stringify(result).match(/"text":"x"/g)?.length).toBe(1);
  } finally {
    f.doc.destroy();
  }
});
