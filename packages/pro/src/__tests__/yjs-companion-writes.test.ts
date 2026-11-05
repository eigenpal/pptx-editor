/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Schema singletons under concurrent edits: two peers that both replace one shape's fill or

// Writes into merged property containers after concurrent first copies: a removed
// child leaves every copy, a new child keeps its place, a companion has one parent, and a
// copy's child with no slot stays.

import { describe, expect, test } from 'bun:test';
import * as Y from 'yjs';
import { createPptxEditor, type PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import {
  canonicalPackageFingerprint,
  find,
  serializeXmlPart,
  xmlPartOf,
  type Element,
  type NodeSpec,
} from '@pptx-editor.dev/core/store';
import { buildPptx, para, run, shape, solid } from '../../../../scripts/fixtures/builder.ts';
import { collaborationModule, createYjsCollaborationSession } from '../index.ts';
import { F, NODES_KEY, ROOM_LIMITS } from '../collaboration/yjs/schema.ts';

const IN = 914_400;
const NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const NS_MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
const deck = (): Uint8Array =>
  buildPptx({
    slides: [
      {
        layout: 'title',
        shapes: shape({
          id: 2,
          x: IN,
          y: IN,
          cx: 2 * IN,
          cy: IN,
          fill: solid('4472C4'),
          text: para(run('Hello')),
        }),
      },
    ],
  });

function link(a: Y.Doc, b: Y.Doc) {
  const queue: [Y.Doc, Uint8Array][] = [];
  let paused = false;
  const send = (to: Y.Doc) => (update: Uint8Array, origin: unknown) => {
    if (origin === 'link') return;
    if (paused) queue.push([to, update]);
    else Y.applyUpdate(to, update, 'link');
  };
  a.on('update', send(b));
  b.on('update', send(a));
  return {
    pause: () => (paused = true),
    flush: () => {
      paused = false;
      for (const [to, update] of queue.splice(0)) Y.applyUpdate(to, update, 'link');
    },
  };
}

async function peers(ids: readonly number[]) {
  const docs = ids.map((id) => {
    const d = new Y.Doc();
    d.clientID = id;
    return d;
  });
  const links: ReturnType<typeof link>[] = [];
  for (let i = 0; i < docs.length; i++)
    for (let j = i + 1; j < docs.length; j++) links.push(link(docs[i]!, docs[j]!));
  const bytes = deck();
  const eds: PptxEditorInstance[] = [];
  const sessions: ReturnType<typeof createYjsCollaborationSession>[] = [];
  for (const [i, doc] of docs.entries()) {
    const session = createYjsCollaborationSession({
      doc,
      documentId: 'deck',
      identity: { actorId: `p${i}`, name: `p${i}` },
    });
    const editor = createPptxEditor({
      fixedMeasurer: true,
      substitutes: false,
      modules: [collaborationModule({ session })],
    });
    expect((await editor.load(bytes)).ok).toBe(true);
    eds.push(editor);
    sessions.push(session);
  }
  return {
    eds,
    docs,
    sessions,
    link,
    pause: () => links.forEach((l) => l.pause()),
    flush: () => links.forEach((l) => l.flush()),
  };
}

const fingerprint = (e: PptxEditorInstance) =>
  canonicalPackageFingerprint(e.session!.store.package);
const xml = (e: PptxEditorInstance) =>
  serializeXmlPart(xmlPartOf(e.session!.store.package, e.session!.info.slides[0].partName)!);
const el = (
  local: string,
  children: NodeSpec[] = [],
  attrs: [string, string][] = []
): NodeSpec => ({
  ns: NS_A,
  local,
  prefix: 'a',
  attrs: attrs.map(([l, value]) => ({ ns: '', local: l, value })),
  children,
});

/** Insert `spec` at `index` in the first `name` element, as one store write. */
function insertAt(e: PptxEditorInstance, name: string, index: number, spec: NodeSpec) {
  const part = e.session!.info.slides[0].partName;
  const parent = find(xmlPartOf(e.session!.store.package, part)!.root, name) as Element;
  const r = e.session!.store.transact([{ op: 'insertNode', part, parent: parent.id, index, spec }]);
  expect(r.ok).toBe(true);
}
function removeFirst(e: PptxEditorInstance, name: string) {
  const part = e.session!.info.slides[0].partName;
  const node = find(xmlPartOf(e.session!.store.package, part)!.root, name) as Element;
  expect(e.session!.store.transact([{ op: 'removeNode', part, node: node.id }]).ok).toBe(true);
}
function selectFirst(e: PptxEditorInstance) {
  const id = e.session!.layout(0)!.items[0]!.element.nodeId;
  e.setSelection({ kind: 'shapes', nodeIds: [id] });
}
/** Every child list in the room: no key twice, and no key in two lists. */
function oneParentEach(doc: Y.Doc) {
  const parentOf = new Map<string, string>();
  for (const [key, rec] of doc.getMap(NODES_KEY)) {
    const list = rec instanceof Y.Map ? rec.get(F.children) : undefined;
    if (!(list instanceof Y.Array)) continue;
    const kids = list.toArray() as string[];
    expect(new Set(kids).size).toBe(kids.length);
    for (const k of kids) {
      expect(parentOf.get(k) ?? key).toBe(key);
      parentOf.set(k, key);
    }
  }
}

for (const ids of [
  [1, 2],
  [2, 1],
] as const) {
  describe(`writes into merged copies (client ids ${ids})`, () => {
    test('a child that a peer removes leaves every copy', async () => {
      const { eds, pause, flush } = await peers(ids);
      const [a, b] = eds as [PptxEditorInstance, PptxEditorInstance];
      pause();
      insertAt(a, 'a:p', 0, el('pPr', [el('defRPr', [], [['b', '1']])]));
      insertAt(b, 'a:p', 0, el('pPr', [el('defRPr', [], [['i', '1']])]));
      flush();
      expect(fingerprint(a)).toBe(fingerprint(b));
      removeFirst(a, 'a:defRPr');
      expect(fingerprint(a)).toBe(fingerprint(b));
      expect(xml(b)).not.toContain('<a:defRPr');
      // A later edit by the other peer does not bring it back.
      selectFirst(b);
      expect(b.exec({ type: 'setAlign', align: 'ctr' }).ok).toBe(true);
      expect(fingerprint(a)).toBe(fingerprint(b));
      expect(xml(a)).not.toContain('<a:defRPr');
    });

    test('a new child keeps its place before a peer child that the read moved', async () => {
      const { eds, pause, flush } = await peers(ids);
      const [a, b] = eds as [PptxEditorInstance, PptxEditorInstance];
      pause();
      insertAt(a, 'a:p', 0, el('pPr', [], [['algn', 'r']]));
      insertAt(b, 'a:p', 0, el('r', [el('t', [{ text: 'BBB' } as unknown as NodeSpec])]));
      flush();
      expect(fingerprint(a)).toBe(fingerprint(b));
      insertAt(a, 'a:p', 1, el('r', [el('t', [{ text: 'NEW' } as unknown as NodeSpec])]));
      expect(fingerprint(a)).toBe(fingerprint(b));
      const text = xml(a);
      expect(text.indexOf('NEW')).toBeLessThan(text.indexOf('BBB'));
    });

    test("a copy's child with no slot stays when the copy's writer edits it", async () => {
      const { eds, pause, flush } = await peers(ids);
      const [a, b] = eds as [PptxEditorInstance, PptxEditorInstance];
      const alt: NodeSpec = {
        ns: NS_MC,
        local: 'AlternateContent',
        prefix: 'mc',
        nsDecls: [{ prefix: 'mc', uri: NS_MC }],
        children: [{ ns: NS_MC, local: 'Fallback', prefix: 'mc' }],
      };
      pause();
      insertAt(a, 'a:p', 0, el('pPr', [el('lnSpc', [el('spcPct', [], [['val', '150000']])])]));
      insertAt(b, 'a:p', 0, el('pPr', [alt], [['marL', '10']]));
      flush();
      expect(fingerprint(a)).toBe(fingerprint(b));
      expect(xml(a)).toContain('AlternateContent');
      selectFirst(b);
      expect(b.exec({ type: 'setAlign', align: 'r' }).ok).toBe(true);
      expect(a.exec({ type: 'undo' }).ok).toBe(true);
      expect(fingerprint(a)).toBe(fingerprint(b));
      expect(xml(a)).toContain('AlternateContent');
    });
  });
}

test('three peers: a companion keeps one parent and no list grows on writes', async () => {
  for (const ids of [
    [1, 2, 3],
    [1, 3, 2],
    [2, 1, 3],
    [3, 2, 1],
  ]) {
    const { eds, docs, pause, flush } = await peers(ids);
    pause();
    insertAt(eds[0]!, 'a:p', 0, el('pPr'));
    insertAt(eds[1]!, 'a:p', 0, el('pPr', [el('defRPr', [], [['b', '1']])]));
    insertAt(eds[2]!, 'a:p', 0, el('pPr', [el('defRPr', [], [['i', '1']])]));
    flush();
    selectFirst(eds[0]!);
    for (let i = 0; i < 6; i++)
      expect(eds[0]!.exec({ type: 'setAlign', align: i % 2 ? 'ctr' : 'r' }).ok).toBe(true);
    for (const d of docs) oneParentEach(d);
    for (const e of eds) expect(fingerprint(e)).toBe(fingerprint(eds[0]!));
  }
});

describe('a group over the copy limit', () => {
  for (const ids of [
    [1, 2],
    [2, 1],
  ] as const)
    test(`refuses the part until the copies are gone (client ids ${ids})`, async () => {
      const { eds, docs, sessions } = await peers(ids);
      const [a, b] = eds as [PptxEditorInstance, PptxEditorInstance];
      // A hostile third doc with the room, linked to both peers.
      const c = new Y.Doc();
      c.clientID = 99;
      Y.applyUpdate(c, Y.encodeStateAsUpdate(docs[0]!));
      link(c, docs[0]!);
      link(c, docs[1]!);
      const nodes = c.getMap(NODES_KEY) as Y.Map<Y.Map<unknown>>;
      const shellOf = (rec: Y.Map<unknown>) => rec.get(F.shell) as { ns?: string; local?: string };
      const textUnder = (key: string, depth = 0): string => {
        const rec = nodes.get(key);
        if (!rec || depth > 8) return '';
        const t = rec.get(F.text);
        if (t instanceof Y.Text) return t.toString();
        const kids = rec.get(F.children);
        return kids instanceof Y.Array
          ? (kids.toArray() as string[]).map((k) => textUnder(k, depth + 1)).join('')
          : '';
      };
      const pKey = [...nodes.keys()].find((k) => {
        const s = shellOf(nodes.get(k)!);
        return s?.ns === NS_A && s.local === 'p' && textUnder(k) === 'Hello';
      })!;
      const pRec = nodes.get(pKey)!;
      const part = pRec.get(F.part) as string;
      const hostile = Array.from(
        { length: ROOM_LIMITS.maxCopiesPerGroup + 2 },
        (_, i) => `99.h${i.toString(36)}`
      );
      c.transact(() => {
        for (const k of hostile) {
          const rec = new Y.Map<unknown>();
          rec.set(F.part, part);
          rec.set(F.key, k);
          rec.set(F.shell, { ns: NS_A, local: 'pPr', prefix: 'a', nsDecls: [] });
          rec.set(F.attrs, new Y.Map());
          rec.set(F.children, new Y.Array());
          nodes.set(k, rec);
        }
        (pRec.get(F.children) as Y.Array<string>).insert(0, hostile);
      });
      for (const s of sessions) {
        expect(s.statusSnapshot().reason?.code).toBe('remote-apply-failed');
      }
      expect(xml(a)).toContain('Hello');
      expect(xml(b)).toContain('Hello');
      // A write of the paragraph's parent keeps the paragraph listed in the room.
      insertAt(a, 'p:txBody', 2, el('p'));
      const parentOf = [...nodes.values()].find((rec) => {
        const kids = rec.get(F.children);
        return kids instanceof Y.Array && (kids.toArray() as string[]).includes(pKey);
      });
      expect(parentOf).toBeDefined();
      // The hostile copies go: the part reads again on every replica.
      c.transact(() => {
        const list = pRec.get(F.children) as Y.Array<string>;
        const at = (list.toArray() as string[]).indexOf(hostile[0]!);
        list.delete(at, hostile.length);
        for (const k of hostile) nodes.delete(k);
      });
      for (const s of sessions) expect(s.statusSnapshot().status).toBe('ready');
      expect(xml(a)).toContain('Hello');
      expect(fingerprint(a)).toBe(fingerprint(b));
    });
});
