// Collaboration verification, as AGENTS.md requires: two editors on two linked Y.Docs,
// concurrent edits made with the link paused, deletion while a peer edits, undo, reconnect,
// schema mismatch, save and reopen, presence cleanup, and the caret that follows a peer's edit.

import { describe, expect, test } from 'bun:test';
import * as Y from 'yjs';
import { Awareness, applyAwarenessUpdate, encodeAwarenessUpdate } from 'y-protocols/awareness';
import { createPptxEditor, type PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import {
  canonicalPackageFingerprint,
  readPptx,
  writePptx,
  xmlPartOf,
  type Element,
} from '@pptx-editor.dev/core/store';
import { buildPptx, para, run, shape } from '../../../../scripts/fixtures/builder.ts';
import { collaborationModule, createYjsCollaborationSession } from '../index.ts';
import type { YjsCollaborationSession } from '../index.ts';
import { META_KEY, NODES_KEY } from '../collaboration/yjs/schema.ts';

const IN = 914_400;
const deck = (): Uint8Array =>
  buildPptx({
    slides: [
      {
        layout: 'title',
        shapes:
          shape({ id: 2, x: IN, y: IN, cx: 4 * IN, cy: IN, text: para(run('Hello')) }) +
          shape({ id: 3, x: IN, y: 3 * IN, cx: 4 * IN, cy: IN, text: para(run('World')) }),
      },
      { layout: 'blank', shapes: '' },
    ],
  });

function linkDocs(a: Y.Doc, b: Y.Doc) {
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

function linkAwareness(a: Awareness, b: Awareness): void {
  const send =
    (from: Awareness, to: Awareness) =>
    (c: { added: number[]; updated: number[]; removed: number[] }, origin: unknown) => {
      if (origin === 'link') return;
      const changed = [...c.added, ...c.updated, ...c.removed];
      applyAwarenessUpdate(to, encodeAwarenessUpdate(from, changed), 'link');
    };
  a.on('update', send(a, b));
  b.on('update', send(b, a));
}

interface Replica {
  editor: PptxEditorInstance;
  session: YjsCollaborationSession;
}

async function replica(doc: Y.Doc, actorId: string, awareness?: Awareness): Promise<Replica> {
  const session = createYjsCollaborationSession({
    doc,
    documentId: 'deck',
    identity: { actorId, name: actorId.toUpperCase() },
    ...(awareness ? { awareness } : {}),
  });
  const editor = createPptxEditor({
    fixedMeasurer: true,
    substitutes: false,
    modules: [collaborationModule({ session })],
  });
  expect((await editor.load(deck())).ok).toBe(true);
  return { editor, session };
}

async function pair() {
  const docA = new Y.Doc();
  const docB = new Y.Doc();
  const sync = linkDocs(docA, docB);
  const awA = new Awareness(docA);
  const awB = new Awareness(docB);
  linkAwareness(awA, awB);
  const a = await replica(docA, 'a', awA);
  const b = await replica(docB, 'b', awB);
  return { a, b, sync, docB };
}

const fingerprint = (e: PptxEditorInstance) =>
  canonicalPackageFingerprint(e.session!.store.package);
const part = (e: PptxEditorInstance, i = 0) => e.session!.info.slides[i].partName;

function find(e: PptxEditorInstance, match: (el: Element) => boolean): Element[] {
  const out: Element[] = [];
  const walk = (el: Element) => {
    for (const c of el.children) {
      if (c.type !== 'element') continue;
      if (match(c)) out.push(c);
      walk(c);
    }
  };
  walk(xmlPartOf(e.session!.store.package, part(e))!.root);
  return out;
}
const textEls = (e: PptxEditorInstance) => find(e, (el) => el.local === 't' && el.kind === 'text');
const texts = (e: PptxEditorInstance) =>
  textEls(e).map((t) => t.children.map((c) => (c.type === 'text' ? c.value : '')).join(''));
const shapes = (e: PptxEditorInstance) => find(e, (el) => el.kind === 'sp');
function setText(e: PptxEditorInstance, i: number, text: string) {
  const node = textEls(e)[i];
  const r = e.session!.store.transact([{ op: 'setText', part: part(e), node: node.id, text }]);
  expect(r.ok).toBe(true);
}
/** Save, reopen, and fingerprint: the replica's deck survives a round trip. */
function savedFingerprint(e: PptxEditorInstance): string {
  const reopened = readPptx(writePptx(e.session!.store.package));
  if (!reopened.ok) throw new Error('reopen failed');
  return canonicalPackageFingerprint(reopened.pkg);
}

describe('collaboration verification', () => {
  test('deleting a shape a peer edits, then undoing the delete, keeps the peer text', async () => {
    const { a, b, sync } = await pair();
    a.editor.setSelection({ kind: 'shapes', nodeIds: [shapes(a.editor)[0].id] });
    sync.pause();
    expect(a.editor.exec({ type: 'delete' }).ok).toBe(true);
    setText(b.editor, 0, 'Hello from B');
    sync.flush();
    expect(texts(a.editor)).toEqual(['World']);
    expect(fingerprint(a.editor)).toBe(fingerprint(b.editor));

    expect(a.editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(texts(a.editor)).toEqual(['Hello from B', 'World']);
    expect(texts(b.editor)).toEqual(['Hello from B', 'World']);
    expect(fingerprint(a.editor)).toBe(fingerprint(b.editor));
  });

  test('undo of my typing keeps what a peer typed in the same text node', async () => {
    const { a, b, sync } = await pair();
    sync.pause();
    setText(a.editor, 0, 'Hello A');
    setText(b.editor, 0, 'B Hello');
    sync.flush();
    expect(texts(a.editor)[0]).toBe('B Hello A');
    expect(a.editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(texts(a.editor)[0]).toBe('B Hello');
    expect(texts(b.editor)[0]).toBe('B Hello');
  });

  test('after undo, both replicas save and reopen to the same deck', async () => {
    const { a, b } = await pair();
    expect(a.editor.exec({ type: 'addSlide' }).ok).toBe(true);
    setText(b.editor, 1, 'World from B');
    expect(a.editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(savedFingerprint(a.editor)).toBe(savedFingerprint(b.editor));
    expect(a.editor.session!.info.slides.length).toBe(2);
  });

  test('a replica goes offline, both edit, and they converge on reconnect', async () => {
    const { a, b, sync } = await pair();
    const states: string[] = [];
    a.session.subscribeStatus((status, reason) => states.push(`${status}:${reason ?? ''}`));
    a.session.setConnected(false);
    sync.pause();
    setText(a.editor, 0, 'Offline A');
    setText(b.editor, 1, 'Online B');
    expect(a.session.statusSnapshot().reason?.code).toBe('transport-disconnected');
    sync.flush();
    a.session.setConnected(true);
    expect(texts(a.editor)).toEqual(['Offline A', 'Online B']);
    expect(fingerprint(a.editor)).toBe(fingerprint(b.editor));
    expect(states).toEqual(['disconnected:transport-disconnected', 'ready:']);
  });

  test('a room with another schema version refuses to attach', async () => {
    const doc = new Y.Doc();
    doc.getMap(META_KEY).set('schema', 99);
    const { session } = await replica(doc, 'a');
    expect(session.statusSnapshot()).toMatchObject({
      status: 'error',
      reason: { code: 'schema-version-mismatch' },
    });
    expect(session.attached).toBe(false);
  });

  test('an undo whose rebuild is refused returns false and reports the failure', async () => {
    const { a, b, docB } = await pair();
    setText(a.editor, 0, 'Hello A');
    // A hostile peer breaks the text element A's undo needs to rebuild.
    const nodes = docB.getMap(NODES_KEY);
    for (const rec of nodes.values()) {
      if (!(rec instanceof Y.Map) || rec.get('p') !== part(b.editor)) continue;
      const shell = rec.get('s') as Record<string, unknown> | undefined;
      if (shell?.local === 't') {
        docB.transact(() => rec.set('s', { ...shell, local: 'bad name' }), 'hostile');
        break;
      }
    }
    const before = fingerprint(a.editor);
    expect(a.session.undo!()).toBe(false);
    expect(a.session.statusSnapshot().reason?.code).toBe('remote-apply-failed');
    expect(fingerprint(a.editor)).toBe(before);
  });
});

describe('presence cleanup and caret follow', () => {
  test('detaching the editor and destroying the session remove presence', async () => {
    const { a, b } = await pair();
    a.editor.setSelection({ kind: 'shapes', nodeIds: [shapes(a.editor)[1].id] });
    expect(b.session.remoteSelections!()).toHaveLength(1);
    a.editor.destroy();
    expect(b.session.remoteSelections!()).toEqual([]);
    expect(
      b.session
        .participants()
        .map((p) => p.actorId)
        .sort()
    ).toEqual(['a', 'b']);
    a.session.destroy();
    expect(b.session.participants().map((p) => p.actorId)).toEqual(['b']);
  });

  test('my caret follows its character when a peer types before it', async () => {
    const { a, b } = await pair();
    const shapeId = shapes(a.editor)[0].id;
    const caret = { paragraph: 0, offset: 2 };
    a.editor.setSelection({ kind: 'text', nodeId: shapeId, anchor: caret, focus: caret });
    setText(b.editor, 0, 'XXHello');
    const sel = a.editor.snapshot().selection;
    expect(sel).toEqual({
      kind: 'text',
      nodeId: shapeId,
      anchor: { paragraph: 0, offset: 4 },
      focus: { paragraph: 0, offset: 4 },
    });
  });

  test('undo returns to the slide and selection of the undone change', async () => {
    const { a } = await pair();
    const shapeId = shapes(a.editor)[0].id;
    a.editor.setSelection({ kind: 'shapes', nodeIds: [shapeId] });
    expect(a.editor.exec({ type: 'delete' }).ok).toBe(true);
    a.editor.goToSlide(1);
    expect(a.editor.snapshot().currentSlide).toBe(1);
    expect(a.editor.exec({ type: 'undo' }).ok).toBe(true);
    const snap = a.editor.snapshot();
    expect(snap.currentSlide).toBe(0);
    expect(snap.selection).toMatchObject({ kind: 'shapes', nodeIds: [shapeId] });
  });
});
