import { describe, expect, test } from 'bun:test';
import * as Y from 'yjs';
import { Awareness, applyAwarenessUpdate, encodeAwarenessUpdate } from 'y-protocols/awareness';
import {
  createPptxEditor,
  type EditorCollaborationSession,
  type PptxEditorInstance,
} from '@pptx-editor.dev/core/editor';
import {
  canonicalPackageFingerprint,
  xmlPartOf,
  type Element,
  type TextNode,
} from '@pptx-editor.dev/core/store';
import { buildPptx, para, run, shape } from '../../../../scripts/fixtures/builder.ts';
import {
  collaborationModule,
  createYjsCollaborationSession,
  CollaborationSchemaError,
} from '../index.ts';
import { NODES_KEY } from '../collaboration/yjs/schema.ts';

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
    ],
  });

/** Two docs and a sync link that can pause, to make edits concurrent. */
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

/** Two awareness instances linked like a provider would. */
function linkAwareness(a: Awareness, b: Awareness): void {
  const send =
    (from: Awareness, to: Awareness) =>
    (
      { added, updated, removed }: { added: number[]; updated: number[]; removed: number[] },
      origin: unknown
    ) => {
      if (origin === 'link') return;
      const changed = [...added, ...updated, ...removed];
      applyAwarenessUpdate(to, encodeAwarenessUpdate(from, changed), 'link');
    };
  a.on('update', send(a, b));
  b.on('update', send(b, a));
}

async function replica(
  doc: Y.Doc,
  actorId: string,
  awareness?: Awareness
): Promise<{ editor: PptxEditorInstance; session: EditorCollaborationSession }> {
  const session = createYjsCollaborationSession({
    doc,
    documentId: 'deck',
    identity: { actorId, name: actorId.toUpperCase(), color: '#3366ff' },
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

async function pair(withAwareness = false) {
  const docA = new Y.Doc();
  const docB = new Y.Doc();
  const sync = linkDocs(docA, docB);
  const awA = withAwareness ? new Awareness(docA) : undefined;
  const awB = withAwareness ? new Awareness(docB) : undefined;
  if (awA && awB) linkAwareness(awA, awB);
  const a = await replica(docA, 'a', awA);
  const b = await replica(docB, 'b', awB);
  return { a, b, sync, docB, awA, awB };
}

const fingerprint = (e: PptxEditorInstance) =>
  canonicalPackageFingerprint(e.session!.store.package);
const slide1 = (e: PptxEditorInstance) => e.session!.info.slides[0].partName;

function find(e: PptxEditorInstance, match: (el: Element) => boolean): Element[] {
  const out: Element[] = [];
  const walk = (el: Element) => {
    for (const c of el.children) {
      if (c.type !== 'element') continue;
      if (match(c)) out.push(c);
      walk(c);
    }
  };
  walk(xmlPartOf(e.session!.store.package, slide1(e))!.root);
  return out;
}
const textEls = (e: PptxEditorInstance) => find(e, (el) => el.local === 't' && el.kind === 'text');
const texts = (e: PptxEditorInstance) =>
  textEls(e).map((t) => t.children.map((c) => (c.type === 'text' ? c.value : '')).join(''));
const textNode = (e: PptxEditorInstance, i: number) => textEls(e)[i].children[0] as TextNode;
const shapes = (e: PptxEditorInstance) => find(e, (el) => el.kind === 'sp');
function setText(e: PptxEditorInstance, i: number, text: string, historyGroup?: symbol) {
  const node = textEls(e)[i];
  const r = e.session!.store.transact([{ op: 'setText', part: slide1(e), node: node.id, text }], {
    ...(historyGroup ? { historyGroup } : {}),
  });
  expect(r.ok).toBe(true);
}

describe('collaboration-owned undo', () => {
  test('undo reverts my edit and keeps a concurrent peer edit; redo restores it', async () => {
    const { a, b, sync } = await pair();
    sync.pause();
    expect(a.editor.exec({ type: 'addSlide' }).ok).toBe(true);
    setText(b.editor, 0, 'Hello from B');
    sync.flush();
    expect(a.editor.session!.info.slides.length).toBe(2);

    expect(a.editor.can({ type: 'undo' }).ok).toBe(true);
    expect(a.editor.exec({ type: 'undo' }).ok).toBe(true);
    for (const e of [a.editor, b.editor]) {
      expect(e.session!.info.slides.length).toBe(1);
      expect(texts(e)[0]).toBe('Hello from B');
    }
    expect(fingerprint(a.editor)).toBe(fingerprint(b.editor));

    expect(a.editor.exec({ type: 'redo' }).ok).toBe(true);
    expect(b.editor.session!.info.slides.length).toBe(2);
    expect(fingerprint(a.editor)).toBe(fingerprint(b.editor));
  });

  test('undo after a remote change reverts only my change', async () => {
    const { a, b } = await pair();
    setText(a.editor, 1, 'World from A');
    setText(b.editor, 0, 'Hello from B');
    expect(a.editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(texts(a.editor)).toEqual(['Hello from B', 'World']);
    expect(texts(b.editor)).toEqual(['Hello from B', 'World']);
    // B's history holds only B's change.
    expect(b.editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(texts(a.editor)).toEqual(['Hello', 'World']);
  });

  test('a grouped store entry (typing) is one undo step; nothing left is refused', async () => {
    const { a, b } = await pair();
    const typing = Symbol('typing');
    setText(a.editor, 0, 'Hello w', typing);
    setText(a.editor, 0, 'Hello wo', typing);
    setText(a.editor, 0, 'Hello wor', typing);
    expect(texts(b.editor)[0]).toBe('Hello wor');
    expect(a.editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(texts(b.editor)[0]).toBe('Hello');
    expect(a.editor.can({ type: 'undo' }).ok).toBe(false);
    expect(a.editor.exec({ type: 'undo' }).ok).toBe(false);
  });

  test("a peer's edit in the middle of a gesture does not split the gesture's undo step", async () => {
    const { a, b } = await pair();
    const drag = Symbol('drag');
    setText(a.editor, 0, 'Hello w', drag);
    setText(b.editor, 1, 'World from B');
    setText(a.editor, 0, 'Hello wo', drag);
    expect(texts(b.editor)).toEqual(['Hello wo', 'World from B']);
    expect(a.editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(texts(b.editor)).toEqual(['Hello', 'World from B']);
    expect(a.editor.exec({ type: 'undo' }).ok).toBe(false);
  });

  test('a frame of an open handle after a room undo is a new step, not part of an older one', async () => {
    const { a, b } = await pair();
    setText(a.editor, 1, 'World 2');
    const group = a.editor.beginHistoryGroup();
    const id = shapes(a.editor)[0].id;
    a.editor.setSelection({
      kind: 'text',
      nodeId: id,
      anchor: { paragraph: 0, offset: 0 },
      focus: { paragraph: 0, offset: 5 },
    });
    expect(a.editor.exec({ type: 'toggleBold' }, { historyGroup: group }).ok).toBe(true);
    expect(a.editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(a.editor.exec({ type: 'toggleItalic' }, { historyGroup: group }).ok).toBe(true);
    group.end();
    // The italic step goes alone; the older text edit stays until its own undo.
    expect(a.editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(texts(b.editor)).toEqual(['Hello', 'World 2']);
    expect(a.editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(texts(b.editor)).toEqual(['Hello', 'World']);
  });
});

describe('status and failures', () => {
  test('the snapshot keeps its reference until a value changes', async () => {
    const { a } = await pair();
    const first = a.session.statusSnapshot();
    expect(first.status).toBe('ready');
    expect(a.session.statusSnapshot()).toBe(first);
    expect(Object.isFrozen(first)).toBe(true);
  });

  test('an invalid remote node is a recovering error that heals on the next clean update', async () => {
    const { a, b, docB } = await pair();
    const seen: (string | undefined)[] = [];
    a.session.subscribeStatus((status, reason) => seen.push(`${status}:${reason ?? ''}`));
    const nodes = docB.getMap(NODES_KEY);
    let target: Y.Map<unknown> | undefined;
    let shell: Record<string, unknown> | undefined;
    for (const rec of nodes.values()) {
      if (!(rec instanceof Y.Map) || rec.get('p') !== slide1(b.editor)) continue;
      const s = rec.get('s') as Record<string, unknown> | undefined;
      if (s?.local === 't') {
        target = rec;
        shell = s;
        break;
      }
    }
    docB.transact(() => target!.set('s', { ...shell, local: 'bad name' }), 'hostile');
    expect(a.session.statusSnapshot()).toMatchObject({
      status: 'error',
      reason: { code: 'remote-apply-failed' },
      recovering: true,
      failureCount: 1,
    });
    docB.transact(() => target!.set('s', shell), 'repair');
    const healed = a.session.statusSnapshot();
    expect(healed.status).toBe('ready');
    expect(healed.reason).toBeUndefined();
    expect(healed.lastFailure?.code).toBe('remote-apply-failed');
    expect(seen).toEqual(['error:remote-apply-failed', 'ready:']);
  });

  test('the factory refuses an invalid document id, identity, or color', () => {
    const doc = new Y.Doc();
    const code = (f: () => unknown) => {
      try {
        f();
      } catch (e) {
        return e instanceof CollaborationSchemaError ? e.code : 'other';
      }
      return 'none';
    };
    const make = (
      documentId: string,
      identity: { actorId: string; name: string; color?: string }
    ) => code(() => createYjsCollaborationSession({ doc, documentId, identity }));
    expect(make('', { actorId: 'a', name: 'A' })).toBe('invalid-document-id');
    expect(make('d', { actorId: ' ', name: 'A' })).toBe('invalid-identity');
    expect(make('d', { actorId: 'a', name: 'A', color: 'red;}' })).toBe('invalid-identity-color');
  });

  test('destroy reports destroyed with its reason', async () => {
    const { a } = await pair();
    a.editor.destroy();
    a.session.destroy();
    expect(a.session.statusSnapshot().status).toBe('destroyed');
    expect(a.session.statusSnapshot().reason?.code).toBe('collaboration-session-destroyed');
  });
});

describe('presence', () => {
  test('participants join and leave', async () => {
    const { a, b } = await pair(true);
    const listed = (s: EditorCollaborationSession) =>
      s
        .participants()
        .map((p) => `${p.actorId}:${p.isLocal}`)
        .sort();
    expect(listed(a.session)).toEqual(['a:true', 'b:false']);
    expect(listed(b.session)).toEqual(['a:false', 'b:true']);
    // Leaving is the session's job: destroy clears this editor's awareness record.
    b.editor.destroy();
    b.session.destroy();
    expect(listed(a.session)).toEqual(['a:true']);
  });

  test('a shape selection maps onto the peer node ids', async () => {
    const { a, b } = await pair(true);
    const updates: number[] = [];
    b.session.subscribeRemoteSelections!((s) => updates.push(s.length));
    a.session.setLocalSelection!({
      kind: 'shapes',
      part: slide1(a.editor),
      shapes: [shapes(a.editor)[1].id],
    });
    const remote = b.session.remoteSelections!();
    expect(remote).toHaveLength(1);
    expect(remote[0]).toMatchObject({ actorId: 'a', name: 'A', color: '#3366ff' });
    expect(remote[0].selection).toEqual({
      kind: 'shapes',
      part: slide1(b.editor),
      shapes: [shapes(b.editor)[1].id],
    });
    expect(updates.at(-1)).toBe(1);
    a.session.setLocalSelection!(null);
    expect(b.session.remoteSelections!()).toEqual([]);
  });

  test('a text caret follows its character when the peer types before it', async () => {
    const { a, b } = await pair(true);
    const node = textNode(a.editor, 0).id;
    a.session.setLocalSelection!({
      kind: 'text',
      part: slide1(a.editor),
      anchor: { node, offset: 2 },
      head: { node, offset: 2 },
    });
    const caret = () => {
      const s = b.session.remoteSelections!()[0]?.selection;
      return s?.kind === 'text' ? s.anchor : null;
    };
    expect(caret()).toEqual({ node: textNode(b.editor, 0).id, offset: 2 });
    setText(b.editor, 0, 'XXHello');
    expect(caret()).toEqual({ node: textNode(b.editor, 0).id, offset: 4 });
  });
});
