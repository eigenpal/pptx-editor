import { describe, expect, test } from 'bun:test';
import * as Y from 'yjs';
import { createPptxEditor, type PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import {
  canonicalPackageFingerprint,
  xmlPartOf,
  type Element,
  type PptxPackage,
} from '@pptx-editor.dev/core/store';
import { buildPptx, para, run, shape } from '../../../../scripts/fixtures/builder.ts';
import { collaborationModule, createYjsCollaborationSession } from '../index.ts';
import { NODES_KEY } from '../collaboration/yjs/schema.ts';

const IN = 914_400;
const deck = (first = 'Hello', second = 'World'): Uint8Array =>
  buildPptx({
    slides: [
      {
        layout: 'title',
        shapes:
          shape({ id: 2, x: IN, y: IN, cx: 4 * IN, cy: IN, text: para(run(first)) }) +
          shape({ id: 3, x: IN, y: 3 * IN, cx: 4 * IN, cy: IN, text: para(run(second)) }),
      },
    ],
  });

/** Two docs and a sync link that can pause, to make edits concurrent. */
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

async function replica(doc: Y.Doc, actorId: string, bytes: Uint8Array) {
  const session = createYjsCollaborationSession({
    doc,
    documentId: 'deck',
    identity: { actorId, name: actorId },
  });
  const editor = createPptxEditor({
    fixedMeasurer: true,
    substitutes: false,
    modules: [collaborationModule({ session })],
  });
  expect((await editor.load(bytes)).ok).toBe(true);
  return { editor, session };
}

const fingerprint = (e: PptxEditorInstance) =>
  canonicalPackageFingerprint(e.session!.store.package);
const slide1 = (e: PptxEditorInstance) => e.session!.info.slides[0].partName;

/** The `a:t` elements of a part, in document order. */
function textElements(pkg: PptxPackage, part: string): Element[] {
  const out: Element[] = [];
  const walk = (el: Element) => {
    for (const c of el.children) {
      if (c.type !== 'element') continue;
      if (c.local === 't' && c.kind === 'text') out.push(c);
      else walk(c);
    }
  };
  walk(xmlPartOf(pkg, part)!.root);
  return out;
}
const texts = (e: PptxEditorInstance) =>
  textElements(e.session!.store.package, slide1(e)).map((t) =>
    t.children.map((c) => (c.type === 'text' ? c.value : '')).join('')
  );
function setText(e: PptxEditorInstance, index: number, text: string) {
  const node = textElements(e.session!.store.package, slide1(e))[index];
  const r = e.session!.store.transact([{ op: 'setText', part: slide1(e), node: node.id, text }]);
  expect(r.ok).toBe(true);
}

describe('Yjs collaboration session', () => {
  test('the first replica seeds the room; the next adopts it, whatever it opened', async () => {
    const docA = new Y.Doc();
    const docB = new Y.Doc();
    link(docA, docB);
    const a = await replica(docA, 'a', deck());
    const b = await replica(docB, 'b', deck('Other', 'Deck'));
    expect(a.session.status()).toBe('ready');
    expect(b.session.status()).toBe('ready');
    expect(texts(b.editor)).toEqual(['Hello', 'World']);
    expect(fingerprint(b.editor)).toBe(fingerprint(a.editor));
  });

  test('edits, undo, and redo reach the other replica', async () => {
    const docA = new Y.Doc();
    const docB = new Y.Doc();
    link(docA, docB);
    const a = await replica(docA, 'a', deck());
    const b = await replica(docB, 'b', deck());

    expect(a.editor.exec({ type: 'addSlide' }).ok).toBe(true);
    expect(b.editor.session!.info.slides.length).toBe(2);
    expect(fingerprint(b.editor)).toBe(fingerprint(a.editor));

    expect(b.editor.exec({ type: 'insertTextBox' }).ok).toBe(true);
    expect(fingerprint(a.editor)).toBe(fingerprint(b.editor));

    expect(b.editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(fingerprint(a.editor)).toBe(fingerprint(b.editor));
    expect(b.editor.exec({ type: 'redo' }).ok).toBe(true);
    expect(fingerprint(a.editor)).toBe(fingerprint(b.editor));
  });

  test('concurrent edits on different slides both survive', async () => {
    const docA = new Y.Doc();
    const docB = new Y.Doc();
    const sync = link(docA, docB);
    const a = await replica(docA, 'a', deck());
    const b = await replica(docB, 'b', deck());

    sync.pause();
    expect(a.editor.exec({ type: 'addSlide' }).ok).toBe(true);
    setText(b.editor, 0, 'Hello from B');
    sync.flush();

    expect(fingerprint(a.editor)).toBe(fingerprint(b.editor));
    expect(a.editor.session!.info.slides.length).toBe(2);
    expect(texts(a.editor)[0]).toBe('Hello from B');
  });

  test('concurrent typing in one text node merges both insertions', async () => {
    const docA = new Y.Doc();
    const docB = new Y.Doc();
    const sync = link(docA, docB);
    const a = await replica(docA, 'a', deck());
    const b = await replica(docB, 'b', deck());

    sync.pause();
    setText(a.editor, 0, 'Hello A');
    setText(b.editor, 0, 'B Hello');
    sync.flush();

    expect(texts(a.editor)[0]).toBe('B Hello A');
    expect(texts(b.editor)[0]).toBe('B Hello A');
    expect(fingerprint(a.editor)).toBe(fingerprint(b.editor));
  });

  test('a remote edit keeps the node ids of untouched shapes', async () => {
    const docA = new Y.Doc();
    const docB = new Y.Doc();
    link(docA, docB);
    const a = await replica(docA, 'a', deck());
    const b = await replica(docB, 'b', deck());
    const before = textElements(a.editor.session!.store.package, slide1(a.editor))[1].id;
    setText(b.editor, 0, 'Changed');
    expect(texts(a.editor)[0]).toBe('Changed');
    expect(textElements(a.editor.session!.store.package, slide1(a.editor))[1].id).toBe(before);
  });

  test('an invalid remote node is refused and the deck does not change', async () => {
    const docA = new Y.Doc();
    const docB = new Y.Doc();
    link(docA, docB);
    const a = await replica(docA, 'a', deck());
    const b = await replica(docB, 'b', deck());
    const good = fingerprint(a.editor);

    // A hostile peer renames an element to a name XML cannot hold.
    const nodes = docB.getMap(NODES_KEY);
    docB.transact(() => {
      for (const rec of nodes.values()) {
        if (!(rec instanceof Y.Map) || rec.get('p') !== slide1(b.editor)) continue;
        const shell = rec.get('s') as { local?: string } | undefined;
        if (shell?.local === 't') {
          rec.set('s', { ...shell, local: 'bad name' });
          break;
        }
      }
    }, 'hostile');

    expect(a.session.status()).toBe('error');
    expect(a.session.statusSnapshot().reason?.code).toBe('remote-apply-failed');
    expect(a.session.statusSnapshot().recovering).toBe(true);
    expect(fingerprint(a.editor)).toBe(good);
  });

  test('destroy detaches and reports destroyed', async () => {
    const doc = new Y.Doc();
    const a = await replica(doc, 'a', deck());
    expect(a.session.attached).toBe(true);
    a.editor.destroy();
    expect(a.session.attached).toBe(false);
    a.session.destroy();
    expect(a.session.status()).toBe('destroyed');
  });
});
