// Text Direction and Columns (`setTextDirection`, `setColumns`) through two editors on two
// linked Y.Docs, as AGENTS.md "Collaboration verification" asks: concurrent edits with the link
// paused, undo while a peer edits, and a deletion while a peer edits the same shape. The
// replicas converge, before and after save and reopen. Two refits that both add the shape's
// first `a:normAutofit` keep one child, through the mirror's companion records.

import { describe, expect, test } from 'bun:test';
import * as Y from 'yjs';
import { createPptxEditor, type PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import {
  attr,
  canonicalPackageFingerprint,
  child,
  elements,
  find,
  readPptx,
  writePptx,
  xmlPartOf,
  type Element,
} from '@pptx-editor.dev/core/store';
import { buildPptx, para, run, shape } from '../../../../scripts/fixtures/builder.ts';
import { collaborationModule, createYjsCollaborationSession } from '../index.ts';

const IN = 914_400;
const deck = (): Uint8Array =>
  buildPptx({
    slides: [
      {
        layout: 'blank',
        shapes: shape({
          id: 2,
          name: 'T',
          x: IN,
          y: IN,
          cx: 4 * IN,
          cy: 2 * IN,
          text: para(run('one two three')),
          bodyPr: '<a:bodyPr wrap="square"><a:noAutofit/></a:bodyPr>',
        }),
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

async function replicas(bytes: () => Uint8Array = deck) {
  const docs = [new Y.Doc(), new Y.Doc()] as const;
  const sync = link(docs[0], docs[1]);
  const editors: PptxEditorInstance[] = [];
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
    expect((await editor.load(bytes())).ok).toBe(true);
    editors.push(editor);
  }
  return { a: editors[0]!, b: editors[1]!, sync };
}

const fingerprint = (e: PptxEditorInstance) =>
  canonicalPackageFingerprint(e.session!.store.package);
/** Both replicas match, live and after save and reopen. */
function converged(a: PptxEditorInstance, b: PptxEditorInstance): void {
  expect(fingerprint(a)).toBe(fingerprint(b));
  const saved = (e: PptxEditorInstance) => {
    const reopened = readPptx(writePptx(e.session!.store.package));
    if (!reopened.ok) throw new Error('reopen failed');
    return canonicalPackageFingerprint(reopened.pkg);
  };
  expect(saved(a)).toBe(saved(b));
}
const root = (e: PptxEditorInstance): Element =>
  xmlPartOf(e.session!.store.package, e.session!.info.slides[0].partName)!.root;
const shapeOf = (e: PptxEditorInstance): Element | undefined =>
  elements(find(root(e), 'p:spTree')!).find((c) => attr(find(c, 'p:cNvPr'), 'name') === 'T');
const bodyPr = (e: PptxEditorInstance): Element =>
  child(child(shapeOf(e)!, 'p:txBody')!, 'a:bodyPr')!;
const selectShape = (e: PptxEditorInstance) =>
  e.setSelection({ kind: 'shapes', nodeIds: [shapeOf(e)!.id] });

describe('Text Direction and Columns under collaboration', () => {
  test('two peers set the text direction at once: one value', async () => {
    const { a, b, sync } = await replicas();
    selectShape(a);
    selectShape(b);
    sync.pause();
    expect(a.exec({ type: 'setTextDirection', direction: 'vertical' }).ok).toBe(true);
    expect(b.exec({ type: 'setTextDirection', direction: 'vertical270' }).ok).toBe(true);
    sync.flush();
    converged(a, b);
    expect(['vert', 'vert270']).toContain(attr(bodyPr(a), 'vert') ?? '');
  });

  test('one peer sets the direction and the other the columns: both survive', async () => {
    const { a, b, sync } = await replicas();
    selectShape(a);
    selectShape(b);
    sync.pause();
    expect(a.exec({ type: 'setTextDirection', direction: 'vertical' }).ok).toBe(true);
    expect(b.exec({ type: 'setColumns', count: 2, gapEmu: 91_440 }).ok).toBe(true);
    sync.flush();
    converged(a, b);
    expect([attr(bodyPr(a), 'vert'), attr(bodyPr(a), 'numCol'), attr(bodyPr(a), 'spcCol')]).toEqual(
      ['vert', '2', '91440']
    );
  });

  test("undo reverts this peer's change and keeps the other peer's", async () => {
    const { a, b, sync } = await replicas();
    selectShape(a);
    selectShape(b);
    sync.pause();
    expect(a.exec({ type: 'setColumns', count: 3 }).ok).toBe(true);
    expect(b.exec({ type: 'setTextDirection', direction: 'vertical270' }).ok).toBe(true);
    sync.flush();
    expect(a.exec({ type: 'undo' }).ok).toBe(true);
    converged(a, b);
    expect(attr(bodyPr(a), 'numCol')).toBeUndefined();
    expect(attr(bodyPr(a), 'vert')).toBe('vert270');
    // Redo brings this peer's columns back on both replicas.
    expect(a.exec({ type: 'redo' }).ok).toBe(true);
    converged(a, b);
    expect(attr(bodyPr(b), 'numCol')).toBe('3');
  });

  test('a peer deletes the shape while the other sets its columns: both end without it', async () => {
    const { a, b, sync } = await replicas();
    selectShape(a);
    selectShape(b);
    sync.pause();
    expect(a.exec({ type: 'setColumns', count: 2 }).ok).toBe(true);
    expect(b.exec({ type: 'delete' }).ok).toBe(true);
    sync.flush();
    converged(a, b);
    expect(shapeOf(a)).toBeUndefined();
  });

  test('two refits add the first a:normAutofit of an inherited autofit at once: one child', async () => {
    const lots = [1, 2, 3, 4, 5, 6]
      .map(() => para(run('Lorem ipsum dolor sit amet, consectetur adipiscing elit, sed do.')))
      .join('');
    // A body placeholder with an empty a:bodyPr inherits the master's a:normAutofit.
    const inherited = (): Uint8Array =>
      buildPptx({
        slides: [
          {
            layout: 'blank',
            shapes: shape({
              id: 2,
              name: 'T',
              x: IN,
              y: IN,
              cx: 4 * IN,
              cy: 2 * IN,
              ph: '<p:ph type="body" idx="1"/>',
              bodyPr: '<a:bodyPr/>',
              text: lots,
            }),
          },
        ],
      });
    const { a, b, sync } = await replicas(inherited);
    expect(find(bodyPr(a), 'a:normAutofit')).toBeUndefined();
    selectShape(a);
    selectShape(b);
    sync.pause();
    expect(a.exec({ type: 'setColumns', count: 2 }).ok).toBe(true);
    expect(b.exec({ type: 'setTextDirection', direction: 'vertical' }).ok).toBe(true);
    expect(find(bodyPr(a), 'a:normAutofit')).toBeDefined();
    expect(find(bodyPr(b), 'a:normAutofit')).toBeDefined();
    sync.flush();
    converged(a, b);
    expect(elements(bodyPr(a)).map((c) => c.local)).toEqual(['normAutofit']);
  });
});
