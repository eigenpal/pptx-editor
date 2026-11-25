// Replace All through the automation host (`replaceText`, which `Presentation.replaceAll` in
// `@pptx-editor.dev/editor-api` sends: `packages/editor-api/src/__tests__/replace-all.test.ts`
// checks that the member sends exactly this operation) on two editors over two linked Y.Docs,
// as AGENTS.md "Collaboration verification" asks. One peer replaces while the other types in a
// matched paragraph, with the link paused. The replicas converge, live and after save and
// reopen, and an undo of the replacement converges too.

import { describe, expect, test } from 'bun:test';
import * as Y from 'yjs';
import type { AutomationOperation } from '@pptx-editor.dev/core/automation';
import {
  createBrowserAutomationHost,
  createPptxEditor,
  type PptxEditorInstance,
} from '@pptx-editor.dev/core/editor';
import {
  attr,
  canonicalPackageFingerprint,
  child,
  children,
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
          cx: 6 * IN,
          cy: 2 * IN,
          text: para(run('Acme', 'b="1"') + run(' plan for Acme')) + para(run('Other Acme line')),
          bodyPr: '<a:bodyPr wrap="square"><a:noAutofit/></a:bodyPr>',
        }),
      },
    ],
  });

/** What `Presentation.replaceAll('Acme', 'Globex', { matchCase: true })` sends. */
const REPLACE: AutomationOperation = {
  op: 'replaceText',
  query: 'Acme',
  replacement: 'Globex',
  options: { matchCase: true },
};

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

/**
 * Two replicas. Yjs orders two inserts at one place by client id, so each test runs with both
 * orders of the ids.
 */
async function replicas(ids: readonly [number, number]) {
  const docs = [new Y.Doc(), new Y.Doc()] as const;
  docs[0].clientID = ids[0];
  docs[1].clientID = ids[1];
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
    expect((await editor.load(deck())).ok).toBe(true);
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

const shapeOf = (e: PptxEditorInstance): Element => {
  const root = xmlPartOf(e.session!.store.package, e.session!.info.slides[0]!.partName)!.root;
  return elements(find(root, 'p:spTree')!).find((c) => c.local === 'sp')!;
};

/** The text of each paragraph of the shape. */
const paragraphs = (e: PptxEditorInstance): string[] =>
  children(child(shapeOf(e), 'p:txBody'), 'a:p').map((p) =>
    children(p, 'a:r')
      .map((r) => {
        const t = child(r, 'a:t');
        return t?.children.map((c) => (c.type === 'text' ? c.value : '')).join('') ?? '';
      })
      .join('')
  );

/** Whether the run that starts the first paragraph is bold. */
const firstRunBold = (e: PptxEditorInstance): boolean => {
  const r = children(children(child(shapeOf(e), 'p:txBody'), 'a:p')[0], 'a:r')[0];
  return attr(child(r, 'a:rPr'), 'b') === '1';
};

function replace(e: PptxEditorInstance): number {
  const response = createBrowserAutomationHost(e).execute({ operations: [REPLACE] });
  if (!response.ok) throw new Error(JSON.stringify(response.results));
  const result = response.results[0];
  return result?.status === 'ok' && result.value.kind === 'number' ? result.value.value : -1;
}

/** Put the caret at `offset` in paragraph `paragraph` of the shape, and type `text`. */
function typeAt(e: PptxEditorInstance, paragraph: number, offset: number, text: string): void {
  const nodeId = shapeOf(e).id;
  e.setSelection({
    kind: 'text',
    nodeId,
    anchor: { paragraph, offset },
    focus: { paragraph, offset },
  });
  expect(e.input('text', text).ok).toBe(true);
}

const ORDERS: readonly (readonly [number, number])[] = [
  [1_001, 2_002],
  [2_002, 1_001],
];

describe.each(ORDERS)('Replace All under collaboration (client ids %d, %d)', (...ids) => {
  test('a peer types in a matched paragraph during the replacement', async () => {
    const { a, b, sync } = await replicas(ids);
    sync.pause();
    expect(replace(a)).toBe(3);
    typeAt(b, 0, 'Acme plan'.length, ' today');
    sync.flush();
    converged(a, b);
    expect(paragraphs(a)).toEqual(['Globex plan today for Globex', 'Other Globex line']);
    // The new text kept the bold run of the word it replaced.
    expect(firstRunBold(a)).toBe(true);
  });

  test('a peer types right after a matched word during the replacement', async () => {
    const { a, b, sync } = await replicas(ids);
    sync.pause();
    expect(replace(a)).toBe(3);
    typeAt(b, 0, 'Acme plan for Acme'.length, '!');
    sync.flush();
    converged(a, b);
    // Two inserts at one place: Yjs puts them in client id order. Both survive.
    expect(['Globex plan for Globex!', 'Globex plan for !Globex']).toContain(paragraphs(a)[0]!);
  });

  test('a peer types inside a matched word during the replacement', async () => {
    const { a, b, sync } = await replicas(ids);
    sync.pause();
    expect(replace(a)).toBe(3);
    typeAt(b, 1, 'Other Ac'.length, 'XY');
    sync.flush();
    converged(a, b);
    // The typed characters survive the merge; the rest of the deck took the replacement.
    const [first, second] = paragraphs(a);
    expect(first).toBe('Globex plan for Globex');
    expect(second).toContain('XY');
    expect(second!.startsWith('Other ')).toBe(true);
    expect(second!.endsWith(' line')).toBe(true);
  });

  test('undo of the replacement converges and keeps the peer typing', async () => {
    const { a, b, sync } = await replicas(ids);
    sync.pause();
    expect(replace(a)).toBe(3);
    typeAt(b, 0, 'Acme plan'.length, '!');
    sync.flush();
    converged(a, b);
    // The sync was one undo step on the replacing peer.
    expect(a.exec({ type: 'undo' }).ok).toBe(true);
    converged(a, b);
    expect(paragraphs(a)).toEqual(['Acme plan! for Acme', 'Other Acme line']);
    expect(a.exec({ type: 'redo' }).ok).toBe(true);
    converged(a, b);
    expect(paragraphs(a)).toEqual(['Globex plan! for Globex', 'Other Globex line']);
  });
});
