// The Paragraph dialog's command (`setParagraphFormat`) through two editors on two linked Y.Docs,
// as AGENTS.md "Collaboration verification" asks: concurrent edits with the link paused, undo
// while a peer edits, and a deletion while a peer edits the same shape. The replicas converge,
// before and after save and reopen. Concurrent inserts of the same child, or of two ordered
// children, keep one child in schema order through the mirror's companion records (eigenpal/pptx-editor-miro#57), in both
// client orders.

import { describe, expect, test } from 'bun:test';
import * as Y from 'yjs';
import {
  createPptxEditor,
  type EditorCommand,
  type PptxEditorInstance,
} from '@pptx-editor.dev/core/editor';
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
const SPACED = '<a:pPr><a:spcBef><a:spcPts val="600"/></a:spcBef></a:pPr>';
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
          text: para(run('one'), SPACED) + para(run('two'), SPACED),
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

async function replicas(bytes: () => Uint8Array = deck, swap = false) {
  const docs = [new Y.Doc(), new Y.Doc()] as const;
  // Fixed client ids, so the client order that decides a tie is the same on every run.
  docs[0].clientID = swap ? 2_002 : 1_001;
  docs[1].clientID = swap ? 1_001 : 2_002;
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
const pPrs = (e: PptxEditorInstance, i = 0): Element[] =>
  elements(elements(find(shapeOf(e)!, 'p:txBody')!).filter((c) => c.local === 'p')[i]!).filter(
    (c) => c.local === 'pPr'
  );
const select = (e: PptxEditorInstance, paragraph = 0) =>
  e.setSelection({
    kind: 'text',
    nodeId: shapeOf(e)!.id,
    anchor: { paragraph, offset: 0 },
    focus: { paragraph, offset: 1 },
  });
const spaceBefore = (e: PptxEditorInstance, i = 0): string[] =>
  elements(pPrs(e, i)[0])
    .filter((c) => c.local === 'spcBef')
    .map((c) => attr(child(c, 'a:spcPts'), 'val') ?? '?');

describe('setParagraphFormat under collaboration', () => {
  test('two peers set the same space before: one a:spcBef, one value, same deck', async () => {
    const { a, b, sync } = await replicas();
    select(a);
    select(b);
    sync.pause();
    expect(a.exec({ type: 'setParagraphFormat', spaceBeforePt: 12 }).ok).toBe(true);
    expect(b.exec({ type: 'setParagraphFormat', spaceBeforePt: 18 }).ok).toBe(true);
    sync.flush();
    converged(a, b);
    expect(pPrs(a)).toHaveLength(1);
    // The same field merges as one attribute write: one of the two values wins.
    expect(spaceBefore(a)).toHaveLength(1);
    expect(['1200', '1800']).toContain(spaceBefore(a)[0]);
    expect(spaceBefore(b)).toEqual(spaceBefore(a));
  });

  test('two peers set different fields of one paragraph: both survive', async () => {
    const { a, b, sync } = await replicas();
    select(a);
    select(b);
    sync.pause();
    expect(a.exec({ type: 'setParagraphFormat', alignment: 'center' }).ok).toBe(true);
    expect(b.exec({ type: 'setParagraphFormat', spaceAfterPt: 6 }).ok).toBe(true);
    sync.flush();
    converged(a, b);
    const [pPr] = pPrs(a);
    expect(attr(pPr, 'algn')).toBe('ctr');
    expect(attr(find(find(pPr!, 'a:spcAft')!, 'a:spcPts'), 'val')).toBe('600');
  });

  test('two peers set the toolbar line spacing over an existing multiple: one a:lnSpc', async () => {
    const { a, b, sync } = await replicas();
    select(a);
    select(b);
    expect(a.exec({ type: 'setLineSpacing', multiple: 1 }).ok).toBe(true);
    sync.pause();
    expect(a.exec({ type: 'setLineSpacing', multiple: 1.5 }).ok).toBe(true);
    expect(b.exec({ type: 'setLineSpacing', multiple: 2 }).ok).toBe(true);
    sync.flush();
    converged(a, b);
    expect(elements(pPrs(a)[0]).filter((c) => c.local === 'lnSpc')).toHaveLength(1);
  });

  test("undo reverts this peer's change and keeps the other peer's", async () => {
    const { a, b, sync } = await replicas();
    select(a, 0);
    select(b, 1);
    sync.pause();
    expect(a.exec({ type: 'setParagraphFormat', spaceBeforePt: 24 }).ok).toBe(true);
    expect(b.exec({ type: 'setParagraphFormat', spaceBeforePt: 30 }).ok).toBe(true);
    sync.flush();
    expect(a.exec({ type: 'undo' }).ok).toBe(true);
    converged(a, b);
    expect(spaceBefore(a, 0)).toEqual(['600']);
    expect(spaceBefore(a, 1)).toEqual(['3000']);
    expect(a.exec({ type: 'redo' }).ok).toBe(true);
    converged(a, b);
    expect(spaceBefore(b, 0)).toEqual(['2400']);
  });

  test('a peer deletes the shape while the other formats it: both end without it', async () => {
    const { a, b, sync } = await replicas();
    select(a);
    b.setSelection({ kind: 'shapes', nodeIds: [shapeOf(b)!.id] });
    sync.pause();
    expect(a.exec({ type: 'setParagraphFormat', spaceBeforePt: 12 }).ok).toBe(true);
    expect(b.exec({ type: 'delete' }).ok).toBe(true);
    sync.flush();
    converged(a, b);
    expect(shapeOf(a)).toBeUndefined();
  });

  // Children that a paragraph holds once: two peers who add one at once keep one, in schema
  // order (companion records, eigenpal/pptx-editor-miro#57).
  const gapsDeck = (): Uint8Array =>
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
            text: para(run('one'), '<a:pPr/>') + para(run('two')) + para(run('three'), SPACED),
          }),
        },
      ],
    });
  const kids = (e: PptxEditorInstance, i: number) => elements(pPrs(e, i)[0]).map((c) => c.local);
  async function both(paragraph: number, ca: EditorCommand, cb: EditorCommand, swap: boolean) {
    const { a, b, sync } = await replicas(gapsDeck, swap);
    select(a, paragraph);
    select(b, paragraph);
    sync.pause();
    expect(a.exec(ca).ok).toBe(true);
    expect(b.exec(cb).ok).toBe(true);
    sync.flush();
    converged(a, b);
    return a;
  }
  for (const swap of [false, true]) {
    const order = swap ? '(client order B, A)' : '(client order A, B)';
    test(`two peers add the same missing spacing child: one a:spcBef ${order}`, async () => {
      const a = await both(
        0,
        { type: 'setParagraphFormat', spaceBeforePt: 12 },
        { type: 'setParagraphFormat', spaceBeforePt: 18 },
        swap
      );
      expect(kids(a, 0).filter((k) => k === 'spcBef')).toHaveLength(1);
    });
    test(`two peers add Before and After at once: schema order kept ${order}`, async () => {
      const a = await both(
        0,
        { type: 'setParagraphFormat', spaceAfterPt: 6 },
        { type: 'setParagraphFormat', spaceBeforePt: 6 },
        swap
      );
      expect(kids(a, 0)).toEqual(['spcBef', 'spcAft']);
    });
    test(`two peers create a paragraph's first a:pPr: one a:pPr ${order}`, async () => {
      const a = await both(
        1,
        { type: 'setParagraphFormat', alignment: 'center' },
        { type: 'setParagraphFormat', alignment: 'right' },
        swap
      );
      expect(pPrs(a, 1)).toHaveLength(1);
    });
    test(`two peers set the toolbar line spacing with no a:lnSpc: one a:lnSpc ${order}`, async () => {
      const a = await both(
        0,
        { type: 'setLineSpacing', multiple: 1.5 },
        { type: 'setLineSpacing', multiple: 2 },
        swap
      );
      expect(kids(a, 0).filter((k) => k === 'lnSpc')).toHaveLength(1);
    });
    test(`two peers change one spacing value from points to lines: one child ${order}`, async () => {
      const a = await both(
        2,
        { type: 'setParagraphFormat', spaceBeforeLines: 1 },
        { type: 'setParagraphFormat', spaceBeforeLines: 2 },
        swap
      );
      expect(kids(a, 2).filter((k) => k === 'spcBef')).toHaveLength(1);
      expect(elements(elements(pPrs(a, 2)[0])[0]!)).toHaveLength(1);
    });
  }
});
