/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Test support for the text lineage cases: N editors on N linked Y.Docs (a full mesh whose
// links pause together), readers of the slide's paragraphs, and the table merge ops.

import { expect } from 'bun:test';
import * as Y from 'yjs';
import { createPptxEditor, type PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import {
  canonicalPackageFingerprint,
  readPptx,
  xmlPartOf,
  type Element,
  type PptxOp,
} from '@pptx-editor.dev/core/store';
import { buildPptx, cell, para, run, shape, table } from '../../../../scripts/fixtures/builder.ts';
import {
  collaborationModule,
  createYjsCollaborationSession,
  type YjsCollaborationSession,
} from '../index.ts';

export const IN = 914_400;
const A_NS = 'http://schemas.openxmlformats.org/drawingml/2006/main';

/** A full mesh of sync links that pause and flush together. */
function mesh(docs: readonly Y.Doc[]) {
  const queue: [Y.Doc, Uint8Array][] = [];
  let paused = false;
  for (const from of docs)
    from.on('update', (update: Uint8Array, origin: unknown) => {
      if (origin === 'link') return;
      for (const to of docs) {
        if (to === from) continue;
        if (paused) queue.push([to, update]);
        else Y.applyUpdate(to, update, 'link');
      }
    });
  return {
    pause: () => (paused = true),
    flush: () => {
      paused = false;
      for (const [to, update] of queue.splice(0)) Y.applyUpdate(to, update, 'link');
    },
    /** Deliver only what waits for `to`, and keep the links paused. */
    deliverTo: (to: Y.Doc) => {
      const rest: [Y.Doc, Uint8Array][] = [];
      for (const [target, update] of queue.splice(0))
        if (target === to) Y.applyUpdate(to, update, 'link');
        else rest.push([target, update]);
      queue.push(...rest);
    },
  };
}

/** The session of each editor that `peers` made, so `converged` can check its status. */
const sessionOf = new WeakMap<PptxEditorInstance, YjsCollaborationSession>();

/** Editors on linked Y.Docs with the given client ids, all loaded from `bytes`. */
export async function peers(bytes: Uint8Array, clientIds: readonly number[]) {
  const docs = clientIds.map((id) => {
    const doc = new Y.Doc();
    doc.clientID = id;
    return doc;
  });
  const sync = mesh(docs);
  const editors: PptxEditorInstance[] = [];
  for (const [i, doc] of docs.entries()) editors.push(await openPeer(doc, `p${i}`, bytes));
  return { editors, sync, docs };
}

/** An editor with a collaboration session on `doc`, loaded from `bytes`. */
export async function openPeer(
  doc: Y.Doc,
  actorId: string,
  bytes: Uint8Array
): Promise<PptxEditorInstance> {
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
  expect(session.status()).toBe('ready');
  sessionOf.set(editor, session);
  return editor;
}

/** The sessions of editors that `peers` made. */
export const sessionsOf = (editors: readonly PptxEditorInstance[]): YjsCollaborationSession[] =>
  editors.map((e) => sessionOf.get(e)!);

/**
 * Two editors, A and B. Client ids are fixed, so the tie rule (the smaller client id) is the
 * same on every run; `swap` gives A the larger id.
 */
export async function replicas(bytes: Uint8Array, swap = false) {
  const { editors, sync, docs } = await peers(bytes, swap ? [2_002, 1_001] : [1_001, 2_002]);
  return { a: editors[0]!, b: editors[1]!, sync, docs: docs as [Y.Doc, Y.Doc] };
}

/** The client-id orders a two-peer case runs in. */
export const TWO_ORDERS: readonly (readonly number[])[] = [
  [1_001, 2_002],
  [2_002, 1_001],
];

export const slide1 = (e: PptxEditorInstance) => e.session!.info.slides[0]!.partName;
export const root = (e: PptxEditorInstance) => xmlPartOf(e.session!.store.package, slide1(e))!.root;
export const kids = (el: Element) => el.children.filter((c): c is Element => c.type === 'element');
export function all(el: Element, local: string, out: Element[] = []): Element[] {
  for (const c of kids(el)) {
    if (c.ns === A_NS && c.local === local) out.push(c);
    all(c, local, out);
  }
  return out;
}
export const textOf = (el: Element): string =>
  all(el, 't')
    .map((t) => t.children.map((c) => (c.type === 'text' ? c.value : '')).join(''))
    .join('');
/** Each paragraph's text, in document order. */
export const paragraphs = (e: PptxEditorInstance) => all(root(e), 'p').map(textOf);
/** The runs and breaks of the first paragraph that holds `text`: run texts, `<br>` for a break. */
export function inlineOf(e: PptxEditorInstance, text: string): string[] {
  const p = all(root(e), 'p').find((x) => textOf(x).includes(text));
  return p
    ? kids(p).flatMap((c) => (c.local === 'br' ? ['<br>'] : c.local === 'r' ? [textOf(c)] : []))
    : [];
}

/**
 * Every replica applied every remote change (its session is `ready`), the replicas match, and
 * the saved file reopens with the same paragraphs. Returns the paragraphs.
 */
export function converged(...editors: readonly PptxEditorInstance[]): string[] {
  for (const e of editors)
    expect(sessionOf.get(e)?.statusSnapshot()).toMatchObject({ status: 'ready' });
  const [first, ...rest] = editors;
  for (const e of rest) {
    expect(canonicalPackageFingerprint(e.session!.store.package)).toBe(
      canonicalPackageFingerprint(first!.session!.store.package)
    );
    expect(paragraphs(e)).toEqual(paragraphs(first!));
  }
  const r = readPptx(first!.save()!);
  if (!r.ok) throw new Error('reopen failed');
  expect(all(xmlPartOf(r.pkg, slide1(first!))!.root, 'p').map(textOf)).toEqual(paragraphs(first!));
  return paragraphs(first!);
}

const textBox = (body: string) =>
  buildPptx({
    slides: [
      {
        layout: 'blank',
        shapes: shape({ id: 2, x: IN, y: IN, cx: 4 * IN, cy: IN, text: body }),
      },
    ],
  });
export const oneRun = (text: string) => textBox(para(run(text)));
export const twoParagraphs = () => textBox(para(run('First')) + para(run('Hello world')));

/** The bold and italic attributes of the run that holds exactly `text`. */
export function formatOf(e: PptxEditorInstance, text: string): Record<string, string> {
  const r = all(root(e), 'r').find((x) => textOf(x) === text);
  const rPr = r && kids(r).find((c) => c.local === 'rPr');
  const out: Record<string, string> = {};
  for (const at of rPr?.attrs ?? [])
    if (at.local === 'b' || at.local === 'i') out[at.local] = at.value;
  return out;
}

/** Selects `from`..`to` in a paragraph of the slide's first shape. */
export function selectIn(e: PptxEditorInstance, paragraph: number, from: number, to = from) {
  const nodeId = e.session!.layout(0)!.items[0]!.element.nodeId;
  e.setSelection({
    kind: 'text',
    nodeId,
    anchor: { paragraph, offset: from },
    focus: { paragraph, offset: to },
  });
}

/** A two-row table: "A" above, `lower` (paragraphs split at `\n`) below. */
export const tableDeck = (lower = 'B') =>
  buildPptx({
    slides: [
      {
        layout: 'blank',
        shapes: table({
          id: 2,
          x: IN,
          y: IN,
          cx: 4 * IN,
          cy: 2 * IN,
          columns: [4 * IN],
          tblPr: '<a:tblPr firstRow="1"/>',
          rows: `<a:tr h="${IN}">${cell('A')}</a:tr><a:tr h="${IN}">${cell(lower)}</a:tr>`,
        }),
      },
    ],
  });

/** The ops of a vertical merge of the two cells: the text moves up. */
export function mergeOps(e: PptxEditorInstance): PptxOp[] {
  const part = slide1(e);
  const [top, bottom] = all(root(e), 'tc');
  const topBody = kids(top!).find((c) => c.local === 'txBody')!;
  const bottomBody = kids(bottom!).find((c) => c.local === 'txBody')!;
  const moved = kids(bottomBody).filter((c) => c.local === 'p');
  return [
    ...moved.map((p, i): PptxOp => ({
      op: 'moveNode',
      part,
      node: p.id,
      parent: topBody.id,
      index: topBody.children.length + i,
    })),
    {
      op: 'insertNode',
      part,
      parent: bottomBody.id,
      index: bottomBody.children.length - moved.length,
      spec: { ns: A_NS, local: 'p', prefix: 'a' },
    },
    { op: 'setAttribute', part, node: top!.id, name: { ns: '', local: 'rowSpan' }, value: '2' },
    { op: 'setAttribute', part, node: bottom!.id, name: { ns: '', local: 'vMerge' }, value: '1' },
  ];
}
export const firstRow = (e: PptxEditorInstance) => all(root(e), 'tr')[0]!;

/**
 * The ops of a delete of the first row by a peer that saw the merge: the row goes, and the cell
 * below it loses `vMerge`. The store refuses a cell with `vMerge` and no merged block above it
 * (`table-merge` in `packages/core/src/store/store/part-checks.ts`), so a removal alone is not
 * a valid local edit after the merge.
 */
export function deleteMergedFirstRowOps(e: PptxEditorInstance): PptxOp[] {
  const part = slide1(e);
  const below = kids(all(root(e), 'tr')[1]!).find((c) => c.local === 'tc')!;
  return [
    { op: 'removeNode', part, node: firstRow(e).id },
    { op: 'setAttribute', part, node: below.id, name: { ns: '', local: 'vMerge' }, value: null },
  ];
}
export const transact = (e: PptxEditorInstance, ops: PptxOp[]) =>
  expect(e.session!.store.transact(ops).ok).toBe(true);
