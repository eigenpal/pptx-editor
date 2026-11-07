/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Deleting an animated shape on one replica while another edits it. The delete
// removes the shape's `p:timing` entries and the connector ends glued to it in the same
// transaction; both replicas must hold the same slide, and the saved deck must name no removed
// shape. A `p:timing` that no one changed saves as the file wrote it, on each replica.

import { describe, expect, test } from 'bun:test';
import * as Y from 'yjs';
import { createPptxEditor, type PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import {
  attr,
  canonicalPackageFingerprint,
  elements,
  find,
  readPptx,
  semanticDigest,
  serializeXmlPart,
  xmlPartOf,
  type Element,
  type PptxPackage,
} from '@pptx-editor.dev/core/store';
import { buildPptx, para, run, shape, solid } from '../../../../scripts/fixtures/builder.ts';
import { clickEntrances, timingWriter } from '../../../../scripts/fixtures/timing.ts';
import { collaborationModule, createYjsCollaborationSession } from '../index.ts';

const IN = 914_400;

/**
 * Shapes `A` (id 2), `B` (id 3), and `C` (id 4), and `extra` after them; a click entrance for
 * each id of `animated`, or the timing `ext`.
 */
const deck = (
  animated: readonly number[],
  extra = '',
  ext = clickEntrances(animated)
): Uint8Array =>
  buildPptx({
    slides: [
      {
        layout: 'blank',
        shapes:
          ['A', 'B', 'C']
            .map((name, i) =>
              shape({
                id: i + 2,
                name,
                x: (1 + 3 * i) * IN,
                y: IN,
                cx: 2 * IN,
                cy: IN,
                fill: solid('4472C4'),
                text: para(run(name)),
              })
            )
            .join('') + extra,
        ext,
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

async function replica(
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
  return editor;
}

async function pair(bytes = deck([2, 3, 4])) {
  const docA = new Y.Doc();
  const docB = new Y.Doc();
  const l = link(docA, docB);
  return { a: await replica(docA, 'a', bytes), b: await replica(docB, 'b', bytes), l };
}

const fingerprint = (e: PptxEditorInstance) =>
  canonicalPackageFingerprint(e.session!.store.package);
const idOf = (e: PptxEditorInstance, name: string) =>
  e.snapshot().shapes!.find((s) => s.name === name)!.nodeId;
const on = (e: PptxEditorInstance, ...names: string[]) =>
  e.setSelection({ kind: 'shapes', nodeIds: names.map((n) => idOf(e, n)) });
const slideRoot = (pkg: PptxPackage, e: PptxEditorInstance): Element =>
  xmlPartOf(pkg, e.session!.info.slides[0].partName)!.root;

/** Every `spid` in the slide's `p:timing`, sorted. */
function spids(root: Element): string[] {
  const out: string[] = [];
  const walk = (el: Element): void => {
    const v = attr(el, 'spid');
    if (v !== undefined) out.push(v);
    for (const c of elements(el)) walk(c);
  };
  const timing = find(root, 'p:timing');
  if (timing) walk(timing);
  return out.sort();
}

/** The `p:cNvPr/@id` values of the slide's shapes. */
function shapeIds(root: Element): Set<string> {
  const ids = new Set<string>();
  const walk = (el: Element): void => {
    if (el.local === 'cNvPr') ids.add(attr(el, 'id')!);
    for (const c of elements(el)) walk(c);
  };
  walk(find(root, 'p:spTree')!);
  return ids;
}

/** Each time node list and build list keeps a child (`CT_TimeNodeList`, `CT_BuildList`). */
function expectNoEmptyLists(root: Element): void {
  const walk = (el: Element): void => {
    if (['tnLst', 'childTnLst', 'subTnLst', 'bldLst'].includes(el.local))
      expect(elements(el).length).toBeGreaterThan(0);
    for (const c of elements(el)) walk(c);
  };
  const timing = find(root, 'p:timing');
  if (timing) walk(timing);
}

/** Both replicas converge, and each save reopens with valid timing and no dangling `spid`. */
function expectConvergedAndSaved(a: PptxEditorInstance, b: PptxEditorInstance): Element {
  expect(fingerprint(a)).toBe(fingerprint(b));
  let root: Element | undefined;
  for (const e of [a, b]) {
    const r = readPptx(e.save()!);
    expect(r.ok).toBe(true);
    if (!r.ok) throw new Error('reopen failed');
    expect(semanticDigest(r.pkg)).toBe(canonicalPackageFingerprint(r.pkg));
    root = slideRoot(r.pkg, e);
    const ids = shapeIds(root);
    for (const spid of spids(root)) expect(ids.has(spid)).toBe(true);
    expectNoEmptyLists(root);
  }
  return root!;
}

describe('deleting an animated shape on two replicas', () => {
  test('one peer edits the shape while the other deletes it', async () => {
    const { a, b, l } = await pair();
    l.pause();
    on(a, 'A');
    expect(a.exec({ type: 'setShapeEffect', effect: { kind: 'glow', size: 8 } }).ok).toBe(true);
    on(b, 'A');
    expect(b.exec({ type: 'delete' }).ok).toBe(true);
    l.flush();
    const root = expectConvergedAndSaved(a, b);
    expect(spids(root)).toEqual(['3', '3', '4', '4']);
    expect(a.snapshot().shapes!.some((s) => s.name === 'A')).toBe(false);
  });

  test('two peers delete different animated shapes at once', async () => {
    const { a, b, l } = await pair();
    l.pause();
    on(a, 'A');
    expect(a.exec({ type: 'delete' }).ok).toBe(true);
    on(b, 'B');
    expect(b.exec({ type: 'delete' }).ok).toBe(true);
    l.flush();
    const root = expectConvergedAndSaved(a, b);
    expect(spids(root)).toEqual(['4', '4']);
  });

  // Each replica keeps the main sequence for the other shape's step, so the merge leaves it, and
  // the build list, with no child. The writer saves the slide without them.
  test('two peers delete the only two animated shapes at once', async () => {
    const { a, b, l } = await pair(deck([2, 3]));
    l.pause();
    on(a, 'A');
    expect(a.exec({ type: 'delete' }).ok).toBe(true);
    on(b, 'B');
    expect(b.exec({ type: 'delete' }).ok).toBe(true);
    l.flush();
    const root = expectConvergedAndSaved(a, b);
    expect(find(root, 'p:timing')).toBeUndefined();
  });

  test('two peers delete the same animated shape at once', async () => {
    const { a, b, l } = await pair();
    l.pause();
    on(a, 'C');
    expect(a.exec({ type: 'delete' }).ok).toBe(true);
    on(b, 'C');
    expect(b.exec({ type: 'delete' }).ok).toBe(true);
    l.flush();
    const root = expectConvergedAndSaved(a, b);
    expect(spids(root)).toEqual(['2', '2', '3', '3']);
  });

  test('undo of the delete brings the shape and its animation back on both replicas', async () => {
    const { a, b } = await pair();
    const before = fingerprint(a);
    on(a, 'B', 'C');
    expect(a.exec({ type: 'delete' }).ok).toBe(true);
    expect(spids(slideRoot(b.session!.store.package, b))).toEqual(['2', '2']);
    expect(a.exec({ type: 'undo' }).ok).toBe(true);
    expect(fingerprint(a)).toBe(before);
    expect(fingerprint(b)).toBe(before);
    expect(a.exec({ type: 'redo' }).ok).toBe(true);
    expectConvergedAndSaved(a, b);
    expect(spids(slideRoot(b.session!.store.package, b))).toEqual(['2', '2']);
  });
  test('undo after a peer deleted the same shapes brings the shape back without its animation', async () => {
    // Peer b's delete took the whole `p:timing`, so a's undo has no main sequence to put the
    // step back in. The shape comes back, and the slide stays valid on both replicas.
    const { a, b, l } = await pair(deck([2, 3]));
    l.pause();
    on(a, 'A');
    expect(a.exec({ type: 'delete' }).ok).toBe(true);
    on(b, 'A', 'B');
    expect(b.exec({ type: 'delete' }).ok).toBe(true);
    l.flush();
    expect(a.exec({ type: 'undo' }).ok).toBe(true);
    expect(a.snapshot().shapes!.some((s) => s.name === 'A')).toBe(true);
    const root = expectConvergedAndSaved(a, b);
    expect(find(root, 'p:timing')).toBeUndefined();
  });

  test('a peer deletes a shape glued by a connector while the other moves the connector', async () => {
    const cxn =
      '<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="7" name="Line"/><p:cNvCxnSpPr><a:stCxn id="2" idx="3"/>' +
      '<a:endCxn id="3" idx="1"/></p:cNvCxnSpPr><p:nvPr/></p:nvCxnSpPr><p:spPr><a:xfrm>' +
      `<a:off x="${3 * IN}" y="${IN + IN / 2}"/><a:ext cx="${IN}" cy="0"/></a:xfrm>` +
      '<a:prstGeom prst="straightConnector1"><a:avLst/></a:prstGeom></p:spPr></p:cxnSp>';
    const { a, b, l } = await pair(deck([2, 3], cxn));
    l.pause();
    on(a, 'B');
    expect(a.exec({ type: 'delete' }).ok).toBe(true);
    // The connector moves with A, so it stays glued to A; it lets go of B, which stays.
    on(b, 'Line', 'A');
    expect(b.exec({ type: 'nudge', dx: IN / 4, dy: 0 }).ok).toBe(true);
    l.flush();
    const root = expectConvergedAndSaved(a, b);
    const ends = elements(find(root, 'p:cNvCxnSpPr')).map((e) => `${e.local}:${attr(e, 'id')}`);
    expect(ends).toEqual(['stCxn:2']);
  });

  test('two peers move the two shapes an elbow connector glues at once', async () => {
    const cxn =
      '<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="7" name="Line"/><p:cNvCxnSpPr><a:stCxn id="2" idx="3"/>' +
      '<a:endCxn id="3" idx="1"/></p:cNvCxnSpPr><p:nvPr/></p:nvCxnSpPr><p:spPr><a:xfrm>' +
      `<a:off x="${3 * IN}" y="${IN + IN / 2}"/><a:ext cx="${IN}" cy="0"/></a:xfrm>` +
      '<a:prstGeom prst="bentConnector3"><a:avLst/></a:prstGeom></p:spPr></p:cxnSp>';
    const { a, b, l } = await pair(deck([2, 3], cxn));
    l.pause();
    // Each move routes the connector again: both replicas end with one of the two routes.
    on(a, 'A');
    expect(a.exec({ type: 'nudge', dx: 0, dy: IN }).ok).toBe(true);
    on(b, 'B');
    expect(b.exec({ type: 'nudge', dx: 0, dy: -IN }).ok).toBe(true);
    l.flush();
    const root = expectConvergedAndSaved(a, b);
    const ends = elements(find(root, 'p:cNvCxnSpPr')).map((e) => `${e.local}:${attr(e, 'id')}`);
    expect(ends).toEqual(['stCxn:2', 'endCxn:3']);
  });

  test('a p:timing that no one changed saves as the file wrote it on both replicas', async () => {
    // The file holds an empty click step. Peer b deletes a shape with no animation; peer a gets
    // the slide rebuilt from the room, with the same `p:timing`.
    const w = timingWriter();
    const ext = w.timing(w.mainSeq(w.click('') + w.click(w.effect(2))), [2]);
    const { a, b } = await pair(deck([], '', ext));
    on(b, 'C');
    expect(b.exec({ type: 'delete' }).ok).toBe(true);
    expect(fingerprint(a)).toBe(fingerprint(b));
    const saved = (e: PptxEditorInstance): string => {
      const r = readPptx(e.save()!);
      if (!r.ok) throw new Error('reopen failed');
      return serializeXmlPart(xmlPartOf(r.pkg, e.session!.info.slides[0].partName)!);
    };
    expect(saved(a)).toContain('<p:childTnLst/>');
    expect(saved(a)).toBe(saved(b));
  });
});
