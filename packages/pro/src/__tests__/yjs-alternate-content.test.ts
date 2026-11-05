/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/

// Two replicas edit a shape in `mc:AlternateContent` whose drawn copy is a `mc:Choice`
// (ECMA-376 Part 3, 10.2.1). Each edit removes the out-of-date copy in `mc:Fallback` in its own
// transaction (`editor/alternate-copies.ts`). Both replicas must converge on one slide, and each
// save must hold no stale `mc:Fallback` copy.

import { describe, expect, test } from 'bun:test';
import * as Y from 'yjs';
import { createPptxEditor, type PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import {
  canonicalPackageFingerprint,
  readPptx,
  semanticDigest,
  serializeXmlPart,
  xmlPartOf,
} from '@pptx-editor.dev/core/store';
import { buildPptx, para, run, shape, solid } from '../../../../scripts/fixtures/builder.ts';
import { collaborationModule, createYjsCollaborationSession } from '../index.ts';

const IN = 914_400;
const MC = 'http://schemas.openxmlformats.org/markup-compatibility/2006';
const P14 = 'http://schemas.microsoft.com/office/powerpoint/2010/main';
const SLIDE = '/ppt/slides/slide1.xml';

const box = (color: string, text: string): string =>
  shape({
    id: 3,
    name: 'Box',
    x: IN,
    y: IN,
    cx: 3 * IN,
    cy: IN,
    fill: solid(color),
    text: para(run(text)),
  });

const deck = (): Uint8Array =>
  buildPptx({
    slides: [
      {
        layout: 'blank',
        shapes:
          `<mc:AlternateContent xmlns:mc="${MC}">` +
          `<mc:Choice xmlns:p14="${P14}" Requires="p14">${box('C00000', 'drawn')}</mc:Choice>` +
          `<mc:Fallback>${box('70AD47', 'stale')}</mc:Fallback></mc:AlternateContent>`,
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

async function replica(doc: Y.Doc, actorId: string): Promise<PptxEditorInstance> {
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
  expect((await editor.load(deck())).ok).toBe(true);
  return editor;
}

async function pair() {
  const docA = new Y.Doc();
  const docB = new Y.Doc();
  const l = link(docA, docB);
  return { a: await replica(docA, 'a'), b: await replica(docB, 'b'), l };
}

const fingerprint = (e: PptxEditorInstance) =>
  canonicalPackageFingerprint(e.session!.store.package);
const boxId = (e: PptxEditorInstance) => e.snapshot().shapes!.find((s) => s.name === 'Box')?.nodeId;

/**
 * Both replicas converge, and each save reopens equal. With `fresh`, the slide holds no stale
 * copy. Returns the slide.
 */
function expectConvergedAndSaved(
  a: PptxEditorInstance,
  b: PptxEditorInstance,
  fresh = true
): string {
  expect(fingerprint(a)).toBe(fingerprint(b));
  const saved = [a, b].map((e) => {
    const r = readPptx(e.save()!);
    if (!r.ok) throw new Error('reopen failed');
    expect(semanticDigest(r.pkg)).toBe(canonicalPackageFingerprint(r.pkg));
    return canonicalPackageFingerprint(r.pkg);
  });
  expect(saved[0]).toBe(saved[1]);
  const xml = serializeXmlPart(xmlPartOf(a.session!.store.package, SLIDE)!);
  if (fresh) expect(xml).not.toContain('stale');
  return xml;
}

describe('a shape in mc:AlternateContent on two replicas', () => {
  test('one peer moves the drawn copy while the other types in it', async () => {
    const { a, b, l } = await pair();
    l.pause();
    a.setSelection({ kind: 'shapes', nodeIds: [boxId(a)!] });
    expect(a.exec({ type: 'nudge', dx: IN, dy: 0 }).ok).toBe(true);
    const at = { paragraph: 0, offset: 0 };
    b.setSelection({ kind: 'text', nodeId: boxId(b)!, anchor: at, focus: at });
    for (const ch of 'New ') expect(b.input('text', ch).ok).toBe(true);
    l.flush();
    const xml = expectConvergedAndSaved(a, b);
    expect(xml).toContain('New drawn');
    expect(xml).toContain(`x="${2 * IN}"`);
    expect(xml).toContain('<mc:Fallback/>');
  });

  test('one peer fills and the other moves; undo and redo on one converge', async () => {
    const { a, b, l } = await pair();
    l.pause();
    a.setSelection({ kind: 'shapes', nodeIds: [boxId(a)!] });
    expect(a.exec({ type: 'setShapeFill', color: '0000FF' }).ok).toBe(true);
    b.setSelection({ kind: 'shapes', nodeIds: [boxId(b)!] });
    expect(b.exec({ type: 'nudge', dx: 0, dy: IN }).ok).toBe(true);
    l.flush();
    expectConvergedAndSaved(a, b);
    // Undo on `a` takes back its own step: its fill and its removal of the copy. The copy then
    // comes back while the move of `b` stays, as the fidelity page lists. Both replicas agree.
    expect(a.exec({ type: 'undo' }).ok).toBe(true);
    l.flush();
    const undone = expectConvergedAndSaved(a, b, false);
    expect(undone).not.toContain('0000FF');
    expect(undone).toContain(`y="${2 * IN}"`);
    expect(a.exec({ type: 'redo' }).ok).toBe(true);
    l.flush();
    expect(expectConvergedAndSaved(a, b)).toContain('0000FF');
  });

  test('one peer deletes the shape while the other edits it', async () => {
    const { a, b, l } = await pair();
    l.pause();
    a.setSelection({ kind: 'shapes', nodeIds: [boxId(a)!] });
    expect(a.exec({ type: 'delete' }).ok).toBe(true);
    b.setSelection({ kind: 'shapes', nodeIds: [boxId(b)!] });
    expect(b.exec({ type: 'setShapeFill', color: '0000FF' }).ok).toBe(true);
    l.flush();
    expectConvergedAndSaved(a, b);
    expect(boxId(a)).toBeUndefined();
    expect(boxId(b)).toBeUndefined();
  });
});
