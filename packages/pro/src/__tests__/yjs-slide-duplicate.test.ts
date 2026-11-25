/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// The automation `duplicateSlide` operation, which `Slide.duplicate()` of the editor API sends,
// through two editors on two linked Y.Docs (`AGENTS.md`, "Verify collaboration convergence").
// The copy is a new `p:sldId` in `p:sldIdLst` (ECMA-376 Part 1, 19.2.1.34) and a new Slide part
// with its own Notes Slide part. Both replicas must hold the same package after a copy, after a
// copy while the other peer edits the source slide, and after two concurrent copies.

import { describe, expect, test } from 'bun:test';
import * as Y from 'yjs';
import type { AutomationHost, AutomationBatchRequest } from '@pptx-editor.dev/core/automation';
import {
  createBrowserAutomationHost,
  createPptxEditor,
  type PptxEditorInstance,
} from '@pptx-editor.dev/core/editor';
import {
  canonicalPackageFingerprint,
  presentationInfo,
  readPptx,
  type PptxPackage,
} from '@pptx-editor.dev/core/store';
import { buildPptx, para, run, shape } from '../../../../scripts/fixtures/builder.ts';
import { collaborationModule, createYjsCollaborationSession } from '../index.ts';

const IN = 914_400;

/** Two slides. Slide 1 has a text box `Box` (id 2) and speaker notes. */
const deck = (): Uint8Array =>
  buildPptx({
    slides: [
      {
        shapes: shape({
          id: 2,
          name: 'Box',
          x: IN,
          y: IN,
          cx: 3 * IN,
          cy: IN,
          text: para(run('Source')),
        }),
        notes: 'Notes of the source',
      },
      { layout: 'blank', shapes: '' },
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
  const sync = link(docA, docB);
  const a = await replica(docA, 'a');
  const b = await replica(docB, 'b');
  return { sync, a, b };
}

function execute(host: AutomationHost, request: AutomationBatchRequest): void {
  const response = host.execute(request);
  if (!response.ok) throw new Error(JSON.stringify(response.results));
}

/** Copy slide `index`, as `slides.getItemAt(index).duplicate()` does in one sync. */
function duplicate(editor: PptxEditorInstance, index: number): void {
  execute(createBrowserAutomationHost(editor), {
    operations: [
      { op: 'getSlide', key: { index } },
      { op: 'duplicateSlide', slide: { result: 0 } },
    ],
  });
}

/** Replace the text of `Box` on slide `index`. */
function setText(editor: PptxEditorInstance, index: number, text: string): void {
  execute(createBrowserAutomationHost(editor), {
    operations: [
      { op: 'getSlide', key: { index } },
      { op: 'getShape', container: { result: 0 }, key: { id: '2' } },
      { op: 'setText', shape: { result: 1 }, text },
    ],
  });
}

const pkgOf = (e: PptxEditorInstance): PptxPackage => e.session!.store.package;

/** Both replicas hold the same package, and both save and reopen to the same package. */
function expectConverged(a: PptxEditorInstance, b: PptxEditorInstance): void {
  expect(canonicalPackageFingerprint(pkgOf(a))).toBe(canonicalPackageFingerprint(pkgOf(b)));
  const reopened = [a, b].map((e) => {
    const r = readPptx(e.save()!);
    if (!r.ok) throw new Error('the saved presentation does not read back');
    return canonicalPackageFingerprint(r.pkg);
  });
  expect(reopened[0]).toBe(reopened[1]!);
}

/** Each slide's text in `Box`, or `-` without one, and whether it has its own notes part. */
function slides(e: PptxEditorInstance): string[] {
  const pkg = pkgOf(e);
  const decode = (name: string): string => {
    const part = pkg.parts.get(name);
    if (!part) return '';
    return part.kind === 'xml'
      ? JSON.stringify(part.xml.root)
      : new TextDecoder().decode(part.bytes);
  };
  const info = presentationInfo(pkg);
  const notes = new Set(info.slides.map((s) => s.notesPart));
  expect(notes.size).toBe(info.slides.length);
  return info.slides.map((s) => {
    const xml = decode(s.partName);
    const text = ['Source', 'Edited'].find((t) => xml.includes(t)) ?? '-';
    return `${text}${s.notesPart ? '+notes' : ''}`;
  });
}

describe('Slide.duplicate on two peers', () => {
  test('a copy reaches the other peer, and the undo of it too', async () => {
    const { a, b } = await pair();
    duplicate(a, 0);
    expectConverged(a, b);
    expect(slides(b)).toEqual(['Source+notes', 'Source+notes', '-']);
    expect(a.exec({ type: 'undo' }).ok).toBe(true);
    expectConverged(a, b);
    expect(slides(b)).toEqual(['Source+notes', '-']);
  });

  test('a copy while the other peer edits the source slide converges', async () => {
    const { sync, a, b } = await pair();
    sync.pause();
    duplicate(a, 0);
    setText(b, 0, 'Edited');
    sync.flush();
    expectConverged(a, b);
    // The copy holds the slide as the copier saw it; the edit reaches the source.
    expect(slides(a)).toEqual(['Edited+notes', 'Source+notes', '-']);
    // The copier's undo removes only the copy and keeps the peer's edit.
    expect(a.exec({ type: 'undo' }).ok).toBe(true);
    expectConverged(a, b);
    expect(slides(a)).toEqual(['Edited+notes', '-']);
  });

  test('two concurrent copies of one slide both stay', async () => {
    const { sync, a, b } = await pair();
    sync.pause();
    duplicate(a, 0);
    duplicate(b, 0);
    sync.flush();
    expectConverged(a, b);
    expect(slides(a)).toEqual(['Source+notes', 'Source+notes', 'Source+notes', '-']);
  });
});
