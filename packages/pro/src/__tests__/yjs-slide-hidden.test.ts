// The automation `setSlideHidden` operation, which `Slide.hidden` of the object model sends,
// through two editors on two linked Y.Docs (`AGENTS.md` "Collaboration verification"). The
// write is `p:sld/@show` (ECMA-376 Part 1, §19.3.1.38); both replicas must converge.

import { describe, expect, test } from 'bun:test';
import * as Y from 'yjs';
import type { AutomationHost, AutomationOperation } from '@pptx-editor.dev/core/automation';
import {
  createBrowserAutomationHost,
  createPptxEditor,
  type PptxEditorInstance,
} from '@pptx-editor.dev/core/editor';
import { canonicalPackageFingerprint, readPptx } from '@pptx-editor.dev/core/store';
import { buildPptx } from '../../../../scripts/fixtures/builder.ts';
import { collaborationModule, createYjsCollaborationSession } from '../index.ts';

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

/** Three slides; the third is hidden in the file. */
const deck = (): Uint8Array =>
  buildPptx({
    slides: [1, 2, 3].map((n) => ({ layout: 'blank' as const, hidden: n === 3, shapes: '' })),
  });

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

/**
 * Hide or show the slide at `index` with the operations `Slide.hidden` sends: the lookups, then
 * the write. `packages/editor-api/src/__tests__/slide-hidden.test.ts` checks that it sends these.
 */
function setHidden(host: AutomationHost, index: number, hidden: boolean): void {
  const operations: AutomationOperation[] = [
    { op: 'getPresentation' },
    { op: 'getSlide', key: { index } },
    { op: 'setSlideHidden', slide: { result: 1 }, hidden },
  ];
  const response = host.execute({ operations });
  if (!response.ok) throw new Error(JSON.stringify(response.results));
}

const fingerprint = (e: PptxEditorInstance): string =>
  canonicalPackageFingerprint(e.session!.store.package);
const hiddenOf = (e: PptxEditorInstance): boolean[] => e.snapshot().slides.map((s) => s.hidden);

async function pair() {
  const docA = new Y.Doc();
  const docB = new Y.Doc();
  const sync = link(docA, docB);
  return { sync, a: await replica(docA, 'a'), b: await replica(docB, 'b') };
}

describe('Slide.hidden on two peers', () => {
  test('a change on one peer reaches the other, and both save the same deck', async () => {
    const { a, b } = await pair();
    setHidden(createBrowserAutomationHost(b), 0, true);
    setHidden(createBrowserAutomationHost(b), 2, false);
    expect(hiddenOf(a)).toEqual([true, false, false]);
    expect(fingerprint(a)).toBe(fingerprint(b));
    const reopened = [a, b].map((e) => {
      const r = readPptx(e.save()!);
      if (!r.ok) throw new Error('reopen');
      return canonicalPackageFingerprint(r.pkg);
    });
    expect(reopened[0]).toBe(reopened[1]!);
  });

  test('concurrent changes to different slides keep both changes', async () => {
    const { sync, a, b } = await pair();
    sync.pause();
    setHidden(createBrowserAutomationHost(a), 0, true);
    setHidden(createBrowserAutomationHost(b), 1, true);
    sync.flush();
    expect(fingerprint(a)).toBe(fingerprint(b));
    expect(hiddenOf(a)).toEqual([true, true, true]);
  });

  test('concurrent opposite changes to one slide converge to one state', async () => {
    const { sync, a, b } = await pair();
    sync.pause();
    setHidden(createBrowserAutomationHost(a), 2, false);
    setHidden(createBrowserAutomationHost(b), 1, true);
    setHidden(createBrowserAutomationHost(b), 2, true);
    sync.flush();
    expect(fingerprint(a)).toBe(fingerprint(b));
    expect(hiddenOf(a)).toEqual(hiddenOf(b));
    expect(hiddenOf(a)[1]).toBe(true);
  });

  test('undo of the change on one peer restores the state on both', async () => {
    const { a, b } = await pair();
    setHidden(createBrowserAutomationHost(b), 0, true);
    expect(b.exec({ type: 'undo' }).ok).toBe(true);
    expect(hiddenOf(a)).toEqual([false, false, true]);
    expect(fingerprint(a)).toBe(fingerprint(b));
  });
});
