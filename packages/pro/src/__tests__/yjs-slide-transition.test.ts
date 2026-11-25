// The automation `setTransition` operation, which `Slide.transition` of the object model sends,
// through two editors on two linked Y.Docs (`AGENTS.md` "Collaboration verification"). The write
// is `p:transition` (`CT_SlideTransition`, ECMA-376 Part 1, §19.3.1.50) in its
// markup-compatibility branches (Part 3, §7.5 `AlternateContent`); both replicas must converge.

import { describe, expect, test } from 'bun:test';
import * as Y from 'yjs';
import type {
  AutomationHost,
  AutomationOperation,
  AutomationTransitionWrite,
} from '@pptx-editor.dev/core/automation';
import {
  createBrowserAutomationHost,
  createPptxEditor,
  type PptxEditorInstance,
} from '@pptx-editor.dev/core/editor';
import {
  canonicalPackageFingerprint,
  presentationInfo,
  readPptx,
  serializeXmlPart,
  xmlPartOf,
} from '@pptx-editor.dev/core/store';
import { buildPptx, para, run, shape } from '../../../../scripts/fixtures/builder.ts';
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

/** Two slides with one text box each. */
const deck = (): Uint8Array =>
  buildPptx({
    slides: [1, 2].map((n) => ({
      layout: 'blank' as const,
      shapes: shape({
        id: 2,
        name: 'Box',
        x: 914_400,
        y: 914_400,
        cx: 914_400,
        cy: 914_400,
        text: para(run(`Slide ${n}`)),
      }),
    })),
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

function execute(host: AutomationHost, operations: AutomationOperation[]): void {
  const response = host.execute({ operations });
  if (!response.ok) throw new Error(JSON.stringify(response.results));
}

/**
 * Set the transition of the slide at `index` with the operations `Slide.transition` sends: the
 * lookups, then the write. `packages/editor-api/src/__tests__/slide-transition.test.ts` checks
 * that it sends these.
 */
function setTransition(
  host: AutomationHost,
  index: number,
  transition: AutomationTransitionWrite
): void {
  execute(host, [
    { op: 'getPresentation' },
    { op: 'getSlide', key: { index } },
    { op: 'setTransition', slide: { result: 1 }, transition },
  ]);
}

/** Replace the text of the box on the slide at `index`, as `TextRange.text` does. */
function setText(host: AutomationHost, index: number, text: string): void {
  execute(host, [
    { op: 'getPresentation' },
    { op: 'getSlide', key: { index } },
    { op: 'getShape', container: { result: 1 }, key: { id: '2' } },
    { op: 'setText', shape: { result: 2 }, text },
  ]);
}

const fingerprint = (e: PptxEditorInstance): string =>
  canonicalPackageFingerprint(e.session!.store.package);
const transitionOf = (e: PptxEditorInstance, index: number) => {
  const host = createBrowserAutomationHost(e);
  const response = host.execute({
    operations: [
      { op: 'getPresentation' },
      { op: 'getSlide', key: { index } },
      { op: 'readTransition', slide: { result: 1 } },
    ],
  });
  const last = response.results[2];
  if (last?.status !== 'ok' || last.value.kind !== 'transition') throw new Error('no read');
  return last.value.transition;
};

const textOf = (e: PptxEditorInstance, index: number): string => {
  const response = createBrowserAutomationHost(e).execute({
    operations: [
      { op: 'getPresentation' },
      { op: 'getSlide', key: { index } },
      { op: 'getShape', container: { result: 1 }, key: { id: '2' } },
      { op: 'readText', shape: { result: 2 } },
    ],
  });
  const last = response.results[3];
  if (last?.status !== 'ok' || last.value.kind !== 'text') throw new Error('no text');
  return last.value.text;
};

/** The saved XML of the slide at `index`. */
const slideXml = (e: PptxEditorInstance, index: number): string => {
  const r = readPptx(e.save()!);
  if (!r.ok) throw new Error('reopen');
  return serializeXmlPart(xmlPartOf(r.pkg, presentationInfo(r.pkg).slides[index]!.partName)!);
};

async function pair() {
  const docA = new Y.Doc();
  const docB = new Y.Doc();
  const sync = link(docA, docB);
  return { sync, a: await replica(docA, 'a'), b: await replica(docB, 'b') };
}

describe('Slide.transition on two peers', () => {
  test('a change on one peer reaches the other, and both save the same deck', async () => {
    const { a, b } = await pair();
    setTransition(createBrowserAutomationHost(b), 0, { kind: 'fade', durationMs: 1200 });
    setTransition(createBrowserAutomationHost(b), 1, { kind: 'push', advanceAfterMs: 3000 });
    expect(transitionOf(a, 0)).toMatchObject({ kind: 'fade', durationMs: 1200 });
    expect(transitionOf(a, 1)).toMatchObject({ kind: 'push', advanceAfterMs: 3000 });
    expect(fingerprint(a)).toBe(fingerprint(b));
    const reopened = [a, b].map((e) => {
      const r = readPptx(e.save()!);
      if (!r.ok) throw new Error('reopen');
      return canonicalPackageFingerprint(r.pkg);
    });
    expect(reopened[0]).toBe(reopened[1]!);
  });

  test('a transition write and a text edit on the same slide keep both changes', async () => {
    const { sync, a, b } = await pair();
    sync.pause();
    setTransition(createBrowserAutomationHost(a), 0, { kind: 'wipe', durationMs: 900 });
    setText(createBrowserAutomationHost(b), 0, 'Edited');
    sync.flush();
    expect(fingerprint(a)).toBe(fingerprint(b));
    for (const e of [a, b]) {
      expect(transitionOf(e, 0)).toMatchObject({ kind: 'wipe', durationMs: 900 });
      expect(textOf(e, 0)).toBe('Edited');
    }
  });

  test('concurrent transitions on one slide converge to one state', async () => {
    const { sync, a, b } = await pair();
    sync.pause();
    setTransition(createBrowserAutomationHost(a), 0, { kind: 'fade', durationMs: 700 });
    setTransition(createBrowserAutomationHost(b), 0, { kind: 'cut', advanceAfterMs: 2000 });
    sync.flush();
    expect(fingerprint(a)).toBe(fingerprint(b));
    expect(transitionOf(a, 0)).toEqual(transitionOf(b, 0));
    // `CT_Slide` holds at most one transition: one `mc:AlternateContent` holds both copies.
    const xml = slideXml(a, 0);
    expect(xml.match(/<mc:AlternateContent/g)).toHaveLength(1);
    expect(xml.match(/<p:transition /g)).toHaveLength(2);
  });

  test('undo of the change on one peer restores the state on both', async () => {
    const { a, b } = await pair();
    setTransition(createBrowserAutomationHost(b), 0, { kind: 'split', durationMs: 1500 });
    expect(b.exec({ type: 'undo' }).ok).toBe(true);
    expect(transitionOf(a, 0).kind).toBe('none');
    expect(fingerprint(a)).toBe(fingerprint(b));
  });
});
