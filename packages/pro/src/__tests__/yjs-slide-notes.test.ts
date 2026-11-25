// The automation `setSlideNotes` operation, which `Slide.speakerNotes` of the object model sends,
// through two editors on two linked Y.Docs (`AGENTS.md` "Collaboration verification"). The
// write edits the body placeholder of a notes slide (`p:notes`, ECMA-376 Part 1, 19.3.1.26) or
// creates the Notes Slide part (13.3.5); both replicas must converge.

import { describe, expect, test } from 'bun:test';
import * as Y from 'yjs';
import type { AutomationHost, AutomationOperation } from '@pptx-editor.dev/core/automation';
import {
  createBrowserAutomationHost,
  createPptxEditor,
  notesText,
  type PptxEditorInstance,
} from '@pptx-editor.dev/core/editor';
import {
  canonicalPackageFingerprint,
  presentationInfo,
  readPptx,
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

/** Two slides with a text box each; the first has notes, the second has none. */
const deck = (): Uint8Array =>
  buildPptx({
    slides: [1, 2].map((n) => ({
      layout: 'blank' as const,
      notes: n === 1 ? 'First notes' : undefined,
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

function send(host: AutomationHost, operations: AutomationOperation[]): void {
  const response = host.execute({ operations });
  if (!response.ok) throw new Error(JSON.stringify(response.results));
}

/**
 * Set the notes of the slide at `index` with the operations `Slide.speakerNotes` sends: the
 * lookups, then the write. `packages/editor-api/src/__tests__/slide-speaker-notes.test.ts`
 * checks that it sends these.
 */
function setNotes(host: AutomationHost, index: number, text: string): void {
  send(host, [
    { op: 'getPresentation' },
    { op: 'getSlide', key: { index } },
    { op: 'setSlideNotes', slide: { result: 1 }, text },
  ]);
}

/** Set the text of the first shape of the slide at `index`, as `TextRange.text` does. */
function setShapeText(host: AutomationHost, index: number, text: string): void {
  send(host, [
    { op: 'getSlide', key: { index } },
    { op: 'getShape', container: { result: 0 }, key: { index: 0 } },
    { op: 'setText', shape: { result: 1 }, text },
  ]);
}

const pkgOf = (e: PptxEditorInstance) => e.session!.store.package;
const fingerprint = (e: PptxEditorInstance): string => canonicalPackageFingerprint(pkgOf(e));
const notesOf = (e: PptxEditorInstance): string[] =>
  presentationInfo(pkgOf(e)).slides.map((_, i) => notesText(pkgOf(e), i));

async function pair() {
  const docA = new Y.Doc();
  const docB = new Y.Doc();
  const sync = link(docA, docB);
  return { sync, a: await replica(docA, 'a'), b: await replica(docB, 'b') };
}

/** Each peer saves a package that reads back to the same canonical package. */
function sameSave(a: PptxEditorInstance, b: PptxEditorInstance): void {
  const reopened = [a, b].map((e) => {
    const r = readPptx(e.save()!);
    if (!r.ok) throw new Error('reopen');
    return canonicalPackageFingerprint(r.pkg);
  });
  expect(reopened[0]).toBe(reopened[1]!);
}

describe('Slide.speakerNotes on two peers', () => {
  test('an edit and a new notes slide on one peer reach the other, and both save the same deck', async () => {
    const { a, b } = await pair();
    setNotes(createBrowserAutomationHost(b), 0, 'First notes, edited');
    setNotes(createBrowserAutomationHost(b), 1, 'New notes');
    expect(notesOf(a)).toEqual(['First notes, edited', 'New notes']);
    expect(fingerprint(a)).toBe(fingerprint(b));
    sameSave(a, b);
  });

  test('a notes write converges while the peer edits the same slide', async () => {
    const { sync, a, b } = await pair();
    sync.pause();
    setNotes(createBrowserAutomationHost(a), 1, 'Talk track');
    setShapeText(createBrowserAutomationHost(b), 1, 'Edited box');
    setNotes(createBrowserAutomationHost(b), 0, 'Peer notes');
    sync.flush();
    expect(fingerprint(a)).toBe(fingerprint(b));
    expect(notesOf(a)).toEqual(['Peer notes', 'Talk track']);
    sameSave(a, b);
  });

  test('concurrent writes to the notes of one slide converge to one state', async () => {
    const { sync, a, b } = await pair();
    sync.pause();
    setNotes(createBrowserAutomationHost(a), 0, 'First notes from a');
    setNotes(createBrowserAutomationHost(b), 0, 'b: First notes');
    sync.flush();
    expect(fingerprint(a)).toBe(fingerprint(b));
    expect(notesOf(a)).toEqual(notesOf(b));
    sameSave(a, b);
  });

  // Both peers converge, but each creates its own Notes Slide part, so the slide keeps two
  // notes relationships (13.3.5 asks for one). The editor's notes command does the same; a
  // repair rule for it is a follow-up.
  test('concurrent creation of the notes slide of one slide converges to one state', async () => {
    const { sync, a, b } = await pair();
    sync.pause();
    setNotes(createBrowserAutomationHost(a), 1, 'From a');
    setNotes(createBrowserAutomationHost(b), 1, 'From b');
    sync.flush();
    expect(fingerprint(a)).toBe(fingerprint(b));
    expect(notesOf(a)).toEqual(notesOf(b));
    sameSave(a, b);
    const saved = readPptx(a.save()!);
    if (!saved.ok) throw new Error('reopen');
    expect(notesText(saved.pkg, 1)).toBe(notesOf(a)[1]!);
  });

  test('undo of the write on one peer restores the notes on both', async () => {
    const { a, b } = await pair();
    setNotes(createBrowserAutomationHost(b), 1, 'New notes');
    expect(b.exec({ type: 'undo' }).ok).toBe(true);
    expect(notesOf(a)).toEqual(['First notes', '']);
    expect(fingerprint(a)).toBe(fingerprint(b));
  });
});
