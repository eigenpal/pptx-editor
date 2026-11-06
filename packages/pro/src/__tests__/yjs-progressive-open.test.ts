import { afterAll, describe, expect, test } from 'bun:test';
import { GlobalWindow } from 'happy-dom';
import * as Y from 'yjs';
import { createPptxEditor, type PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import { canonicalPackageFingerprint, isDeferredPending } from '@pptx-editor.dev/core/store';
import { textDeck } from '../../../../scripts/bench/text-deck.ts';
import { collaborationModule, createYjsCollaborationSession } from '../index.ts';

// A long deck in a collaboration room. A room needs every part, so an editor with a
// collaboration module opens the deck whole: no slide is left to parse when the room takes
// the package, and edits on any slide converge. A peer without the module opens the same deck
// progressively, and its save matches the room's.

const SLIDES = 300;
const BYTES = textDeck(SLIDES);

// A DOM for this file only: the test run registers no global DOM.
const win = new GlobalWindow();
afterAll(() => win.happyDOM.close());

async function replica(ydoc: Y.Doc, actorId: string): Promise<PptxEditorInstance> {
  const session = createYjsCollaborationSession({
    doc: ydoc,
    documentId: 'deck',
    identity: { actorId, name: actorId },
  });
  const editor = createPptxEditor({
    fixedMeasurer: true,
    substitutes: false,
    modules: [collaborationModule({ session })],
  });
  expect((await editor.load(BYTES)).ok).toBe(true);
  return editor;
}

const fingerprint = (e: PptxEditorInstance) =>
  canonicalPackageFingerprint(e.session!.store.package);

const pending = (e: PptxEditorInstance): number =>
  e.session!.info.slides.filter((s) =>
    isDeferredPending(e.session!.store.package.parts.get(s.partName))
  ).length;

function typeAt(editor: PptxEditorInstance, index: number, text: string): void {
  editor.goToSlide(index);
  const item = editor.session!.layout(index)!.items.find((i) => i.element.name === 'Body')!;
  const end = item.kind === 'shape' ? (item.text?.paragraphLengths[0] ?? 0) : 0;
  const at = { paragraph: 0, offset: end };
  editor.setSelection({ kind: 'text', nodeId: item.element.nodeId, anchor: at, focus: at });
  for (const ch of text) editor.input('text', ch);
}

describe('a long deck in a collaboration room', () => {
  test('opens whole, and edits on far slides converge', async () => {
    const docA = new Y.Doc();
    const docB = new Y.Doc();
    docA.on('update', (u: Uint8Array, origin: unknown) => {
      if (origin !== 'link') Y.applyUpdate(docB, u, 'link');
    });
    docB.on('update', (u: Uint8Array, origin: unknown) => {
      if (origin !== 'link') Y.applyUpdate(docA, u, 'link');
    });
    const a = await replica(docA, 'a');
    const b = await replica(docB, 'b');
    expect(a.snapshot().pendingSlides ?? 0).toBe(0);
    expect(b.snapshot().pendingSlides ?? 0).toBe(0);
    expect(pending(a)).toBe(0);
    expect(pending(b)).toBe(0);
    expect(fingerprint(a)).toBe(fingerprint(b));

    typeAt(a, 0, 'first');
    typeAt(b, 250, 'far');
    a.exec({ type: 'deleteSlide', index: 120 });
    expect(fingerprint(a)).toBe(fingerprint(b));
    expect(a.snapshot().slides.length).toBe(SLIDES - 1);

    // A peer with no room opens the deck progressively, makes the same edits before its open
    // ends, and saves what the room holds.
    const solo = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    expect((await solo.load(BYTES)).ok).toBe(true);
    expect(solo.snapshot().pendingSlides).toBeGreaterThan(0);
    typeAt(solo, 0, 'first');
    typeAt(solo, 250, 'far');
    solo.exec({ type: 'deleteSlide', index: 120 });
    expect(await reopened(solo.save()!)).toBe(await reopened(a.save()!));
    for (const e of [a, b, solo]) e.destroy();
  });
});

/** The canonical fingerprint of the deck `bytes` hold, opened whole. */
async function reopened(bytes: Uint8Array): Promise<string> {
  const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await editor.load(bytes)).ok).toBe(true);
  editor.save(); // Parses whatever is still opening.
  const out = fingerprint(editor);
  editor.destroy();
  return out;
}
