// The automation `setHeaderFooter` operation, which `Slide.headerFooter` and
// `Presentation.applyHeaderFooter` of the object model send, through two editors on two linked
// Y.Docs (`AGENTS.md` "Collaboration verification"). The write adds or removes the slide's own
// `sldNum`, `ftr`, and `dt` placeholders (`p:ph`, ECMA-376 Part 1, 19.3.1.36); both replicas
// must converge, also while a peer edits the same slide.

import { describe, expect, test } from 'bun:test';
import * as Y from 'yjs';
import type { AutomationHost, AutomationOperation } from '@pptx-editor.dev/core/automation';
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
import { buildPptx } from '../../../../scripts/fixtures/builder.ts';
import { headerFooterDeck } from '../../../../scripts/fixtures/round-trip.ts';
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
  expect((await editor.load(buildPptx(headerFooterDeck()))).ok).toBe(true);
  return editor;
}

function run(host: AutomationHost, operations: AutomationOperation[]): void {
  const response = host.execute({ operations });
  if (!response.ok) throw new Error(JSON.stringify(response.results));
}

/**
 * The operations that `Slide.headerFooter` sends for slide `index`, or that
 * `Presentation.applyHeaderFooter` sends without one.
 * `packages/editor-api/src/__tests__/header-footer.test.ts` checks that the API sends these.
 */
function setHeaderFooter(
  editor: PptxEditorInstance,
  settings: Record<string, unknown>,
  index?: number
): void {
  run(
    createBrowserAutomationHost(editor),
    index === undefined
      ? [{ op: 'setHeaderFooter', settings }]
      : [
          { op: 'getPresentation' },
          { op: 'getSlide', key: { index } },
          { op: 'setHeaderFooter', slides: [{ result: 1 }], settings },
        ]
  );
}

/** Replace the text of the title of slide `index`. */
function setTitle(editor: PptxEditorInstance, index: number, text: string): void {
  run(createBrowserAutomationHost(editor), [
    { op: 'getPresentation' },
    { op: 'getSlide', key: { index } },
    { op: 'getShape', container: { result: 1 }, key: { index: 0 } },
    { op: 'setText', shape: { result: 2 }, text },
  ]);
}

const fingerprint = (e: PptxEditorInstance): string =>
  canonicalPackageFingerprint(e.session!.store.package);

const slideXml = (e: PptxEditorInstance, index: number): string => {
  const pkg = e.session!.store.package;
  return serializeXmlPart(xmlPartOf(pkg, presentationInfo(pkg).slides[index]!.partName)!);
};

async function pair() {
  const docA = new Y.Doc();
  const docB = new Y.Doc();
  const sync = link(docA, docB);
  return { sync, a: await replica(docA, 'a'), b: await replica(docB, 'b') };
}

describe('setHeaderFooter on two peers', () => {
  test('a change on one peer reaches the other, and both save the same deck', async () => {
    const { a, b } = await pair();
    setHeaderFooter(b, { slideNumberVisible: true, footerText: 'Board review' });
    expect(slideXml(a, 1)).toMatch(/type="sldNum"[^]*Board review/);
    expect(fingerprint(a)).toBe(fingerprint(b));
    const reopened = [a, b].map((e) => {
      const r = readPptx(e.save()!);
      if (!r.ok) throw new Error('reopen');
      return canonicalPackageFingerprint(r.pkg);
    });
    expect(reopened[0]).toBe(reopened[1]!);
  });

  test('converges while the peer edits text on the same slide', async () => {
    const { sync, a, b } = await pair();
    sync.pause();
    setHeaderFooter(
      a,
      { slideNumberVisible: true, footerText: 'Board review', dateText: 'June' },
      1
    );
    setTitle(b, 1, 'Edited by b');
    sync.flush();
    expect(fingerprint(a)).toBe(fingerprint(b));
    // Both edits stay: the title text and the new placeholders.
    const xml = slideXml(a, 1);
    expect(xml).toContain('Edited by b');
    expect(xml).toMatch(/type="sldNum"[^]*type="ftr"[^]*type="dt"/);
  });

  test('concurrent opposite writes to one slide converge to one state', async () => {
    const { sync, a, b } = await pair();
    sync.pause();
    setHeaderFooter(a, { footerVisible: false, slideNumberVisible: false }, 0);
    setHeaderFooter(b, { footerText: 'New text' }, 0);
    sync.flush();
    expect(fingerprint(a)).toBe(fingerprint(b));
    expect(slideXml(a, 0)).toBe(slideXml(b, 0));
  });

  test('concurrent writes to all slides on both peers converge', async () => {
    const { sync, a, b } = await pair();
    sync.pause();
    setHeaderFooter(a, { slideNumberVisible: true, footerText: 'From a' });
    setHeaderFooter(b, { dateVisible: true, footerText: 'From b' });
    sync.flush();
    expect(fingerprint(a)).toBe(fingerprint(b));
    for (const e of [a, b]) expect(readPptx(e.save()!).ok).toBe(true);
  });

  test('undo of the change on one peer restores the state on both', async () => {
    const { a, b } = await pair();
    const before = slideXml(a, 1);
    setHeaderFooter(b, { slideNumberVisible: true, dateVisible: true });
    expect(b.exec({ type: 'undo' }).ok).toBe(true);
    expect(slideXml(a, 1)).toBe(before);
    expect(fingerprint(a)).toBe(fingerprint(b));
  });
});
