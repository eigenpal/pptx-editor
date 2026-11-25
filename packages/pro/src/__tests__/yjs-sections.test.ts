// The automation section operations, which `Presentation.sections` of the object model sends,
// through two editors on two linked Y.Docs (`AGENTS.md` "Collaboration verification"). The
// section list is a `p14:sectionLst` in a `p:ext` of the presentation's `p:extLst` (ECMA-376
// Part 1, 19.2.1.11 and 19.2.1.12); both replicas must converge.

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

const P14 = 'http://schemas.microsoft.com/office/powerpoint/2010/main';

/** Four slides: section Intro holds 256 and 257, section Body holds 258 and 259. */
const deck = (): Uint8Array =>
  buildPptx({
    slides: [1, 2, 3, 4].map(() => ({ layout: 'blank' as const, shapes: '' })),
    presentationExt:
      `<p:extLst><p:ext uri="{521415D9-36F7-43E2-AB2F-B90AF26B5E84}"><p14:sectionLst xmlns:p14="${P14}">` +
      '<p14:section name="Intro" id="{6E1B33A1-4C0F-4B53-9C2B-0D4A4B3E8A11}"><p14:sldIdLst><p14:sldId id="256"/><p14:sldId id="257"/></p14:sldIdLst></p14:section>' +
      '<p14:section name="Body" id="{0B7F5C2D-1E4A-4F3B-8C6D-2A9E7B1C5D22}"><p14:sldIdLst><p14:sldId id="258"/><p14:sldId id="259"/></p14:sldIdLst></p14:section>' +
      '</p14:sectionLst></p:ext></p:extLst>',
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

function run(editor: PptxEditorInstance, operations: AutomationOperation[]): void {
  const host: AutomationHost = createBrowserAutomationHost(editor);
  const response = host.execute({ operations });
  if (!response.ok) throw new Error(JSON.stringify(response.results));
}

// The operations of the object model calls, as `packages/editor-api/src/__tests__/sections.test.ts`
// checks that they are sent.
const at = (index: number): AutomationOperation[] => [
  { op: 'getPresentation' },
  { op: 'getSection', key: { index } },
];
/** `sections.add(name, slideId)`. */
const add = (e: PptxEditorInstance, name: string, slideId: string) =>
  run(e, [
    { op: 'getPresentation' },
    { op: 'getSlide', key: { id: slideId } },
    { op: 'addSection', name, slide: { result: 1 } },
  ]);
/** `sections.getItemAt(index).name = name`. */
const rename = (e: PptxEditorInstance, index: number, name: string) =>
  run(e, [...at(index), { op: 'renameSection', section: { result: 1 }, name }]);
/** `sections.getItemAt(index).moveTo(to)`. */
const move = (e: PptxEditorInstance, index: number, to: number) =>
  run(e, [...at(index), { op: 'moveSection', section: { result: 1 }, index: to }]);
/** `sections.getItemAt(index).delete({ deleteSlides })`. */
const remove = (e: PptxEditorInstance, index: number, deleteSlides: boolean) =>
  run(e, [
    ...at(index),
    { op: 'deleteSection', section: { result: 1 }, ...(deleteSlides ? { deleteSlides } : {}) },
  ]);
/** `slides.getItem(slideId).hidden = true`, a peer edit of the same slide. */
const hide = (e: PptxEditorInstance, slideId: string) =>
  run(e, [
    { op: 'getPresentation' },
    { op: 'getSlide', key: { id: slideId } },
    { op: 'setSlideHidden', slide: { result: 1 }, hidden: true },
  ]);
/** `slides.getItem(slideId).moveTo(index)`. */
const moveSlide = (e: PptxEditorInstance, slideId: string, index: number) =>
  run(e, [
    { op: 'getPresentation' },
    { op: 'getSlide', key: { id: slideId } },
    { op: 'moveSlide', slide: { result: 1 }, index },
  ]);

const fingerprint = (e: PptxEditorInstance): string =>
  canonicalPackageFingerprint(e.session!.store.package);

/** The sections the slide pane shows: name, first slide, and slide count. */
const sectionsOf = (e: PptxEditorInstance): string[] =>
  e.snapshot().sections!.map((s) => `${s.name}@${s.firstSlide}+${s.slideCount}`);

/** The same, read from the saved bytes, so the writer's repair agrees too. */
function savedFingerprint(e: PptxEditorInstance): string {
  const r = readPptx(e.save()!);
  if (!r.ok) throw new Error('reopen');
  return canonicalPackageFingerprint(r.pkg);
}

async function pair() {
  const docA = new Y.Doc();
  const docB = new Y.Doc();
  const sync = link(docA, docB);
  return { sync, a: await replica(docA, 'a'), b: await replica(docB, 'b') };
}

function converged(a: PptxEditorInstance, b: PptxEditorInstance): void {
  expect(fingerprint(a)).toBe(fingerprint(b));
  expect(sectionsOf(a)).toEqual(sectionsOf(b));
  expect(savedFingerprint(a)).toBe(savedFingerprint(b));
}

describe('Presentation.sections on two peers', () => {
  test('each edit on one peer reaches the other, and both save the same deck', async () => {
    const { a, b } = await pair();
    add(b, 'Close', '259');
    rename(b, 2, 'Closing');
    move(b, 2, 0);
    expect(sectionsOf(a)).toEqual(['Closing@0+1', 'Intro@1+2', 'Body@3+1']);
    remove(b, 0, true);
    expect(sectionsOf(a)).toEqual(['Intro@0+2', 'Body@2+1']);
    converged(a, b);
  });

  test('a section added while a peer edits the slide it starts at keeps both edits', async () => {
    const { sync, a, b } = await pair();
    sync.pause();
    add(a, 'Middle', '257');
    hide(b, '257');
    sync.flush();
    converged(a, b);
    expect(sectionsOf(a)).toEqual(['Intro@0+1', 'Middle@1+1', 'Body@2+2']);
    expect(a.snapshot().slides.map((s) => s.hidden)).toEqual([false, true, false, false]);
  });

  test('a section moved while a peer moves one of its slides converges', async () => {
    const { sync, a, b } = await pair();
    sync.pause();
    move(a, 1, 0);
    moveSlide(b, '258', 0);
    sync.flush();
    converged(a, b);
    // Every slide is in exactly one section.
    expect(a.snapshot().sections!.reduce((n, s) => n + s.slideCount, 0)).toBe(4);
  });

  test('a section renamed on one peer and deleted on the other converges', async () => {
    const { sync, a, b } = await pair();
    sync.pause();
    rename(a, 1, 'Main');
    remove(b, 1, false);
    sync.flush();
    converged(a, b);
  });

  test('two peers that add a section at the same time keep both sections', async () => {
    const { sync, a, b } = await pair();
    sync.pause();
    add(a, 'From A', '257');
    add(b, 'From B', '259');
    sync.flush();
    converged(a, b);
    expect(
      a
        .snapshot()
        .sections!.map((s) => s.name)
        .sort()
    ).toEqual(['Body', 'From A', 'From B', 'Intro'].sort());
  });

  test('undo of a section edit on one peer restores the sections on both', async () => {
    const { a, b } = await pair();
    move(b, 1, 0);
    expect(b.exec({ type: 'undo' }).ok).toBe(true);
    expect(sectionsOf(a)).toEqual(['Intro@0+2', 'Body@2+2']);
    converged(a, b);
  });
});
