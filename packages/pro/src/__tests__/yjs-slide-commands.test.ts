// Change layout, reset slide, hide slide, and notes formatting through two editors on two linked
// Y.Docs (`AGENTS.md` "Collaboration verification"): concurrent edits with the link paused,
// undo and redo, an edit inside a placeholder that the other replica removes, and save and
// reopen of both replicas to the same deck.

import { describe, expect, test } from 'bun:test';
import * as Y from 'yjs';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { createPptxEditor, type PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import {
  canonicalPackageFingerprint,
  elements,
  presentationInfo,
  readPptx,
  xmlPartOf,
  type PptxPackage,
} from '@pptx-editor.dev/core/store';
import { buildPptx } from '../../../../scripts/fixtures/builder.ts';
import { buildSlideLayoutsDeck } from '../../../../scripts/fixtures/slide-layouts.ts';
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

async function replica(
  doc: Y.Doc,
  actorId: string,
  bytes: Uint8Array = buildSlideLayoutsDeck()
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

async function pair(bytes?: Uint8Array) {
  const docA = new Y.Doc();
  const docB = new Y.Doc();
  const sync = link(docA, docB);
  return { sync, a: await replica(docA, 'a', bytes), b: await replica(docB, 'b', bytes) };
}

const fingerprint = (e: PptxEditorInstance): string =>
  canonicalPackageFingerprint(e.session!.store.package);

/** The fingerprint of the deck after save and reopen. */
function saved(e: PptxEditorInstance): string {
  const r = readPptx(e.save()!);
  if (!r.ok) throw new Error('reopen');
  return canonicalPackageFingerprint(r.pkg);
}

const layoutId = (e: PptxEditorInstance, name: string): string =>
  e.snapshot().layouts!.find((l) => l.name === name)!.id;
const layoutNames = (e: PptxEditorInstance): string[] =>
  e.snapshot().slides.map((s) => e.snapshot().layouts!.find((l) => l.id === s.layoutId)!.name);

function expectConverged(a: PptxEditorInstance, b: PptxEditorInstance): void {
  expect(fingerprint(a)).toBe(fingerprint(b));
  expect(saved(a)).toBe(saved(b));
}

describe('slide commands on two replicas', () => {
  test('concurrent change layout, hide slide, and notes format all land', async () => {
    const { sync, a, b } = await pair();
    sync.pause();
    expect(
      a.exec({ type: 'applyLayout', indices: [1], layoutId: layoutId(a, 'Two Content') }).ok
    ).toBe(true);
    expect(b.exec({ type: 'setSlidesHidden', indices: [4], hidden: true }).ok).toBe(true);
    expect(b.exec({ type: 'formatNotes', index: 1, start: 0, end: 4, format: 'bold' }).ok).toBe(
      true
    );
    sync.flush();
    expectConverged(a, b);
    expect(layoutNames(b)[1]).toBe('Two Content');
    expect(a.snapshot().slides[4]!.hidden).toBe(true);
    a.goToSlide(1);
    expect(a.snapshot().notesFormatted![0]!.runs[0]).toMatchObject({ text: 'Open', bold: true });
  });

  test('undo and redo of change layout and reset reach the other replica', async () => {
    const { a, b } = await pair();
    const before = fingerprint(b);
    a.exec({ type: 'applyLayout', indices: [0], layoutId: layoutId(a, 'Title and Content') });
    expectConverged(a, b);
    a.exec({ type: 'resetSlides', indices: [1] });
    expectConverged(a, b);
    a.exec({ type: 'undo' });
    a.exec({ type: 'undo' });
    expect(fingerprint(b)).toBe(before);
    a.exec({ type: 'redo' });
    expectConverged(a, b);
    expect(layoutNames(b)[0]).toBe('Title and Content');
  });

  test('typing in a placeholder that the other replica removes converges and saves', async () => {
    const { sync, a, b } = await pair();
    sync.pause();
    // A: slide 5's empty placeholders have no place on Blank, so they go.
    a.exec({ type: 'applyLayout', indices: [4], layoutId: layoutId(a, 'Blank') });
    // B: types into slide 5's empty title at the same time.
    b.goToSlide(4);
    const title = b.session!.layout(4)!.items.find((i) => i.element.name === 'Title 1')!;
    const caret = { paragraph: 0, offset: 0 };
    b.setSelection({ kind: 'text', nodeId: title.element.nodeId, anchor: caret, focus: caret });
    expect(b.input('text', 'Late title').ok).toBe(true);
    sync.flush();
    expectConverged(a, b);
    expect(layoutNames(a)[4]).toBe('Blank');
  });

  test('concurrent notes format and notes text edit both land', async () => {
    const { sync, a, b } = await pair();
    sync.pause();
    a.exec({ type: 'formatNotes', index: 1, start: 0, end: 4, format: 'italic' });
    b.goToSlide(1);
    b.exec({ type: 'setNotes', index: 1, text: `${b.snapshot().notes} Thank them.` });
    sync.flush();
    expectConverged(a, b);
    const runs = b.snapshot().notesFormatted!.flatMap((p) => p.runs);
    expect(runs[0]).toMatchObject({ text: 'Open', italic: true });
    expect(runs.map((r) => r.text).join('')).toContain('Thank them.');
  });
});

/** The saved deck of a replica, reopened. */
function reopen(e: PptxEditorInstance): PptxPackage {
  const r = readPptx(e.save()!);
  if (!r.ok) throw new Error('reopen');
  return r.pkg;
}

describe('section commands on two replicas', () => {
  const plain = (): Uint8Array =>
    buildPptx({ slides: [{ shapes: '' }, { shapes: '' }, { shapes: '' }, { shapes: '' }] });
  const sectionsOf = (e: PptxEditorInstance): string[] =>
    e.snapshot().sections!.map((s) => `${s.name}:${s.firstSlide}+${s.slideCount}`);

  test('concurrent Add Section on a deck without sections saves one valid section list', async () => {
    const { sync, a, b } = await pair(plain());
    sync.pause();
    a.exec({ type: 'addSection', name: 'A', index: 1 });
    b.exec({ type: 'addSection', name: 'B', index: 2 });
    sync.flush();
    expect(fingerprint(a)).toBe(fingerprint(b));
    expect(sectionsOf(a)).toEqual(sectionsOf(b));
    for (const e of [a, b]) {
      const pkg = reopen(e);
      const root = xmlPartOf(pkg, presentationInfo(pkg).presentationPart)!.root;
      expect(elements(root).filter((el) => el.local === 'extLst').length).toBe(1);
      // Every slide is in exactly one section after save.
      const named: string[] = [];
      const walk = (el: typeof root): void => {
        if (el.local === 'sldId' && el.ns.includes('2010'))
          named.push(el.attrs.find((x) => x.local === 'id')!.value);
        for (const c of elements(el)) walk(c);
      };
      walk(root);
      expect(new Set(named).size).toBe(named.length);
    }
    expect(saved(a)).toBe(saved(b));
  });

  test('concurrent rename and move of sections, then undo, converge', async () => {
    const { sync, a, b } = await pair();
    const id = (e: PptxEditorInstance, name: string) =>
      e.snapshot().sections!.find((s) => s.name === name)!.id;
    sync.pause();
    a.exec({ type: 'renameSection', sectionId: id(a, 'Details'), name: 'More' });
    b.exec({ type: 'moveSection', sectionId: id(b, 'Opening'), direction: 1 });
    sync.flush();
    expectConverged(a, b);
    expect(sectionsOf(a)).toEqual(['More:0+4', 'Opening:4+2']);
    b.exec({ type: 'undo' });
    expectConverged(a, b);
    expect(sectionsOf(a)).toEqual(['Opening:0+2', 'More:2+4']);
  });

  test('concurrent delete and move converge, and the rail shows the same sections', async () => {
    const { sync, a, b } = await pair();
    sync.pause();
    a.exec({ type: 'deleteSlides', indices: [5] });
    b.exec({ type: 'moveSlides', indices: [5], to: 0 });
    sync.flush();
    expectConverged(a, b);
    // The rail shows the sections that the file will hold, on both replicas.
    expect(sectionsOf(a)).toEqual(sectionsOf(b));
    const count = a.snapshot().slides.length;
    a.exec({ type: 'addSlide' });
    expectConverged(a, b);
    expect(b.snapshot().slides.length).toBe(count + 1);
    // Every saved slide entry has its own r:id.
    for (const e of [a, b]) {
      const pkg = reopen(e);
      const root = xmlPartOf(pkg, presentationInfo(pkg).presentationPart)!.root;
      const list = elements(root).find((el) => el.local === 'sldIdLst')!;
      const rIds = elements(list).map(
        (el) => el.attrs.find((x) => x.local === 'id' && x.ns)!.value
      );
      expect(new Set(rIds).size).toBe(rIds.length);
    }
  });

  // A concurrent delete and move used to bring back the deleted
  // slide's part and its presentation relationship without its content type or its rels, so the
  // dead `p:sldId` named a part and the writer repair kept it (7 entries, 6 slides). Now the
  // delete wins and the entry goes. Only the entry count is in this test; the r:id check is in
  // the test above.
  test('concurrent delete and move, then New Slide, leave no slide entry without a slide', async () => {
    const { sync, a, b } = await pair();
    sync.pause();
    a.exec({ type: 'deleteSlides', indices: [5] });
    b.exec({ type: 'moveSlides', indices: [5], to: 0 });
    sync.flush();
    a.exec({ type: 'addSlide' });
    for (const e of [a, b]) {
      const pkg = reopen(e);
      const root = xmlPartOf(pkg, presentationInfo(pkg).presentationPart)!.root;
      const list = elements(root).find((el) => el.local === 'sldIdLst')!;
      const rIds = elements(list).map(
        (el) => el.attrs.find((x) => x.local === 'id' && x.ns)!.value
      );
      expect(rIds.length).toBe(presentationInfo(pkg).slides.length);
    }
  });

  test('concurrent section move and slide move: the rail keeps its sections, in slide order', async () => {
    const { sync, a, b } = await pair();
    const details = a.snapshot().sections!.find((s) => s.name === 'Details')!.id;
    sync.pause();
    a.exec({ type: 'moveSection', sectionId: details, direction: -1 });
    b.exec({ type: 'moveSlides', indices: [5], to: 0 });
    sync.flush();
    expectConverged(a, b);
    for (const e of [a, b]) {
      const shown = e.snapshot().sections!;
      let next = 0;
      for (const s of shown) {
        expect(s.firstSlide).toBe(next);
        next += s.slideCount;
      }
      expect(next).toBe(e.snapshot().slides.length);
    }
    // A section command acts on the sections the rail shows, and both replicas agree.
    a.exec({ type: 'renameSection', sectionId: details, name: 'More' });
    expectConverged(a, b);
    expect(sectionsOf(b)).toEqual(sectionsOf(a));
  });

  test('after the repair, two concurrent section renames both land', async () => {
    const { sync, a, b } = await pair();
    const id = (e: PptxEditorInstance, name: string) =>
      e.snapshot().sections!.find((s) => s.name === name)!.id;
    sync.pause();
    a.exec({ type: 'moveSection', sectionId: id(a, 'Details'), direction: -1 });
    b.exec({ type: 'moveSlides', indices: [5], to: 0 });
    sync.flush();
    // The part needs the repair now; each rename runs it first, in place.
    sync.pause();
    a.exec({ type: 'renameSection', sectionId: id(a, 'Opening'), name: 'OpeningA' });
    b.exec({ type: 'renameSection', sectionId: id(b, 'Details'), name: 'DetailsB' });
    sync.flush();
    expectConverged(a, b);
    expect(
      a
        .snapshot()
        .sections!.map((s) => s.name)
        .sort()
    ).toEqual(['DetailsB', 'OpeningA']);
    // Both replicas moved the same entries, so each slide is named once in the merged part.
    const root = xmlPartOf(a.session!.store.package, a.session!.info.presentationPart)!.root;
    const named: string[] = [];
    const walk = (el: typeof root): void => {
      if (el.local === 'sldId' && el.ns.includes('2010'))
        named.push(el.attrs.find((x) => x.local === 'id')!.value);
      for (const c of elements(el)) walk(c);
    };
    walk(root);
    expect(new Set(named).size).toBe(named.length);
  });

  test('delete and New Slide on one replica and a move on the other save one entry per slide', async () => {
    const { sync, a, b } = await pair();
    sync.pause();
    a.exec({ type: 'deleteSlides', indices: [5] });
    a.exec({ type: 'addSlide' });
    b.exec({ type: 'moveSlides', indices: [5], to: 0 });
    sync.flush();
    expectConverged(a, b);
    for (const e of [a, b]) {
      const pkg = reopen(e);
      const root = xmlPartOf(pkg, presentationInfo(pkg).presentationPart)!.root;
      const list = elements(root).find((el) => el.local === 'sldIdLst')!;
      const ids = elements(list).map(
        (el) => el.attrs.find((x) => x.local === 'id' && !x.ns)!.value
      );
      const rIds = elements(list).map(
        (el) => el.attrs.find((x) => x.local === 'id' && x.ns)!.value
      );
      expect(new Set(ids).size).toBe(ids.length);
      expect(new Set(rIds).size).toBe(rIds.length);
    }
  });

  const addDeleteAdd = async () => {
    const { sync, a, b } = await pair();
    sync.pause();
    expect(a.exec({ type: 'addSlide' }).ok).toBe(true);
    expect(b.exec({ type: 'deleteSlides', indices: [1] }).ok).toBe(true);
    expect(b.exec({ type: 'addSlide' }).ok).toBe(true);
    expect(b.exec({ type: 'addSlide' }).ok).toBe(true);
    sync.flush();
    expectConverged(a, b);
    return [a, b] as const;
  };

  test('New Slide on one replica, Delete and two New Slides on the other: the rail shows what the file holds', async () => {
    for (const e of await addDeleteAdd()) {
      const shown = e.snapshot().slides.length;
      const slides = presentationInfo(reopen(e)).slides;
      // One entry per slide part, ids unique, and the same count in the rail and the file.
      expect(slides.length).toBe(shown);
      expect(new Set(slides.map((s) => s.sldId)).size).toBe(slides.length);
      expect(new Set(slides.map((s) => s.partName)).size).toBe(slides.length);
    }
  });

  // Without actor-scoped ids, A's new slide and B's first new slide would get one part name
  // and r:id, and one slide would be lost (7, not 6 + 1 - 1 + 2 = 8). Each replica mints in
  // its own stripe (`replica-ids.ts` in core).
  test('New Slide on one replica, Delete and two New Slides on the other: every slide saves', async () => {
    for (const e of await addDeleteAdd()) {
      expect(e.snapshot().slides.length).toBe(8);
      expect(presentationInfo(reopen(e)).slides.length).toBe(8);
    }
  });

  /** A deck with no slides and no `p:sldIdLst` (the element is optional then). */
  const empty = (): Uint8Array => {
    const f = unzipSync(buildPptx({ slides: [] }));
    const xml = strFromU8(f['ppt/presentation.xml']!).replace(
      /<p:sldIdLst\s*(\/>|><\/p:sldIdLst>)/,
      ''
    );
    expect(xml).not.toContain('sldIdLst');
    f['ppt/presentation.xml'] = strToU8(xml);
    return zipSync(f);
  };
  const slideLists = (pkg: PptxPackage): number =>
    elements(xmlPartOf(pkg, presentationInfo(pkg).presentationPart)!.root).filter(
      (el) => el.local === 'sldIdLst'
    ).length;

  test('New Slide on both replicas of a deck without a slide list saves one list', async () => {
    const { sync, a, b } = await pair(empty());
    sync.pause();
    expect(a.exec({ type: 'addSlide' }).ok).toBe(true);
    expect(b.exec({ type: 'addSlide' }).ok).toBe(true);
    sync.flush();
    expectConverged(a, b);
    // The merged part holds the list of each replica.
    expect(slideLists(a.session!.store.package)).toBe(2);
    for (const e of [a, b]) {
      const pkg = reopen(e);
      // `CT_Presentation` allows one `p:sldIdLst`; the rail shows what the file holds.
      expect(slideLists(pkg)).toBe(1);
      const slides = presentationInfo(pkg).slides;
      expect(slides.length).toBe(e.snapshot().slides.length);
      expect(slides.length).toBeGreaterThanOrEqual(1);
      expect(new Set(slides.map((s) => s.sldId)).size).toBe(slides.length);
    }
    // The next command joins the lists in the part itself, on both replicas.
    expect(a.exec({ type: 'addSlide' }).ok).toBe(true);
    expectConverged(a, b);
    expect(slideLists(a.session!.store.package)).toBe(1);
  });

  // Without actor-scoped ids, both replicas would mint one part name and r:id for their first
  // slide, and the two new slides would be one part after the merge.
  test('New Slide on both replicas of a deck without a slide list keeps both slides', async () => {
    const { sync, a, b } = await pair(empty());
    sync.pause();
    a.exec({ type: 'addSlide' });
    b.exec({ type: 'addSlide' });
    sync.flush();
    for (const e of [a, b]) expect(presentationInfo(reopen(e)).slides.length).toBe(2);
  });

  test('removing a section with its slides while the other replica edits one converges', async () => {
    const { sync, a, b } = await pair();
    const details = a.snapshot().sections!.find((s) => s.name === 'Details')!.id;
    sync.pause();
    a.exec({ type: 'removeSection', sectionId: details, withSlides: true });
    b.exec({ type: 'setSlidesHidden', indices: [3], hidden: false });
    sync.flush();
    expectConverged(a, b);
    expect(sectionsOf(b)).toEqual(['Opening:0+2']);
  });
});
