// A long deck that opens progressively, at its edges: a start slide that is not a whole number,
// and a slide or a notes slide that does not parse. A damaged part refuses the deck at once,
// wherever it is read. With no edits the deck closes; after edits it stays open, read-only
// (`readOnlyReason`), its other slides go on opening, and no write, undo, save, or export gets
// past it.

import { afterAll, describe, expect, test } from 'bun:test';
import { GlobalWindow } from 'happy-dom';
import { readZip, writeZip } from '@eigenpal/ooxml-core';
import { buildPptx, para, run, shape } from '../../../../../scripts/fixtures/builder.ts';
import { textDeck } from '../../../../../scripts/bench/text-deck.ts';
import { presentationInfo, readPptx } from '../../store/index.ts';
import { createBrowserAutomationHost } from '../automation-host.ts';
import { createPptxEditor, onEditorEvent, type PptxEditorInstance } from '../index.ts';

// A DOM for this file only: the engine test run registers no global DOM.
const win = new GlobalWindow();
afterAll(() => win.happyDOM.close());

const SLIDES = 300;
const BYTES = textDeck(SLIDES);
const BROKEN = '<p:sld xmlns:p="urn:x"><p:cSld>';
const wait = (ms: number): Promise<void> => new Promise((r) => setTimeout(r, ms));

const editorFor = () => createPptxEditor({ fixedMeasurer: true, substitutes: false });

/** `bytes` with part `name` replaced by `text`. */
function withPart(bytes: Uint8Array, name: string, text: string): Uint8Array {
  const zip = readZip(bytes);
  if (!zip.ok) throw new Error('zip');
  const entries = new Map(zip.entries);
  entries.set(name, new TextEncoder().encode(text));
  return writeZip(entries);
}

function slidePart(bytes: Uint8Array, index: number): string {
  const r = readPptx(bytes);
  if (!r.ok) throw new Error('read');
  return presentationInfo(r.pkg).slides[index]!.partName;
}

/** Type `text` at the end of the first body paragraph of the current slide. */
function typeHere(editor: PptxEditorInstance, text: string): void {
  const index = editor.snapshot().currentSlide;
  const item = editor.session!.layout(index)!.items.find((i) => i.element.name === 'Body')!;
  const end = item.kind === 'shape' ? (item.text?.paragraphLengths[0] ?? 0) : 0;
  const at = { paragraph: 0, offset: end };
  editor.setSelection({ kind: 'text', nodeId: item.element.nodeId, anchor: at, focus: at });
  for (const ch of text) editor.input('text', ch);
}

async function refused(editor: PptxEditorInstance): Promise<void> {
  for (let i = 0; i < 2000 && !editor.snapshot().openFailure; i += 1) await wait(2);
  expect(editor.snapshot().openFailure).toMatchObject({ code: 'xml', reason: 'damaged' });
}

/** Wait until the deck turns read-only after edits. */
async function readOnly(editor: PptxEditorInstance): Promise<void> {
  for (let i = 0; i < 2000 && !editor.snapshot().readOnlyReason; i += 1) await wait(1);
  expect(editor.snapshot().readOnlyReason).toBe('damaged-slide');
}

describe('the start slide of a progressive open', () => {
  test('a start slide that is not a whole number is normalized, and the deck opens', async () => {
    const cases: [number, number][] = [
      [1.5, 1],
      [Number.NaN, 0],
      [-1, 0],
      [Number.POSITIVE_INFINITY, 0],
      [Number.NEGATIVE_INFINITY, 0],
      [SLIDES + 0.5, SLIDES - 1],
    ];
    for (const [startSlide, shown] of cases) {
      const editor = editorFor();
      expect((await editor.load(BYTES, { startSlide })).ok).toBe(true);
      expect(editor.snapshot().status).toBe('ready');
      expect(editor.snapshot().currentSlide).toBe(shown);
      expect(editor.snapshot().slides[shown]!.pending).toBeUndefined();
      editor.destroy();
    }
  });

  test('a small deck normalizes the start slide the same way', async () => {
    const editor = editorFor();
    expect((await editor.load(textDeck(5), { startSlide: 2.7 })).ok).toBe(true);
    expect(editor.snapshot().currentSlide).toBe(2);
    editor.destroy();
  });
});

describe('a damaged part in a progressive open', () => {
  test('after edits, the deck stays open read-only, and nothing writes past the damage', async () => {
    const bytes = withPart(BYTES, slidePart(BYTES, 3), BROKEN);
    const editor = editorFor();
    let failures = 0;
    onEditorEvent(editor, 'openFailure', () => (failures += 1));
    expect((await editor.load(bytes)).ok).toBe(true);
    typeHere(editor, 'kept');
    const revision = editor.snapshot().revision;
    expect(revision).toBeGreaterThan(0);
    const session = editor.session!;
    // The snapshot that first says read-only: the other slides still wait, with placeholders.
    let first: ReturnType<PptxEditorInstance['snapshot']> | null = null;
    const off = editor.subscribe(() => {
      const snap = editor.snapshot();
      if (!first && snap.readOnlyReason) first = snap;
    });
    await readOnly(editor);
    off();
    expect(first!.pendingSlides ?? 0).toBeGreaterThan(SLIDES / 2);
    expect(first!.slides.filter((s) => s.pending).length).toBe(first!.pendingSlides!);
    // The session and its edits stay on screen, in view mode. The file did open: no failure.
    expect(editor.session).toBe(session);
    expect(editor.snapshot().status).toBe('ready');
    expect(editor.snapshot().mode).toBe('view');
    expect(editor.snapshot().openFailure ?? null).toBeNull();
    expect(failures).toBe(0);
    expect(editor.setMode('edit')).toBeUndefined();
    expect(editor.snapshot().mode).toBe('view');
    // No command, undo, store transaction, or save gets past it.
    expect(editor.exec({ type: 'deleteSlide', index: 0 }).ok).toBe(false);
    expect(editor.exec({ type: 'undo' }).ok).toBe(false);
    expect(session.store.canUndo).toBe(false);
    expect(session.store.transact([{ op: 'removeNode', part: 'x', node: 'y' }]).ok).toBe(false);
    expect(editor.snapshot().revision).toBe(revision);
    expect(editor.save()).toBeNull();
    // The automation host refuses its commits and its export.
    const host = createBrowserAutomationHost(editor);
    const listed = host.execute({ operations: [{ op: 'listSlides' }] });
    expect(listed.ok).toBe(true);
    const handles = listed.results[0]!;
    if (handles.status !== 'ok' || handles.value.kind !== 'handles') throw new Error('list');
    const exported = host.execute({
      operations: [{ op: 'exportSlides', slides: [handles.value.handles[0]!] }],
    });
    expect(exported.ok).toBe(false);
    expect(exported.results[0]).toMatchObject({ error: { code: 'document-unavailable' } });
    expect(failures).toBe(0);
    editor.destroy();
  });

  test('after the damage, no snapshot parses the waiting slides: the background opens them', async () => {
    const bytes = withPart(BYTES, slidePart(BYTES, 2), BROKEN);
    const editor = editorFor();
    expect((await editor.load(bytes)).ok).toBe(true);
    typeHere(editor, 'x');
    const session = editor.session!;
    const waiting = (): number =>
      session.info.slides.filter((_, i) => session.slidePending(i)).length;
    let atDamage = -1;
    let afterReads = -1;
    const off = editor.subscribe(() => {
      if (atDamage >= 0 || !editor.snapshot().readOnlyReason) return;
      atDamage = waiting();
      // New snapshots and a slide change read the summaries and the rail: one slide parses.
      editor.goToSlide(150);
      editor.snapshot();
      editor.goToSlide(0);
      editor.snapshot();
      afterReads = waiting();
    });
    await readOnly(editor);
    off();
    expect(atDamage).toBeGreaterThan(SLIDES / 2);
    expect(atDamage - afterReads).toBeLessThanOrEqual(1);
    for (let i = 0; i < 2000 && waiting() > 0; i += 1) await wait(1);
    expect(waiting()).toBe(0);
    expect(editor.snapshot().slides.some((s) => s.pending)).toBe(false);
    expect(editor.snapshot().readOnlyReason).toBe('damaged-slide');
    editor.destroy();
  });

  test('a reader that finds the damage turns an edited deck read-only once the read returns', async () => {
    const bytes = withPart(BYTES, slidePart(BYTES, 250), BROKEN);
    const editor = editorFor();
    expect((await editor.load(bytes)).ok).toBe(true);
    typeHere(editor, 'x');
    editor.goToSlide(250);
    // The slide on screen lays out: the read finds the damage.
    editor.session!.layout(250);
    expect(editor.session!.damage).not.toBeNull();
    await Promise.resolve();
    expect(editor.snapshot().readOnlyReason).toBe('damaged-slide');
    expect(editor.snapshot().mode).toBe('view');
    expect(editor.snapshot().currentSlide).toBe(250);
    editor.destroy();
  });

  test('the next healthy deck opens in the host mode, and the read-only state fires change once', async () => {
    const editor = editorFor();
    let changes = 0;
    onEditorEvent(editor, 'change', () => (changes += 1));
    expect((await editor.load(withPart(BYTES, slidePart(BYTES, 3), BROKEN))).ok).toBe(true);
    typeHere(editor, 'x');
    const before = changes;
    await readOnly(editor);
    expect(editor.snapshot().mode).toBe('view');
    // One change for the read-only state, and none for the slides that open after it.
    for (let i = 0; i < 2000 && (editor.snapshot().pendingSlides ?? 0) > 0; i += 1) await wait(1);
    expect(changes).toBe(before + 1);
    expect((await editor.load(textDeck(10))).ok).toBe(true);
    expect(editor.snapshot().readOnlyReason ?? null).toBeNull();
    expect(editor.snapshot().mode).toBe('edit');
    expect(editor.exec({ type: 'addSection', index: 1 }).ok).toBe(true);
    editor.destroy();
  });

  test('automation reads of a damaged slide are refused, not answered with an empty slide', async () => {
    const editor = editorFor();
    expect((await editor.load(withPart(BYTES, slidePart(BYTES, SLIDES - 1), BROKEN))).ok).toBe(
      true
    );
    const host = createBrowserAutomationHost(editor);
    const read = host.execute({
      operations: [
        { op: 'getSlide', key: { index: SLIDES - 1 } },
        { op: 'listShapes', container: { result: 0 } },
      ],
    });
    expect(read.ok).toBe(false);
    expect(read.results[1]).toMatchObject({
      status: 'error',
      error: { code: 'document-unavailable', detail: 'damaged' },
    });
    // A healthy slide still reads (the deck closes once the read returns: no edits).
    const fresh = editorFor();
    expect((await fresh.load(withPart(BYTES, slidePart(BYTES, SLIDES - 1), BROKEN))).ok).toBe(true);
    const ok = createBrowserAutomationHost(fresh).execute({
      operations: [
        { op: 'getSlide', key: { index: 0 } },
        { op: 'listShapes', container: { result: 0 } },
      ],
    });
    expect(ok.ok).toBe(true);
    editor.destroy();
    fresh.destroy();
  });

  test('a slide that finishes opening fires no change, so a host that saves on change does not', async () => {
    const editor = editorFor();
    let changes = 0;
    onEditorEvent(editor, 'change', () => (changes += 1));
    expect((await editor.load(BYTES)).ok).toBe(true);
    const after = changes;
    expect(after).toBe(1);
    for (let i = 0; i < 2000 && (editor.snapshot().pendingSlides ?? 0) > 0; i += 1) await wait(2);
    expect(editor.snapshot().pendingSlides ?? 0).toBe(0);
    expect(editor.snapshot().revision).toBe(0);
    expect(changes).toBe(after);
    editor.destroy();
  });

  test('a damaged slide that the user deleted does not refuse the deck', async () => {
    const broken = slidePart(BYTES, SLIDES - 1);
    const bytes = withPart(BYTES, broken, BROKEN);
    const editor = editorFor();
    expect((await editor.load(bytes)).ok).toBe(true);
    editor.goToSlide(0);
    // Delete it without reading it: a selection of the last slide, then Delete.
    expect(editor.exec({ type: 'deleteSlide', index: SLIDES - 1 }).ok).toBe(true);
    for (let i = 0; i < 2000 && (editor.snapshot().pendingSlides ?? 0) > 0; i += 1) await wait(2);
    expect(editor.snapshot().pendingSlides ?? 0).toBe(0);
    expect(editor.snapshot().readOnlyReason ?? null).toBeNull();
    expect(editor.snapshot().openFailure ?? null).toBeNull();
    expect(editor.save()).not.toBeNull();
    // An undo brings it back: the save then reads it, and the deck turns read-only.
    expect(editor.exec({ type: 'undo' }).ok).toBe(true);
    expect(editor.save()).toBeNull();
    await readOnly(editor);
    editor.destroy();
  });

  test('a damaged notes slide refuses the deck once the background reads it', async () => {
    const LONG = 'Quarterly results show steady growth across every region. '.repeat(40);
    const IN = 914_400;
    const base = buildPptx({
      slides: Array.from({ length: SLIDES }, (_, i) => ({
        shapes: shape({
          id: 2,
          name: 'Body',
          x: IN,
          y: IN,
          cx: 8 * IN,
          cy: 4 * IN,
          text: para(run(`${i} ${LONG}`)),
        }),
        ...(i === SLIDES - 2 ? { notes: 'Speaker notes' } : {}),
      })),
    });
    const r = readPptx(base);
    if (!r.ok) throw new Error('read');
    const notes = [...r.pkg.parts.keys()].find((n) => n.includes('/notesSlides/'))!;
    expect(notes).toBeDefined();
    const bytes = withPart(base, notes, '<p:notes xmlns:p="urn:x"><p:cSld>');
    const editor = editorFor();
    expect((await editor.load(bytes)).ok).toBe(true);
    expect(editor.snapshot().pendingSlides).toBeGreaterThan(0);
    await refused(editor);
    // No edits: the deck closes.
    expect(editor.session).toBeNull();
    expect(editor.snapshot().status).toBe('error');
    editor.destroy();
  });

  test('an export through the automation host before the background reaches the damage is refused', async () => {
    const bytes = withPart(BYTES, slidePart(BYTES, SLIDES - 1), BROKEN);
    const editor = editorFor();
    expect((await editor.load(bytes)).ok).toBe(true);
    typeHere(editor, 'x');
    expect(editor.snapshot().openFailure).toBeFalsy();
    const host = createBrowserAutomationHost(editor);
    const listed = host.execute({ operations: [{ op: 'listSlides' }] });
    const handles = listed.results[0]!;
    if (handles.status !== 'ok' || handles.value.kind !== 'handles') throw new Error('list');
    // The export reads the whole deck: it finishes the open, finds the damage, and refuses.
    const exported = host.execute({
      operations: [{ op: 'exportSlides', slides: [handles.value.handles[0]!] }],
    });
    expect(exported.ok).toBe(false);
    expect(exported.results[0]).toMatchObject({
      status: 'error',
      error: { code: 'document-unavailable', detail: 'damaged' },
    });
    // The export finished the open and found the damage: the edited deck is read-only at once.
    expect(editor.snapshot().readOnlyReason).toBe('damaged-slide');
    expect(editor.snapshot().mode).toBe('view');
    expect(editor.save()).toBeNull();
    editor.destroy();
  });

  test('a destroy during the background parse stops it', async () => {
    const editor = editorFor();
    expect((await editor.load(BYTES)).ok).toBe(true);
    const session = editor.session!;
    const open = session.opening!;
    expect(open).not.toBeNull();
    const pending = open.pendingSlides(session.store.package, session.info);
    expect(pending).toBeGreaterThan(0);
    editor.destroy();
    await wait(100);
    // No background task ran on the destroyed session.
    expect(open.pendingSlides(session.store.package, session.info)).toBe(pending);
    expect(editor.snapshot().pendingSlides ?? 0).toBe(0);
  });
});
