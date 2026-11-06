// A long deck that opens progressively, at its edges: a start slide that is not a whole number,
// and a slide or a notes slide that does not parse. A damaged part refuses the deck at once,
// wherever it is read. With no edits the deck closes; after edits it stays open, read-only,
// with an `openFailure`, and no write, undo, save, or export gets past it.

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
    const bytes = withPart(BYTES, slidePart(BYTES, SLIDES - 1), BROKEN);
    const editor = editorFor();
    let failures = 0;
    onEditorEvent(editor, 'openFailure', () => (failures += 1));
    expect((await editor.load(bytes)).ok).toBe(true);
    typeHere(editor, 'kept');
    const revision = editor.snapshot().revision;
    expect(revision).toBeGreaterThan(0);
    const session = editor.session!;
    await refused(editor);
    // The session and its edits stay on screen, in view mode.
    expect(editor.session).toBe(session);
    expect(editor.snapshot().status).toBe('ready');
    expect(editor.snapshot().mode).toBe('view');
    expect(failures).toBe(1);
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
    expect(failures).toBe(1);
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
    await refused(editor);
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
