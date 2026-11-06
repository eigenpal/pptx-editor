// A long deck opens progressively: the start slide parses, lays out, and takes input first, and
// the other slides parse in background tasks. Everything the editor does before they finish
// gives the same result as on a deck that opened whole: navigation, edits, undo, find, and
// save. A slide that does not parse refuses the deck once it is read.

import { afterAll, describe, expect, test } from 'bun:test';
import { GlobalWindow } from 'happy-dom';
import { readZip, writeZip } from '@eigenpal/ooxml-core';
import { textDeck } from '../../../../../scripts/bench/text-deck.ts';
import {
  canonicalPackageFingerprint,
  deferredParseCount,
  isDeferredPending,
  presentationInfo,
  readPptx,
  PPTX_ZIP_LIMITS,
} from '../../store/index.ts';
import { createPptxEditor, onEditorEvent, type PptxEditorInstance } from '../index.ts';
import { openPresentation } from '../document-session.ts';
import {
  OPENING_SLICE_MS,
  openingOrder,
  ProgressiveOpen,
  PROGRESSIVE_OPEN_CONTENT_BYTES,
} from '../progressive-open.ts';

// A DOM for this file only: the engine test run registers no global DOM.
const win = new GlobalWindow();
const doc = win.document as unknown as Document;
afterAll(() => win.happyDOM.close());

const SLIDES = 300;
const BYTES = textDeck(SLIDES);
const wait = (ms: number): Promise<void> => new Promise((r) => setTimeout(r, ms));

async function open(bytes: Uint8Array = BYTES, startSlide?: number): Promise<PptxEditorInstance> {
  const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  const r = await editor.load(bytes, startSlide === undefined ? undefined : { startSlide });
  expect(r.ok).toBe(true);
  return editor;
}

async function settled(editor: PptxEditorInstance): Promise<void> {
  for (let i = 0; i < 2000 && (editor.snapshot().pendingSlides ?? 0) > 0; i += 1) await wait(2);
  expect(editor.snapshot().pendingSlides ?? 0).toBe(0);
}

const pendingParts = (editor: PptxEditorInstance): number =>
  editor.session!.info.slides.filter((s) =>
    isDeferredPending(editor.session!.store.package.parts.get(s.partName))
  ).length;

/** The body box of slide `index`: its node id and the length of its first paragraph. */
function body(editor: PptxEditorInstance, index: number) {
  const item = editor.session!.layout(index)!.items.find((i) => i.element.name === 'Body')!;
  const end = item.kind === 'shape' ? (item.text?.paragraphLengths[0] ?? 0) : 0;
  return { id: item.element.nodeId, end };
}

function typeAt(editor: PptxEditorInstance, index: number, text: string): void {
  editor.goToSlide(index);
  const b = body(editor, index);
  const at = { paragraph: 0, offset: b.end };
  editor.setSelection({ kind: 'text', nodeId: b.id, anchor: at, focus: at });
  for (const ch of text) editor.input('text', ch);
}

describe('a progressive open', () => {
  test('the test deck is long enough to open progressively', () => {
    const r = readPptx(BYTES);
    if (!r.ok) throw new Error('read');
    expect(presentationInfo(r.pkg).slides.length).toBe(SLIDES);
    expect(BYTES.length).toBeGreaterThan(0);
    expect(PROGRESSIVE_OPEN_CONTENT_BYTES).toBe(512 * 1024);
  });

  test('parses only the start slide before it resolves, then the rest in the background', async () => {
    const before = deferredParseCount();
    const editor = await open();
    expect(deferredParseCount() - before).toBe(1);
    const snap = editor.snapshot();
    expect(snap.status).toBe('ready');
    expect(snap.currentSlide).toBe(0);
    expect(snap.pendingSlides).toBe(SLIDES - 1);
    expect(snap.slides.length).toBe(SLIDES);
    expect(snap.slides[0]!.pending).toBeUndefined();
    expect(snap.slides.slice(1).every((s) => s.pending === true)).toBe(true);
    await settled(editor);
    expect(pendingParts(editor)).toBe(0);
    expect(editor.snapshot().slides.some((s) => s.pending)).toBe(false);
    expect(editor.session!.opening).toBeNull();
    editor.destroy();
  });

  test('a small deck opens whole, as before', async () => {
    const before = deferredParseCount();
    const editor = await open(textDeck(20));
    expect(editor.snapshot().pendingSlides).toBe(0);
    expect(deferredParseCount() - before).toBe(20);
    expect(editor.session!.opening).toBeNull();
    editor.destroy();
  });

  test('a requested start slide shows and parses first, and its neighbors next', async () => {
    const before = deferredParseCount();
    const editor = await open(BYTES, 250);
    expect(deferredParseCount() - before).toBe(1);
    expect(editor.snapshot().currentSlide).toBe(250);
    expect(editor.snapshot().slides[250]!.pending).toBeUndefined();
    expect(editor.session!.slidePending(250)).toBe(false);
    expect(editor.session!.slidePending(251)).toBe(true);
    const order = openingOrder(editor.session!.info, 250);
    expect(order.slice(0, 2)).toEqual([
      editor.session!.info.slides[251]!.partName,
      editor.session!.info.slides[252]!.partName,
    ]);
    expect(order.at(-1)).toBe(editor.session!.info.slides[249]!.partName);
    expect(order.length).toBe(SLIDES - 1);
    await settled(editor);
    editor.destroy();
    // Out of range: clamped to the deck.
    const last = await open(BYTES, 10_000);
    expect(last.snapshot().currentSlide).toBe(SLIDES - 1);
    last.destroy();
  });

  test('each background slice keeps to its budget and ends early for input', () => {
    const r = readPptx(BYTES, PPTX_ZIP_LIMITS, { defer: () => true });
    if (!r.ok) throw new Error('read');
    const info = presentationInfo(r.pkg);
    const run = new ProgressiveOpen(r.pkg, openingOrder(info, 0));
    // A clock that moves 1 ms for each reading: a slice reads it once per part.
    let now = 0;
    const clock = () => (now += 1);
    const count = () => run.pendingSlides(r.pkg, info);
    const all = count();
    expect(all).toBe(SLIDES);
    run.step(OPENING_SLICE_MS, clock, () => false);
    const first = all - count();
    expect(first).toBeGreaterThan(1);
    expect(first).toBeLessThanOrEqual(OPENING_SLICE_MS);
    // Input waits: one part, then the slice ends.
    run.step(OPENING_SLICE_MS, clock, () => true);
    expect(all - count()).toBe(first + 1);
    // A part a thumbnail asks for parses next.
    const wanted = info.slides[200]!.partName;
    run.want(wanted);
    const step = run.step(OPENING_SLICE_MS, clock, () => true);
    expect(step.wanted).toBe(true);
    expect(isDeferredPending(r.pkg.parts.get(wanted))).toBe(false);
    expect(isDeferredPending(r.pkg.parts.get(info.slides[first + 2]!.partName))).toBe(true);
    expect(run.complete()).toBeNull();
    expect(run.settled).toBe(true);
    expect(count()).toBe(0);
  });

  test('navigation to a slide still opening lays it out at once, as a whole open does', async () => {
    const editor = await open();
    const whole = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    await whole.load(textDeck(20));
    expect(editor.session!.slidePending(150)).toBe(true);
    editor.goToSlide(150);
    expect(editor.snapshot().currentSlide).toBe(150);
    const lay = editor.session!.layout(150)!;
    expect(lay.items.map((i) => i.element.name)).toEqual(['Title', 'Body', 'Fit']);
    expect(editor.session!.slidePending(150)).toBe(false);
    editor.destroy();
    whole.destroy();
  });

  test('edits, undo, find, and save before the open ends match a whole open', async () => {
    const progressive = await open();
    expect(progressive.snapshot().pendingSlides).toBeGreaterThan(0);
    typeAt(progressive, 0, 'abc');
    typeAt(progressive, 120, 'xyz');
    progressive.exec({ type: 'undo' });
    typeAt(progressive, 121, 'qq');
    const found = progressive.find.matches('Slide 280 title');
    expect(found.length).toBe(1);
    const saved = progressive.save()!;
    // Save parsed the rest, so the open is over.
    expect(progressive.snapshot().pendingSlides ?? 0).toBe(0);
    expect(pendingParts(progressive)).toBe(0);

    // The same session on the deck parsed whole before the first edit.
    const whole = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    expect((await whole.load(BYTES)).ok).toBe(true);
    expect(whole.session!.opening!.complete()).toBeNull();
    typeAt(whole, 0, 'abc');
    typeAt(whole, 120, 'xyz');
    whole.exec({ type: 'undo' });
    typeAt(whole, 121, 'qq');
    expect(saved).toEqual(whole.save()!);
    progressive.destroy();
    whole.destroy();
  });

  test('a saved deck that was never edited is the deck an eager open saves', async () => {
    const progressive = await open();
    const eager = readPptx(BYTES);
    if (!eager.ok) throw new Error('eager');
    const saved = progressive.save()!;
    const reopened = readPptx(saved);
    if (!reopened.ok) throw new Error('reopen');
    expect(canonicalPackageFingerprint(reopened.pkg)).toBe(canonicalPackageFingerprint(eager.pkg));
    progressive.destroy();
  });

  test('an undo brings back a deleted slide that never parsed, and the open still settles', async () => {
    const editor = await open();
    editor.goToSlide(40);
    expect(editor.exec({ type: 'deleteSlide', index: 40 }).ok).toBe(true);
    expect(editor.snapshot().slides.length).toBe(SLIDES - 1);
    editor.exec({ type: 'undo' });
    expect(editor.snapshot().slides.length).toBe(SLIDES);
    await settled(editor);
    expect(pendingParts(editor)).toBe(0);
    editor.destroy();
  });

  test('the rail shows a placeholder until the slide is ready, then paints it', async () => {
    const editor = await open();
    const box = doc.createElement('div') as unknown as HTMLElement;
    doc.body.append(box as unknown as Node);
    const pendingIndex = 299;
    expect(editor.snapshot().slides[pendingIndex]!.pending).toBe(true);
    editor.renderThumbnail(pendingIndex, box, { defer: true });
    expect(box.firstElementChild).toBeNull();
    // A thumbnail that paints at once (no `defer`) opens its slide now.
    const now = doc.createElement('div') as unknown as HTMLElement;
    editor.renderThumbnail(298, now);
    expect(now.firstElementChild?.localName).toBe('svg');
    expect(editor.session!.slidePending(298)).toBe(false);
    await settled(editor);
    expect(editor.snapshot().slides[pendingIndex]!.pending).toBeUndefined();
    editor.renderThumbnail(pendingIndex, box, { defer: true });
    for (let i = 0; i < 200 && !box.firstElementChild; i += 1) await wait(5);
    expect(box.firstElementChild?.localName).toBe('svg');
    editor.destroy();
    box.remove();
  });

  test('progress publishes as selection changes, and a later load stops the old open', async () => {
    const editor = await open();
    let heard = 0;
    const off = onEditorEvent(editor, 'selectionChange', () => (heard += 1));
    await settled(editor);
    expect(heard).toBeGreaterThan(0);
    off();
    // A second load while the first still opens: the first session's tasks stop.
    const first = await open();
    const session = first.session!;
    expect((await first.load(textDeck(20))).ok).toBe(true);
    await wait(50);
    expect(first.session).not.toBe(session);
    expect(session.opening).not.toBeNull();
    first.destroy();
    editor.destroy();
  });

  test('a family that only a slide parsed later names loads its substitute when the open settles', async () => {
    const r = readPptx(BYTES);
    if (!r.ok) throw new Error('read');
    const late = presentationInfo(r.pkg).slides[280]!.partName;
    const zip = readZip(BYTES);
    if (!zip.ok) throw new Error('zip');
    const entries = new Map(zip.entries);
    const xml = new TextDecoder().decode(entries.get(late)!);
    const named = xml.replace(
      /<a:rPr lang="en-US" dirty="0" sz="1800">/,
      '<a:rPr lang="en-US" dirty="0" sz="1800"><a:latin typeface="Courier New"/>'
    );
    expect(named).not.toBe(xml);
    entries.set(late, new TextEncoder().encode(named));
    const opened = await openPresentation(writeZip(entries), { progressive: {} });
    if (!opened.ok) throw new Error('open');
    const session = opened.session;
    const faces = session.fontReport.faces.length;
    const generation = session.fontGeneration;
    expect(session.substitutedFamilies().some((f) => f.family === 'Courier New')).toBe(false);
    expect(session.opening!.complete()).toBeNull();
    expect(await session.addMissingSubstitutes()).toBe(true);
    expect(session.fontReport.faces.length).toBeGreaterThan(faces);
    expect(session.fontGeneration).toBe(generation + 1);
    const courier = session.substitutedFamilies().find((f) => f.family === 'Courier New');
    expect(courier?.metricCompatible).toBe(true);
    expect(await session.addMissingSubstitutes()).toBe(false);
    session.dispose();
  });

  test('a slide that does not parse refuses the deck once it is read', async () => {
    const r = readPptx(BYTES);
    if (!r.ok) throw new Error('read');
    const broken = presentationInfo(r.pkg).slides[200]!.partName;
    const zip = readZip(BYTES);
    if (!zip.ok) throw new Error('zip');
    const entries = new Map(zip.entries);
    entries.set(broken, new TextEncoder().encode('<p:sld xmlns:p="urn:x"><p:cSld>'));
    const bytes = writeZip(entries);
    const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    let failures = 0;
    onEditorEvent(editor, 'openFailure', () => (failures += 1));
    expect((await editor.load(bytes)).ok).toBe(true);
    expect(editor.snapshot().status).toBe('ready');
    for (let i = 0; i < 2000 && editor.snapshot().status === 'ready'; i += 1) await wait(2);
    expect(editor.snapshot().status).toBe('error');
    expect(editor.snapshot().openFailure).toMatchObject({ code: 'xml', reason: 'damaged' });
    expect(editor.session).toBeNull();
    expect(editor.save()).toBeNull();
    expect(failures).toBe(1);
    editor.destroy();
  });

  test('a save before the background reaches a broken slide refuses the deck', async () => {
    const r = readPptx(BYTES);
    if (!r.ok) throw new Error('read');
    const broken = presentationInfo(r.pkg).slides[299]!.partName;
    const zip = readZip(BYTES);
    if (!zip.ok) throw new Error('zip');
    const entries = new Map(zip.entries);
    entries.set(broken, new TextEncoder().encode('<p:sld xmlns:p="urn:x"><p:cSld>'));
    const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    expect((await editor.load(writeZip(entries))).ok).toBe(true);
    expect(editor.save()).toBeNull();
    expect(editor.snapshot().status).toBe('error');
    editor.destroy();
  });
});
