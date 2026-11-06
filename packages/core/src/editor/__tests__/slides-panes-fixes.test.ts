// Core rules behind the slide pane fixes: the open-failure reasons, `setNotes` with the shown
// formats, the focus after a slide opens from the sorter, and the focus after a show ends.

import { afterAll, beforeAll, describe, expect, test } from 'bun:test';
import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { buildPptx, para, run, shape } from '../../../../../scripts/fixtures/builder.ts';
import { createPptxEditor } from '../pptx-editor.ts';
import { documentOpenFailure, shownOpenFailure } from '../loading.ts';
import {
  focusCurrentSlideOrCanvas,
  openSlideInNormalView,
  refocusAfterSorterHides,
} from '../pane-support.ts';
import { enterSlideshowScreen } from '../slideshow-screen.ts';

const IN = 914_400;
const deck = (notes?: string) =>
  buildPptx({
    slides: [0, 1, 2].map((i) => ({
      layout: 'blank' as const,
      shapes: shape({ id: 2, x: IN, y: IN, cx: 2 * IN, cy: IN, text: para(run(`S${i}`)) }),
      ...(notes && i === 0 ? { notes } : {}),
    })),
  });
const frame = () => new Promise((r) => setTimeout(r, 30));

describe('open failure', () => {
  test('each engine code has a reason group, and each refusal is a new object', () => {
    expect(documentOpenFailure('zip').reason).toBe('damaged');
    expect(documentOpenFailure('xml').reason).toBe('damaged');
    expect(documentOpenFailure('empty').reason).toBe('not-pptx');
    expect(documentOpenFailure('not-a-presentation').reason).toBe('not-pptx');
    expect(documentOpenFailure('too-many-parts').reason).toBe('too-large');
    expect(documentOpenFailure('constructor').reason).toBe('unknown');
    expect(documentOpenFailure('zip')).not.toBe(documentOpenFailure('zip'));
  });

  test('a refused file keeps the open deck, and the loading state hides the failure', async () => {
    const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    await editor.load(deck());
    const r = await editor.load(new Uint8Array([1, 2, 3]));
    expect(r.ok).toBe(false);
    const snap = editor.snapshot();
    expect(snap.status).toBe('ready');
    expect(snap.slides.length).toBe(3);
    expect(shownOpenFailure(snap)?.code).toBe('zip');
    expect(shownOpenFailure({ ...snap, status: 'loading' })).toBe(null);
  });
});

describe('setNotes with the shown formats', () => {
  test('each character takes the format it was shown with, in one undo step', async () => {
    const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    await editor.load(deck('Hi'));
    const off = { italic: false, underline: false, strike: false };
    const r = editor.exec({
      type: 'setNotes',
      index: 0,
      text: 'Hi bold\nmore',
      paragraphs: [
        {
          runs: [
            { text: 'Hi', bold: false, ...off },
            { text: ' bold', bold: true, ...off },
          ],
        },
        { runs: [{ text: 'more', bold: true, ...off }] },
      ],
    });
    expect(r.ok).toBe(true);
    expect(
      editor.snapshot().notesFormatted!.map((p) => p.runs.map((x) => [x.text, x.bold]))
    ).toEqual([
      [
        ['Hi', false],
        [' bold', true],
      ],
      [['more', true]],
    ]);
    editor.exec({ type: 'undo' });
    expect(editor.snapshot().notes).toBe('Hi');
  });

  test('paragraphs with another text, or of the wrong shape, write the text alone', async () => {
    const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    await editor.load(deck('Hi'));
    const bad = [{ runs: [{ text: 'other', bold: true }] }] as never;
    expect(editor.exec({ type: 'setNotes', index: 0, text: 'Hi!', paragraphs: bad }).ok).toBe(true);
    expect(editor.snapshot().notesFormatted![0]!.runs.every((x) => !x.bold)).toBe(true);
  });
});

describe('focus', () => {
  // The engine run has no DOM (scripts/test/preload-dom.ts): these tests register one, and
  // remove it after, so the other engine tests stay DOM-free.
  let registered = false;
  beforeAll(() => {
    if (GlobalRegistrator.isRegistered) return;
    GlobalRegistrator.register();
    registered = true;
  });
  afterAll(async () => {
    if (registered) await GlobalRegistrator.unregister();
  });

  test('a slide opened from the sorter gives focus to the current slide of the rail', async () => {
    const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    await editor.load(deck());
    const scope = document.createElement('div');
    scope.className = 'pptx-editor';
    const list = document.createElement('ol');
    scope.append(list);
    document.body.append(scope);
    editor.exec({ type: 'setPresentationView', view: 'slideSorter' });
    openSlideInNormalView(editor, 2, list);
    // The adapter renders the rail in place of the sorter.
    list.remove();
    const rail = document.createElement('ol');
    rail.className = 'pptx-rail__list';
    for (const i of [0, 1, 2]) {
      const li = document.createElement('li');
      li.tabIndex = -1;
      li.className = 'pptx-rail__item';
      li.setAttribute('aria-current', String(i === 2));
      rail.append(li);
    }
    scope.append(rail);
    await frame();
    expect(editor.snapshot().presentationView).toBe('normal');
    expect(document.activeElement).toBe(rail.children[2]!);
    scope.remove();
  });

  test('a slide opened from the sorter with the rail hidden or absent gives focus to the slide', async () => {
    const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    await editor.load(deck());
    const scope = document.createElement('div');
    scope.className = 'pptx-editor';
    const list = document.createElement('ol');
    scope.append(list);
    document.body.append(scope);
    // A narrow window hides the rail: its current item cannot take focus.
    const rail = document.createElement('ol');
    rail.className = 'pptx-rail__list';
    const item = document.createElement('li');
    item.tabIndex = -1;
    item.setAttribute('aria-current', 'true');
    item.focus = () => {};
    rail.append(item);
    const surface = document.createElement('div');
    surface.className = 'pptx-surface';
    surface.tabIndex = 0;
    editor.exec({ type: 'setPresentationView', view: 'slideSorter' });
    openSlideInNormalView(editor, 1, list);
    list.remove();
    scope.append(rail, surface);
    await frame();
    expect(document.activeElement === surface).toBe(true);
    // No rail at all: the slide again, never the page body.
    surface.blur();
    rail.remove();
    focusCurrentSlideOrCanvas(scope, editor);
    expect(document.activeElement === surface).toBe(true);
    scope.remove();
  });

  test('a sorter hidden while its slide has focus gives focus to the current slide of the rail', async () => {
    const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    await editor.load(deck());
    const scope = document.createElement('div');
    scope.className = 'pptx-editor';
    const list = document.createElement('ol');
    const tile = document.createElement('li');
    tile.tabIndex = 0;
    list.append(tile);
    scope.append(list);
    document.body.append(scope);
    editor.goToSlide(1);
    tile.focus();
    // The Normal view button: the adapter calls the helper, then renders the rail in place.
    editor.exec({ type: 'setPresentationView', view: 'normal' });
    refocusAfterSorterHides(list, editor);
    list.remove();
    const rail = document.createElement('ol');
    rail.className = 'pptx-rail__list';
    for (const i of [0, 1, 2]) {
      const li = document.createElement('li');
      li.tabIndex = -1;
      li.setAttribute('aria-current', String(i === 1));
      rail.append(li);
    }
    scope.append(rail);
    await frame();
    expect(document.activeElement).toBe(rail.children[1]!);
    scope.remove();
  });

  test('a sorter hidden while focus is elsewhere leaves focus where it is', async () => {
    const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    await editor.load(deck());
    const scope = document.createElement('div');
    scope.className = 'pptx-editor';
    const list = document.createElement('ol');
    const field = document.createElement('input');
    const rail = document.createElement('ol');
    rail.className = 'pptx-rail__list';
    const item = document.createElement('li');
    item.tabIndex = -1;
    item.setAttribute('aria-current', 'true');
    rail.append(item);
    scope.append(list, rail);
    document.body.append(scope, field);
    field.focus();
    refocusAfterSorterHides(list, editor);
    list.remove();
    // Focus left on nothing, but it was never in the sorter: the helper does not take it.
    field.blur();
    await frame();
    expect(document.activeElement === item).toBe(false);
    scope.remove();
    field.remove();
  });

  test('after a show, focus goes to the current slide item, not the one the show began on', async () => {
    const scope = document.createElement('div');
    scope.className = 'pptx-editor';
    const rail = document.createElement('ol');
    rail.className = 'pptx-rail__list';
    const items = [0, 1, 2].map((i) => {
      const li = document.createElement('li');
      li.tabIndex = -1;
      li.className = 'pptx-rail__item';
      li.setAttribute('aria-current', String(i === 0));
      rail.append(li);
      return li;
    });
    const show = document.createElement('div');
    show.tabIndex = -1;
    scope.append(rail, show);
    document.body.append(scope);
    items[0]!.focus();
    const leave = enterSlideshowScreen(show);
    items[0]!.setAttribute('aria-current', 'false');
    items[2]!.setAttribute('aria-current', 'true');
    leave();
    show.remove();
    await frame();
    expect(document.activeElement).toBe(items[2]!);
    scope.remove();
  });
});
