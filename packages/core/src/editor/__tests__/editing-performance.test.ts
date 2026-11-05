// Regression tests for editing performance and for the
// SVG id prefixes of two editors on one page.

import { afterAll, describe, expect, test } from 'bun:test';
import { GlobalWindow } from 'happy-dom';
import { buildPptx, para, run, shape } from '../../../../../scripts/fixtures/builder.ts';
import { textDeck } from '../../../../../scripts/bench/text-deck.ts';
import { find, presentationInfo, xmlPartOf } from '../../store/index.ts';
import { createPptxEditor, type PptxEditorInstance } from '../index.ts';
import { renderKey, slideDependencies } from '../slide-dependencies.ts';

// A DOM for this file only: the engine test run registers no global DOM.
const win = new GlobalWindow();
const doc = win.document as unknown as Document;
afterAll(() => win.happyDOM.close());

const IN = 914_400;

async function open(bytes: Uint8Array): Promise<PptxEditorInstance> {
  const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  const r = await editor.load(bytes);
  expect(r.ok).toBe(true);
  return editor;
}

/** Put the caret at the end of the first paragraph of the shape named `name` on `slide`. */
function caretIn(editor: PptxEditorInstance, slide: number, name: string): void {
  editor.goToSlide(slide);
  const lay = editor.session!.layout(slide)!;
  const item = lay.items.find((i) => i.kind === 'shape' && i.element.name === name)!;
  const end = item.kind === 'shape' ? (item.text?.paragraphLengths[0] ?? 0) : 0;
  const caret = { paragraph: 0, offset: end };
  editor.setSelection({ kind: 'text', nodeId: item.element.nodeId, anchor: caret, focus: caret });
}

const wait = (ms: number): Promise<void> => new Promise((r) => setTimeout(r, ms));

describe('P1: layouts and thumbnails are keyed by the parts a slide reads', () => {
  test('an edit on one slide keeps the layouts of the other slides', async () => {
    const editor = await open(textDeck(4));
    const session = editor.session!;
    const before = [0, 1, 2, 3].map((i) => session.layout(i));
    caretIn(editor, 1, 'Body');
    expect(editor.input('text', 'x').ok).toBe(true);
    const after = [0, 1, 2, 3].map((i) => session.layout(i));
    expect(after[0]).toBe(before[0]!);
    expect(after[1]).not.toBe(before[1]!);
    expect(after[2]).toBe(before[2]!);
    expect(after[3]).toBe(before[3]!);
    editor.destroy();
  });

  test('a slide summary keeps its object unless its slide renders differently', async () => {
    const editor = await open(textDeck(3));
    const before = editor.snapshot().slides;
    caretIn(editor, 2, 'Body');
    editor.input('text', 'x');
    const after = editor.snapshot().slides;
    expect(after[0]).toBe(before[0]!);
    expect(after[1]).toBe(before[1]!);
    expect(after[2]).not.toBe(before[2]!);
    editor.destroy();
  });

  test('an edit on the master changes every slide that uses it', async () => {
    const editor = await open(textDeck(3));
    const session = editor.session!;
    const pkg = session.store.package;
    const keys = [0, 1, 2].map((i) => renderKey(pkg, i, session.fontGeneration));
    const layouts = [0, 1, 2].map((i) => session.layout(i));
    const master = presentationInfo(pkg).slides[0]!.masterPart!;
    const off = find(xmlPartOf(pkg, master)!.root, 'a:off')!;
    const r = session.store.transact([
      { op: 'setAttribute', part: master, node: off.id, name: { ns: '', local: 'x' }, value: '7' },
    ]);
    expect(r.ok).toBe(true);
    const next = session.store.package;
    for (const i of [0, 1, 2]) {
      expect(renderKey(next, i, session.fontGeneration)).not.toBe(keys[i]!);
      expect(session.layout(i)).not.toBe(layouts[i]!);
    }
    editor.destroy();
  });

  test('moving a slide changes its dependencies (slide numbers)', async () => {
    const editor = await open(textDeck(3));
    const pkg = editor.session!.store.package;
    const a = slideDependencies(pkg, 0)!;
    const b = slideDependencies(pkg, 1)!;
    expect(a[0]).toBe(0);
    expect(b[0]).toBe(1);
    expect(slideDependencies(pkg, 9)).toBeUndefined();
    editor.destroy();
  });

  test('a thumbnail of an unchanged slide is not repainted; a changed one repaints when idle', async () => {
    const editor = await open(textDeck(3));
    const boxes = [0, 1, 2].map(() => doc.createElement('div') as unknown as HTMLElement);
    boxes.forEach((box, i) => editor.renderThumbnail(i, box));
    const first = boxes.map((b) => b.firstElementChild);
    expect(first.every((svg) => svg?.localName === 'svg')).toBe(true);

    caretIn(editor, 1, 'Body');
    editor.input('text', 'x');
    boxes.forEach((box, i) => editor.renderThumbnail(i, box));
    // Nothing repaints during the edit: slide 1 waits for an idle time.
    expect(boxes.map((b) => b.firstElementChild)).toEqual(first);

    await wait(400);
    expect(boxes[0]!.firstElementChild).toBe(first[0]!);
    expect(boxes[1]!.firstElementChild).not.toBe(first[1]!);
    expect(boxes[2]!.firstElementChild).toBe(first[2]!);
    editor.destroy();
  });

  test('a container that shows another slide paints at once', async () => {
    const editor = await open(textDeck(3));
    const box = doc.createElement('div') as unknown as HTMLElement;
    editor.renderThumbnail(0, box);
    const first = box.firstElementChild;
    editor.renderThumbnail(2, box);
    expect(box.firstElementChild).not.toBe(first);
    expect(box.firstElementChild?.localName).toBe('svg');
    editor.destroy();
  });

  test('a change made on the store directly reaches the chrome', async () => {
    const editor = await open(textDeck(2));
    const session = editor.session!;
    let calls = 0;
    editor.subscribe(() => (calls += 1));
    const pkg = session.store.package;
    const part = presentationInfo(pkg).slides[0]!.partName;
    const off = find(xmlPartOf(pkg, part)!.root, 'a:off')!;
    session.store.transact([
      { op: 'setAttribute', part, node: off.id, name: { ns: '', local: 'x' }, value: '9' },
    ]);
    expect(calls).toBe(1);
    expect(editor.snapshot().revision).toBe(session.store.revision);
    editor.destroy();
  });
});

describe('P2: "can this command run" does not re-fit autofit text', () => {
  test('a probe in an overflowing normAutofit box measures no text', async () => {
    const editor = await open(textDeck(3));
    // Slide 1 is the middle slide: its `Fit` box overflows.
    caretIn(editor, 1, 'Fit');
    editor.snapshot();
    const measurer = editor.session!.measurer as { advances: (...a: unknown[]) => unknown };
    const advances = measurer.advances.bind(measurer);
    let calls = 0;
    measurer.advances = (...a: unknown[]) => {
      calls += 1;
      return advances(...a);
    };
    expect(editor.can({ type: 'toggleBold' })).toEqual({ ok: true });
    expect(editor.can({ type: 'setFontSize', points: 40 })).toEqual({ ok: true });
    expect(calls).toBe(0);
    // The command itself still re-fits.
    expect(editor.exec({ type: 'setFontSize', points: 40 }).ok).toBe(true);
    expect(calls).toBeGreaterThan(0);
    editor.destroy();
  });

  test('answers stay cached while only the zoom changes', async () => {
    const editor = await open(textDeck(3));
    caretIn(editor, 0, 'Body');
    const measurer = editor.session!.measurer as { resolve: (...a: unknown[]) => unknown };
    expect(editor.can({ type: 'toggleItalic' }).ok).toBe(true);
    const resolve = measurer.resolve.bind(measurer);
    let calls = 0;
    measurer.resolve = (...a: unknown[]) => {
      calls += 1;
      return resolve(...a);
    };
    editor.setZoom(0.5);
    expect(editor.can({ type: 'toggleItalic' }).ok).toBe(true);
    expect(calls).toBe(0);
    editor.destroy();
  });
});

describe('P4: snapshot caches', () => {
  test('font substitutions are not recomputed when the package did not change', async () => {
    const editor = await open(textDeck(2));
    const session = editor.session!;
    const original = session.substitutedFamilies.bind(session);
    let calls = 0;
    session.substitutedFamilies = () => {
      calls += 1;
      return original();
    };
    editor.snapshot();
    const base = calls;
    for (const z of [0.5, 0.75, 1]) {
      editor.setZoom(z);
      editor.snapshot();
    }
    editor.goToSlide(1);
    editor.snapshot();
    expect(calls).toBe(base);
    caretIn(editor, 0, 'Body');
    editor.input('text', 'x');
    editor.snapshot();
    expect(calls).toBe(base + 1);
    editor.destroy();
  });

  test('the notes text follows notes edits and slide switches', async () => {
    const bytes = buildPptx({
      slides: [
        {
          shapes: shape({ id: 2, x: IN, y: IN, cx: 4 * IN, cy: IN, text: para(run('A')) }),
          notes: 'First note',
        },
        { shapes: '', layout: 'blank', notes: 'Second note' },
      ],
    });
    const editor = await open(bytes);
    expect(editor.snapshot().notes).toBe('First note');
    editor.setZoom(0.5);
    expect(editor.snapshot().notes).toBe('First note');
    editor.goToSlide(1);
    expect(editor.snapshot().notes).toBe('Second note');
    expect(editor.exec({ type: 'setNotes', index: 1, text: 'Changed' }).ok).toBe(true);
    expect(editor.snapshot().notes).toBe('Changed');
    editor.goToSlide(0);
    expect(editor.snapshot().notes).toBe('First note');
    editor.destroy();
  });
});

describe('SVG ids', () => {
  const gradient =
    '<a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs>' +
    '<a:gs pos="100000"><a:srgbClr val="0000FF"/></a:gs></a:gsLst><a:lin ang="0"/></a:gradFill>';
  const deck = (): Uint8Array =>
    buildPptx({
      slides: [0, 1].map(() => ({
        shapes: shape({ id: 2, x: IN, y: IN, cx: 4 * IN, cy: IN, fill: gradient }),
        layout: 'blank',
      })),
    });

  const ids = (root: Element): string[] =>
    [...root.querySelectorAll('[id]')].map((e) => e.getAttribute('id')!);

  test('two editors on one page never share an id', async () => {
    // The surface watches its size; this file has no global DOM, so lend it happy-dom's.
    const g = globalThis as { ResizeObserver?: unknown };
    const had = 'ResizeObserver' in g;
    g.ResizeObserver ??= (win as unknown as { ResizeObserver: unknown }).ResizeObserver;
    const page = doc.createElement('div') as unknown as HTMLElement;
    const editors = [await open(deck()), await open(deck())];
    for (const editor of editors) {
      const surface = doc.createElement('div') as unknown as HTMLElement;
      page.append(surface);
      editor.attach(surface);
      for (const i of [0, 1]) {
        const box = doc.createElement('div') as unknown as HTMLElement;
        page.append(box);
        editor.renderThumbnail(i, box);
      }
    }
    const all = ids(page);
    expect(all.length).toBeGreaterThan(4);
    expect(new Set(all).size).toBe(all.length);
    for (const editor of editors) editor.destroy();
    if (!had) delete g.ResizeObserver;
  });

  test('a repainted thumbnail takes new ids', async () => {
    const editor = await open(deck());
    const a = doc.createElement('div') as unknown as HTMLElement;
    const b = doc.createElement('div') as unknown as HTMLElement;
    editor.renderThumbnail(0, a);
    editor.renderThumbnail(0, b);
    const both = [...ids(a), ...ids(b)];
    expect(both.length).toBeGreaterThan(1);
    expect(new Set(both).size).toBe(both.length);
    editor.destroy();
  });
});
