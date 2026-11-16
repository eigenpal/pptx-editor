// The document refresh controller, after DOCX Editor's document-refresh.test.ts and
// document-refresh-edit.test.ts: capture, cumulative results, refusals, both modes, change
// locations, the change flash, and a refresh that never blinks: no loading screen, no empty
// canvas, and no repaint of what did not change.

import { afterAll, describe, expect, test } from 'bun:test';
import { GlobalWindow } from 'happy-dom';
import { buildPptx, para, run, shape } from '../../../../../scripts/fixtures/builder.ts';
import {
  createDocumentRefresh,
  createPptxEditor,
  onEditorEvent,
  type ComputedRefreshChange,
  type DocumentChange,
  type EditorModule,
  type PptxEditorInstance,
  type RefreshChangeInput,
} from '../index.ts';

const window = new GlobalWindow();
const doc = window.document as unknown as Document;
afterAll(() => window.happyDOM.close());
const g = globalThis as { ResizeObserver?: unknown };
g.ResizeObserver ??= (window as unknown as { ResizeObserver: unknown }).ResizeObserver;

const box = (id: number, text: string, x = 400_000) =>
  shape({ id, x, y: 400_000, cx: 4_000_000, cy: 800_000, text: para(run(text)) });
const deck = (...slides: string[][]) =>
  buildPptx({
    slides: slides.map((texts) => ({ shapes: texts.map((t, i) => box(i + 2, t)).join('') })),
  });

/** An editor in a page laid out as the adapters lay it out: a rail beside the canvas. */
async function mount(bytes: Uint8Array, modules: EditorModule[] = []) {
  const shell = doc.createElement('div');
  shell.className = 'pptx-editor';
  const rail = doc.createElement('ol');
  rail.className = 'pptx-rail';
  const host = doc.createElement('div');
  shell.append(rail, host);
  doc.body.append(shell);
  const editor = createPptxEditor({ fixedMeasurer: true, modules });
  editor.attach(host);
  expect((await editor.load(bytes)).ok).toBe(true);
  const thumbs = () => [...rail.querySelectorAll<HTMLElement>('.pptx-rail__thumb')];
  const renderRail = () => {
    rail.replaceChildren();
    editor.snapshot().slides.forEach((s) => {
      const li = doc.createElement('li');
      li.setAttribute('data-index', String(s.index));
      const thumb = doc.createElement('div');
      thumb.className = 'pptx-rail__thumb';
      li.append(thumb);
      rail.append(li);
      editor.renderThumbnail(s.index, thumb);
    });
  };
  renderRail();
  const done = () => {
    editor.destroy();
    shell.remove();
  };
  return { editor, host, rail, thumbs, renderRail, done };
}

/** The changes a review module computes: here, the ones the test gives. */
function reviewWith(changes: () => readonly ComputedRefreshChange[]): EditorModule {
  return {
    id: 'review',
    review: { displayModes: ['final'], collectReviewItems: () => [], refreshChanges: changes },
  };
}

const slideText = (editor: PptxEditorInstance, index: number) => {
  editor.goToSlide(index);
  return editor
    .snapshot()
    .shapes?.map((s) => s.name)
    .join(',');
};

describe('document refresh: reload', () => {
  test('keeps the instance, subscriptions, current slide, and zoom; resets undo', async () => {
    const { editor, done } = await mount(deck(['A'], ['B'], ['C']));
    expect(editor.exec({ type: 'duplicateSlide', index: 0 }).ok).toBe(true);
    editor.goToSlide(2);
    editor.setZoom(1.5);
    const refresh = createDocumentRefresh(editor);
    expect(createDocumentRefresh(editor)).toBe(refresh);
    const submission = await refresh.capture();
    let notified = 0;
    editor.subscribe(() => void (notified += 1));
    const result = await refresh.applyUpdate({
      submission,
      sequence: 1,
      bytes: deck(['A'], ['B changed'], ['C'], ['D']),
    });
    expect(result.ok).toBe(true);
    if (result.ok) expect(result.mode).toBe('reload');
    expect(notified).toBeGreaterThan(0);
    expect(editor.snapshot().slides.length).toBe(4);
    expect(editor.snapshot().currentSlide).toBe(2);
    expect(editor.snapshot().zoom).toBe(1.5);
    expect(editor.snapshot().canUndo).toBe(false);
    refresh.finish(submission);
    done();
  });

  test('never shows the loading screen or an empty canvas, and repaints only what changed', async () => {
    const { editor, host, thumbs, done } = await mount(deck(['A'], ['B'], ['C']));
    const before = thumbs().map((t) => t.firstElementChild);
    const statuses = new Set<string>();
    let emptyCanvas = false;
    editor.subscribe(() => {
      statuses.add(editor.snapshot().status);
      if (!host.querySelector('.pptx-surface__slide')) emptyCanvas = true;
    });
    const refresh = createDocumentRefresh(editor);
    const submission = await refresh.capture();
    const result = await refresh.applyUpdate({
      submission,
      sequence: 1,
      bytes: deck(['A'], ['B changed'], ['C']),
    });
    expect(result.ok).toBe(true);
    expect([...statuses]).toEqual(['ready']);
    expect(emptyCanvas).toBe(false);
    // The rail asks again for each slide; only the changed slide paints a new picture.
    editor.snapshot().slides.forEach((s, i) => editor.renderThumbnail(s.index, thumbs()[i]!));
    const after = thumbs().map((t) => t.firstElementChild);
    expect(after[0]).toBe(before[0]);
    expect(after[2]).toBe(before[2]);
    expect(editor.session!.store.package.parts.get('/ppt/slides/slide1.xml')).toBeDefined();
    done();
  });

  test('the change event of a reload reports source refresh', async () => {
    const { editor, done } = await mount(deck(['A']));
    const seen: DocumentChange[] = [];
    onEditorEvent(editor, 'change', (_s, change) => void seen.push(change));
    const refresh = createDocumentRefresh(editor);
    const submission = await refresh.capture();
    await refresh.applyUpdate({ submission, sequence: 1, bytes: deck(['B']) });
    expect(seen.map((c) => c.source)).toEqual(['refresh']);
    done();
  });

  test('refuses local edits after capture, even through undo', async () => {
    const { editor, done } = await mount(deck(['A'], ['B']));
    const refresh = createDocumentRefresh(editor);
    const submission = await refresh.capture();
    editor.exec({ type: 'duplicateSlide', index: 0 });
    editor.exec({ type: 'undo' });
    const result = await refresh.applyUpdate({ submission, sequence: 1, bytes: deck(['X']) });
    expect(result.ok).toBe(false);
    if (!result.ok) expect(result.code).toBe('local-edits');
    expect(editor.snapshot().slides.length).toBe(2);
    done();
  });

  test('refuses invalid bytes before it changes anything', async () => {
    const { editor, done } = await mount(deck(['A'], ['B']));
    const refresh = createDocumentRefresh(editor);
    const submission = await refresh.capture();
    const result = await refresh.applyUpdate({
      submission,
      sequence: 1,
      bytes: new Uint8Array([1, 2, 3]),
    });
    expect(result).toMatchObject({ ok: false, code: 'invalid-document' });
    expect(editor.snapshot().slides.length).toBe(2);
    done();
  });

  test('orders cumulative results by sequence and marks only new changes', async () => {
    const { editor, done } = await mount(deck(['A'], ['B']));
    const refresh = createDocumentRefresh(editor);
    const submission = await refresh.capture();
    const first: RefreshChangeInput[] = [
      { id: 'one', description: 'First', location: { slideIndex: 0 } },
    ];
    const r1 = await refresh.applyUpdate({
      submission,
      sequence: 2,
      bytes: deck(['A1'], ['B']),
      changes: first,
    });
    expect(r1.ok && r1.changes.map((c) => [c.id, c.isNew])).toEqual([['one', true]]);
    const stale = await refresh.applyUpdate({ submission, sequence: 1, bytes: deck(['X']) });
    expect(stale).toMatchObject({ ok: false, code: 'out-of-order' });
    const second = [...first, { id: 'two', location: { slideIndex: 1 } }];
    const r2 = await refresh.applyUpdate({
      submission,
      sequence: 3,
      bytes: deck(['A1'], ['B1']),
      changes: second,
    });
    expect(r2.ok && r2.changes.map((c) => [c.id, c.isNew])).toEqual([
      ['one', false],
      ['two', true],
    ]);
    done();
  });

  test('refuses superseded, foreign, and cancelled submissions', async () => {
    const { editor, done } = await mount(deck(['A']));
    const refresh = createDocumentRefresh(editor);
    const old = await refresh.capture();
    const current = await refresh.capture();
    const superseded = await refresh.applyUpdate({
      submission: old,
      sequence: 1,
      bytes: deck(['X']),
    });
    expect(superseded).toMatchObject({ code: 'superseded' });
    const copy = { ...current };
    expect(
      await refresh.applyUpdate({ submission: copy, sequence: 1, bytes: deck(['X']) })
    ).toMatchObject({
      code: 'superseded',
    });
    refresh.cancel();
    expect(
      await refresh.applyUpdate({ submission: current, sequence: 1, bytes: deck(['X']) })
    ).toMatchObject({ code: 'superseded' });
    done();
  });

  test('a plain result reports unavailable change information', async () => {
    const { editor, done } = await mount(deck(['A']));
    const refresh = createDocumentRefresh(editor);
    const submission = await refresh.capture();
    const result = await refresh.applyUpdate({ submission, sequence: 1, bytes: deck(['B']) });
    expect(result).toMatchObject({ ok: true, changeInformation: 'unavailable', changes: [] });
    expect(refresh.highlightChanges()).toBe(0);
    done();
  });

  test('refuses malformed metadata with invalid-result', async () => {
    const { editor, done } = await mount(deck(['A']));
    const refresh = createDocumentRefresh(editor);
    const submission = await refresh.capture();
    const bad = [
      { id: '' },
      { id: 'x', description: 'y'.repeat(1001) },
      { id: 'x', location: { slideIndex: 'one' } },
      { id: 'x', unavailableReason: 'gone' },
    ];
    for (const change of bad) {
      const r = await refresh.applyUpdate({
        submission,
        sequence: 1,
        bytes: deck(['B']),
        changes: [change as never],
      });
      expect(r).toMatchObject({ ok: false, code: 'invalid-result' });
    }
    const dup = await refresh.applyUpdate({
      submission,
      sequence: 1,
      bytes: deck(['B']),
      changes: [{ id: 'a' }, { id: 'a' }],
    });
    expect(dup).toMatchObject({ code: 'invalid-result' });
    expect(slideText(editor, 0)).toBe('Shape 2');
    done();
  });
});

describe('document refresh: locations', () => {
  test('checks slides, shapes, paragraphs, and text, and reports why a location fails', async () => {
    const { editor, done } = await mount(deck(['A'], ['B']));
    const refresh = createDocumentRefresh(editor);
    const submission = await refresh.capture();
    const sldId = 256; // The builder numbers slide ids from 256.
    const result = await refresh.applyUpdate({
      submission,
      sequence: 1,
      bytes: deck(['Hello world'], ['B']),
      changes: [
        { id: 'slide', location: { slideIndex: 1 } },
        { id: 'by-id', location: { slideId: sldId } },
        { id: 'shape', location: { slideIndex: 0, shapeId: 2 } },
        {
          id: 'text',
          location: {
            slideIndex: 0,
            shapeId: 2,
            paragraphIndex: 0,
            start: 6,
            end: 11,
            text: 'world',
          },
        },
        {
          id: 'mismatch',
          location: {
            slideIndex: 0,
            shapeId: 2,
            paragraphIndex: 0,
            start: 0,
            end: 5,
            text: 'Howdy',
          },
        },
        {
          id: 'offsets',
          location: { slideIndex: 0, shapeId: 2, paragraphIndex: 0, start: 5, end: 99, text: 'x' },
        },
        { id: 'no-shape', location: { slideIndex: 0, shapeId: 77 } },
        { id: 'no-slide', location: { slideIndex: 9 } },
        { id: 'both', location: { slideIndex: 0, slideId: sldId } },
        { id: 'bad-id', location: { slideId: 3 } },
        { id: 'deleted', unavailableReason: 'deleted' },
        { id: 'nowhere', description: 'Updated the properties' },
      ],
    });
    if (!result.ok) throw new Error(result.code);
    const by = Object.fromEntries(
      result.changes.map((c) => [c.id, [c.status, c.diagnostic?.code]])
    );
    expect(by).toEqual({
      slide: ['available', undefined],
      'by-id': ['available', undefined],
      shape: ['available', undefined],
      text: ['available', undefined],
      mismatch: ['invalid', 'text-mismatch'],
      offsets: ['invalid', 'invalid-offsets'],
      'no-shape': ['invalid', 'shape-not-found'],
      'no-slide': ['invalid', 'slide-not-found'],
      both: ['invalid', 'invalid-selector'],
      'bad-id': ['invalid', 'invalid-slide-id'],
      deleted: ['deleted', 'deleted'],
      nowhere: ['unavailable', 'missing-location'],
    });
    expect(refresh.highlightChanges()).toBe(4);
    done();
  });

  test('the review module computes changes when the update has none', async () => {
    const computed: ComputedRefreshChange[] = [
      {
        id: 'c1',
        location: { slideIndex: 1 },
        computed: { scope: 'slide', action: 'changed', aspects: ['text'] },
      },
      { id: 'c2', computed: { scope: 'layout', action: 'changed', affectedSlides: [0, 1] } },
    ];
    let calls = 0;
    const module = reviewWith(() => ((calls += 1), computed));
    const { editor, done } = await mount(deck(['A'], ['B']), [module]);
    const refresh = createDocumentRefresh(editor);
    const submission = await refresh.capture();
    const result = await refresh.applyUpdate({
      submission,
      sequence: 1,
      bytes: deck(['A'], ['B2']),
    });
    expect(calls).toBe(1);
    expect(result).toMatchObject({ ok: true, changeInformation: 'available' });
    if (result.ok) {
      expect(result.changes.map((c) => c.status)).toEqual(['available', 'available']);
      expect(result.changes[0]!.computed).toEqual(computed[0]!.computed);
    }
    // Server metadata replaces the computed list.
    const r2 = await refresh.applyUpdate({
      submission,
      sequence: 2,
      bytes: deck(['A'], ['B3']),
      changes: [{ id: 'server' }],
    });
    expect(calls).toBe(1);
    expect(r2.ok && r2.changes.map((c) => c.id)).toEqual(['server']);
    done();
  });
});

describe('document refresh: edit mode', () => {
  test('adds one undo step above earlier history and reports source refresh', async () => {
    const { editor, done } = await mount(deck(['A'], ['B']));
    expect(editor.exec({ type: 'duplicateSlide', index: 0 }).ok).toBe(true);
    const seen: DocumentChange[] = [];
    onEditorEvent(editor, 'change', (_s, change) => void seen.push(change));
    const refresh = createDocumentRefresh(editor);
    const submission = await refresh.capture();
    const r1 = await refresh.applyUpdate({
      submission,
      sequence: 1,
      bytes: deck(['X'], ['Y']),
      mode: 'edit',
    });
    expect(r1).toMatchObject({ ok: true, mode: 'edit' });
    const r2 = await refresh.applyUpdate({
      submission,
      sequence: 2,
      bytes: deck(['X'], ['Y'], ['Z']),
      mode: 'edit',
    });
    expect(r2.ok).toBe(true);
    expect(editor.snapshot().slides.length).toBe(3);
    expect(seen.map((c) => c.source)).toEqual(['refresh', 'refresh']);
    // Each cumulative result is one undo step; the earlier edit stays below them.
    editor.exec({ type: 'undo' });
    expect(editor.snapshot().slides.length).toBe(2);
    editor.exec({ type: 'undo' });
    expect(editor.snapshot().slides.length).toBe(3);
    editor.exec({ type: 'undo' });
    expect(editor.snapshot().slides.length).toBe(2);
    done();
  });

  test('viewing mode refuses with edit-not-permitted and changes nothing', async () => {
    const { editor, done } = await mount(deck(['A']));
    const refresh = createDocumentRefresh(editor);
    const submission = await refresh.capture();
    editor.setMode('view');
    const r = await refresh.applyUpdate({
      submission,
      sequence: 1,
      bytes: deck(['X'], ['Y']),
      mode: 'edit',
    });
    expect(r).toMatchObject({ ok: false, code: 'edit-not-permitted', reason: 'viewing' });
    expect(editor.snapshot().slides.length).toBe(1);
    done();
  });

  test('an identical result succeeds without an undo step', async () => {
    const bytes = deck(['A']);
    const { editor, done } = await mount(bytes);
    const refresh = createDocumentRefresh(editor);
    const submission = await refresh.capture();
    const r = await refresh.applyUpdate({ submission, sequence: 1, bytes, mode: 'edit' });
    expect(r.ok).toBe(true);
    expect(editor.snapshot().canUndo).toBe(false);
    done();
  });
});

describe('document refresh: change flash', () => {
  test('flashes changed thumbnails and the changed shape, then clears on a local edit', async () => {
    const { editor, host, thumbs, done } = await mount(deck(['A', 'Keep'], ['B'], ['C']));
    const refresh = createDocumentRefresh(editor);
    const submission = await refresh.capture();
    const bytes = deck(['A changed', 'Keep'], ['B'], ['C changed']);
    const result = await refresh.applyUpdate({
      submission,
      sequence: 1,
      bytes,
      changes: [
        { id: 'shape', location: { slideIndex: 0, shapeId: 2 } },
        { id: 'slide', location: { slideIndex: 2 } },
      ],
    });
    expect(result.ok).toBe(true);
    const savedBefore = editor.save();
    expect(refresh.highlightChanges({ animation: false })).toBe(2);
    expect(refresh.snapshot().highlightsVisible).toBe(true);
    const marked = thumbs().map((t) => t.getAttribute('data-refresh-highlight'));
    expect(marked).toEqual(['shown', null, 'shown']);
    expect(thumbs()[0]!.parentElement!.getAttribute('aria-description')).toBe('Changed');
    const layer = host.querySelector('.pptx-surface__refresh');
    expect(layer?.querySelectorAll('.pptx-refresh-ring').length).toBe(1);
    // The flash is chrome: the saved file and the undo history do not change.
    expect(editor.save()).toEqual(savedBefore);
    expect(editor.snapshot().canUndo).toBe(false);
    editor.goToSlide(2);
    expect(host.querySelector('.pptx-refresh-ring--slide')).not.toBeNull();
    // A local edit removes the flash at once.
    editor.exec({ type: 'duplicateSlide', index: 0 });
    expect(thumbs().some((t) => t.hasAttribute('data-refresh-highlight'))).toBe(false);
    expect(host.querySelector('.pptx-surface__refresh')).toBeNull();
    expect(refresh.snapshot()).toMatchObject({ changes: [], highlightsVisible: false });
    done();
  });

  test('changeIds, includePrevious, and surfaces select what flashes; options are validated', async () => {
    const { editor, host, thumbs, done } = await mount(deck(['A'], ['B']));
    const refresh = createDocumentRefresh(editor);
    const submission = await refresh.capture();
    await refresh.applyUpdate({
      submission,
      sequence: 1,
      bytes: deck(['A1'], ['B']),
      changes: [{ id: 'one', location: { slideIndex: 0 } }],
    });
    await refresh.applyUpdate({
      submission,
      sequence: 2,
      bytes: deck(['A1'], ['B1']),
      changes: [
        { id: 'one', location: { slideIndex: 0 } },
        { id: 'two', location: { slideIndex: 1 } },
      ],
    });
    expect(refresh.highlightChanges({ animation: false })).toBe(1);
    expect(refresh.highlightChanges({ animation: false, includePrevious: true })).toBe(2);
    expect(refresh.highlightChanges({ animation: false, changeIds: ['one', 'missing'] })).toBe(1);
    expect(refresh.highlightChanges({ animation: false, changeIds: [] })).toBe(0);
    expect(
      refresh.highlightChanges({ animation: false, includePrevious: true, surfaces: ['canvas'] })
    ).toBe(2);
    expect(thumbs().some((t) => t.hasAttribute('data-refresh-highlight'))).toBe(false);
    expect(host.querySelector('.pptx-surface__refresh')).not.toBeNull();
    expect(() => refresh.highlightChanges({ opacity: 2 })).toThrow(RangeError);
    expect(() => refresh.highlightChanges({ timeoutMs: -1 })).toThrow(RangeError);
    expect(() => refresh.highlightChanges({ borderStyle: 'wavy' as never })).toThrow(TypeError);
    expect(() => refresh.highlightChanges({ animation: { easing: 'var(--x)' } })).toThrow(
      TypeError
    );
    expect(() => refresh.highlightChanges({ surfaces: ['rail' as never] })).toThrow(TypeError);
    refresh.clearHighlights({ animation: false });
    expect(host.querySelector('.pptx-surface__refresh')).toBeNull();
    done();
  });

  test('expiry clears the flash and keeps the change list', async () => {
    const { editor, thumbs, done } = await mount(deck(['A']));
    const refresh = createDocumentRefresh(editor);
    const submission = await refresh.capture();
    await refresh.applyUpdate({
      submission,
      sequence: 1,
      bytes: deck(['B']),
      changes: [{ id: 'one', location: { slideIndex: 0 } }],
    });
    refresh.highlightChanges({ animation: false, timeoutMs: 0 });
    await new Promise((r) => setTimeout(r, 10));
    expect(thumbs()[0]!.hasAttribute('data-refresh-highlight')).toBe(false);
    expect(refresh.snapshot().highlightsVisible).toBe(false);
    expect(refresh.snapshot().changes.length).toBe(1);
    done();
  });

  test('navigateToChange shows the slide and, with focus, selects the shape', async () => {
    const { editor, done } = await mount(deck(['A'], ['B', 'C']));
    const refresh = createDocumentRefresh(editor);
    const submission = await refresh.capture();
    await refresh.applyUpdate({
      submission,
      sequence: 1,
      bytes: deck(['A'], ['B', 'C2']),
      changes: [
        { id: 'shape', location: { slideIndex: 1, shapeId: 3 } },
        { id: 'gone', unavailableReason: 'deleted' },
      ],
    });
    expect(refresh.navigateToChange('gone')).toBe(false);
    expect(editor.snapshot().currentSlide).toBe(0);
    expect(refresh.navigateToChange('shape', { focus: true })).toBe(true);
    expect(editor.snapshot().currentSlide).toBe(1);
    const selection = editor.snapshot().selection;
    expect(selection.kind).toBe('shapes');
    expect(() => refresh.navigateToChange('shape', { offsetPx: -1 })).toThrow(RangeError);
    done();
  });
});
