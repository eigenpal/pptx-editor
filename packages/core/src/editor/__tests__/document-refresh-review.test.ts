// Document refresh regressions: the source of the edit after a reload, a file that only a load
// accepts, recovery after a failed load, forged editor marks, change fingerprints, computed
// changes from a server, input that runs when a result arrives (notes drafts, IME composition,
// pointer gestures), the selection and the focus after a reload, and the flash teardown.

import { afterAll, afterEach, describe, expect, test } from 'bun:test';
import { GlobalWindow } from 'happy-dom';
import {
  buildPptx,
  cell,
  para,
  run,
  shape,
  table,
} from '../../../../../scripts/fixtures/builder.ts';
import {
  createDocumentRefresh,
  createPptxEditor,
  describeRefreshChange,
  onEditorEvent,
  type ComputedRefreshChange,
  type EditorModule,
  type PptxEditorInstance,
  type RefreshChange,
} from '../index.ts';
import { elements, type Element } from '../../store/index.ts';
import { isStructureEdited } from '../../store/package/table-marks.ts';
import { REFRESH_INPUT_LIMITS } from '../refresh-input.ts';

const window = new GlobalWindow();
const doc = window.document as unknown as Document;
afterAll(() => window.happyDOM.close());
const g = globalThis as { ResizeObserver?: unknown };
g.ResizeObserver ??= (window as unknown as { ResizeObserver: unknown }).ResizeObserver;
const LIMITS = { ...REFRESH_INPUT_LIMITS };
afterEach(() => Object.assign(REFRESH_INPUT_LIMITS, LIMITS));

const box = (id: number, text: string, x = 400_000) =>
  shape({ id, x, y: 400_000, cx: 4_000_000, cy: 800_000, text: para(run(text)) });
const deck = (...slides: string[][]) =>
  buildPptx({
    slides: slides.map((texts) => ({ shapes: texts.map((t, i) => box(i + 2, t)).join('') })),
  });
const tick = (ms = 0) => new Promise((r) => setTimeout(r, ms));

/** An editor in a page laid out as the adapters lay it out, with a rail and a notes field. */
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
  editor.snapshot().slides.forEach((s) => {
    const li = doc.createElement('li');
    li.setAttribute('data-index', String(s.index));
    const thumb = doc.createElement('div');
    thumb.className = 'pptx-rail__thumb';
    li.append(thumb);
    rail.append(li);
    editor.renderThumbnail(s.index, thumb);
  });
  const done = () => {
    editor.destroy();
    shell.remove();
  };
  return { editor, host, shell, rail, done };
}

/** The review module with the changes the test gives. */
function reviewWith(changes: () => readonly ComputedRefreshChange[]): EditorModule {
  return {
    id: 'review',
    review: { displayModes: ['final'], collectReviewItems: () => [], refreshChanges: changes },
  };
}

const nameOf = (editor: PptxEditorInstance, i = 0) => editor.snapshot().shapes?.[i]?.name;
/** The node id of the shape whose `p:cNvPr/@id` is `id` (the fixtures name it `Shape <id>`). */
const nodeOf = (editor: PptxEditorInstance, id: number) => {
  const node = editor.snapshot().shapes?.find((s) => s.name === `Shape ${id}`)?.nodeId;
  if (!node) throw new Error(`no shape ${id}`);
  return node;
};

describe('document refresh review: change source', () => {
  test('the first edit after a reload in place has no source', async () => {
    const { editor, done } = await mount(deck(['A'], ['B']));
    const seen: unknown[] = [];
    const off = onEditorEvent(editor, 'change', (_s, c) => seen.push(c.source));
    const refresh = createDocumentRefresh(editor);
    const submission = await refresh.capture();
    const r = await refresh.applyUpdate({ submission, sequence: 1, bytes: deck(['Z'], ['B']) });
    expect(r.ok).toBe(true);
    expect(editor.exec({ type: 'duplicateSlide', index: 0 }).ok).toBe(true);
    expect(seen).toEqual(['refresh', undefined]);
    off();
    done();
  });

  test('an identical reload is a refresh change with a new load id; the next edit has no source', async () => {
    const { editor, done } = await mount(deck(['A']));
    const seen: unknown[] = [];
    const off = onEditorEvent(editor, 'change', (_s, c) => seen.push(c.source));
    const refresh = createDocumentRefresh(editor);
    const submission = await refresh.capture();
    const before = editor.getDocumentHandle().loadId;
    expect((await refresh.applyUpdate({ submission, sequence: 1, bytes: deck(['A']) })).ok).toBe(
      true
    );
    expect(editor.getDocumentHandle().loadId).not.toBe(before);
    expect(editor.exec({ type: 'duplicateSlide', index: 0 }).ok).toBe(true);
    expect(seen).toEqual(['refresh', undefined]);
    off();
    done();
  });
});

describe('document refresh review: a file only a load accepts', () => {
  test('duplicate shape ids: edit refuses, reload opens the file whole', async () => {
    const { editor, done } = await mount(
      buildPptx({ slides: [{ shapes: box(2, 'A') + box(3, 'B') }] })
    );
    const next = buildPptx({ slides: [{ shapes: box(2, 'A') + box(2, 'B2') }] });
    const refresh = createDocumentRefresh(editor);
    const submission = await refresh.capture();
    const edit = await refresh.applyUpdate({ submission, sequence: 1, bytes: next, mode: 'edit' });
    expect(edit.ok ? 'ok' : edit.code).toBe('edit-unsupported');
    const seen: unknown[] = [];
    const off = onEditorEvent(editor, 'change', (_s, c) => seen.push(c.source));
    const reload = await refresh.applyUpdate({ submission, sequence: 2, bytes: next });
    expect(reload.ok ? reload.mode : reload.code).toBe('reload');
    expect(nameOf(editor, 1)).toBeDefined();
    expect(refresh.snapshot().recoveryAvailable).toBe(false);
    expect(seen[seen.length - 1]).toBe('refresh');
    off();
    done();
  });
});

describe('document refresh review: recovery', () => {
  test('a load from outside during the quiet load gives document-changed and no recovery', async () => {
    const { editor, done } = await mount(deck(['A']));
    const refresh = createDocumentRefresh(editor);
    const submission = await refresh.capture();
    const load = editor.load.bind(editor);
    let other: Promise<unknown> = Promise.resolve();
    editor.load = (bytes, options) => {
      const r = load(bytes, options);
      if (!options?.quiet) return r;
      other = load(deck(['Other']));
      return r.then(async (x) => (await other, x));
    };
    // Duplicate shape ids reload with `load()`.
    const bytes = buildPptx({ slides: [{ shapes: box(2, 'Z') + box(2, 'Y') }] });
    const r = await refresh.applyUpdate({ submission, sequence: 1, bytes });
    await other;
    editor.load = load;
    expect(r.ok ? 'ok' : r.code).toBe('document-changed');
    expect(refresh.snapshot().recoveryAvailable).toBe(false);
    expect(refresh.recoveryBytes()).toBeNull();
    expect(await refresh.recover()).toBe(false);
    refresh.cancel();
    expect(refresh.snapshot().phase).not.toBe('failed');
    await refresh.capture();
    done();
  });

  test('a quiet load that fails keeps the deck: load-failed, recovered, no recovery bytes', async () => {
    const { editor, done } = await mount(deck(['A']));
    const refresh = createDocumentRefresh(editor);
    const submission = await refresh.capture();
    const load = editor.load.bind(editor);
    editor.load = async (bytes, options) =>
      options?.quiet ? { ok: false, reason: 'errors.unsupported' } : load(bytes, options);
    const bytes = buildPptx({ slides: [{ shapes: box(2, 'Z') + box(2, 'Y') }] });
    const r = await refresh.applyUpdate({ submission, sequence: 1, bytes });
    editor.load = load;
    expect(r.ok ? 'ok' : [r.code, r.recovered]).toEqual(['load-failed', true]);
    expect(nameOf(editor)).toBeDefined();
    expect(refresh.snapshot().recoveryAvailable).toBe(false);
    await refresh.capture();
    done();
  });

  test('no document after a failed load: recovery-failed, then recover() opens the bytes', async () => {
    const { editor, done } = await mount(deck(['Kept']));
    const refresh = createDocumentRefresh(editor);
    const submission = await refresh.capture();
    const load = editor.load.bind(editor);
    const real = Object.getOwnPropertyDescriptor(editor, 'session')!;
    const hide = () =>
      Object.defineProperty(editor, 'session', { get: () => null, configurable: true });
    const show = () => Object.defineProperty(editor, 'session', real);
    // The open deck is gone, and every load fails.
    editor.load = async () => {
      hide();
      return { ok: false, reason: 'errors.unsupported' };
    };
    const bytes = buildPptx({ slides: [{ shapes: box(2, 'Z') + box(2, 'Y') }] });
    const r = await refresh.applyUpdate({ submission, sequence: 1, bytes });
    expect(r.ok ? 'ok' : [r.code, r.recovered]).toEqual(['recovery-failed', false]);
    expect(refresh.snapshot().recoveryAvailable).toBe(true);
    expect(refresh.recoveryBytes()).not.toBeNull();
    await expect(refresh.capture()).rejects.toMatchObject({ code: 'recovery-failed' });
    // A cancel while no document shows keeps the bytes.
    refresh.cancel();
    expect(refresh.snapshot().recoveryAvailable).toBe(true);
    // The real load opens the bytes; the deck shows again once it mounts.
    editor.load = async (bytes, options) => (show(), load(bytes, options));
    expect(await refresh.recover()).toBe(true);
    editor.load = load;
    expect(refresh.snapshot().recoveryAvailable).toBe(false);
    expect(editor.snapshot().shapes?.map((s) => s.name).length).toBeGreaterThan(0);
    await refresh.capture();
    done();
  });

  test('cancel in the pause after a reload commit leaves no recovery', async () => {
    const { editor, done } = await mount(deck(['A']));
    const refresh = createDocumentRefresh(editor);
    const submission = await refresh.capture();
    const off = onEditorEvent(editor, 'change', (_s, c) => {
      if (c.source === 'refresh') queueMicrotask(() => refresh.cancel());
    });
    const r = await refresh.applyUpdate({ submission, sequence: 1, bytes: deck(['Z']) });
    off();
    expect(r.ok ? 'ok' : r.code).toBe('cancelled');
    expect(refresh.snapshot().phase).not.toBe('failed');
    expect(refresh.snapshot().recoveryAvailable).toBe(false);
    expect(refresh.recoveryBytes()).toBeNull();
    await refresh.capture();
    done();
  });
});

describe('document refresh review: returned files read as an open reads them', () => {
  const MARK = 'urn:pptx-editor.dev:table-structure-edit';
  const tbl = (mark: boolean, t: string) => {
    let x = table({
      id: 4,
      x: 0,
      y: 0,
      cx: 1_000_000,
      cy: 400_000,
      columns: [500_000, 500_000],
      tblPr: '<a:tblPr/>',
      rows: `<a:tr h="400000">${cell(t)}${cell('b')}</a:tr>`,
    });
    if (mark)
      x = x.replace(
        /(<p:cNvPr id="4" name="[^"]*")\/>/,
        `$1><a:extLst><a:ext uri="${MARK}"/></a:extLst></p:cNvPr>`
      );
    return x;
  };
  const frames = (root: Element, out: Element[] = []): Element[] => {
    for (const c of elements(root)) {
      if (c.local === 'graphicFrame') out.push(c);
      frames(c, out);
    }
    return out;
  };
  const marked = (editor: PptxEditorInstance) => {
    const part = editor.session!.store.package.parts.get('/ppt/slides/slide1.xml');
    if (part?.kind !== 'xml') throw new Error('no slide');
    return frames(part.xml.root).map((f) => isStructureEdited(f));
  };

  test('replaceDocument and a reload drop a forged structure mark, as load() does', async () => {
    const forged = buildPptx({ slides: [{ shapes: tbl(true, 'z') }] });
    const { editor, done } = await mount(buildPptx({ slides: [{ shapes: tbl(false, 'a') }] }));
    expect(editor.exec({ type: 'replaceDocument', pptx: forged }).ok).toBe(true);
    expect(marked(editor)).toEqual([false]);
    const refresh = createDocumentRefresh(editor);
    const submission = await refresh.capture();
    const next = buildPptx({ slides: [{ shapes: tbl(true, 'y') }] });
    expect((await refresh.applyUpdate({ submission, sequence: 1, bytes: next })).ok).toBe(true);
    expect(marked(editor)).toEqual([false]);
    done();
  });
});

describe('document refresh review: changes', () => {
  test('a paragraph index out of range is paragraph-not-found', async () => {
    const { editor, done } = await mount(deck(['A']));
    const refresh = createDocumentRefresh(editor);
    const submission = await refresh.capture();
    const r = await refresh.applyUpdate({
      submission,
      sequence: 1,
      bytes: deck(['B']),
      changes: [{ id: 'p', location: { slideIndex: 0, shapeId: 2, paragraphIndex: 4 } }],
    });
    if (!r.ok) throw new Error(r.code);
    expect(r.changes[0]?.diagnostic?.code).toBe('paragraph-not-found');
    done();
  });

  test('a change that only moved with an inserted slide is not new', async () => {
    let changes: ComputedRefreshChange[] = [];
    const { editor, done } = await mount(deck(['One'], ['Hello there']), [
      reviewWith(() => changes),
    ]);
    const refresh = createDocumentRefresh(editor);
    const submission = await refresh.capture();
    const text = (slideIndex: number): ComputedRefreshChange => ({
      id: 'text-1',
      location: { slideIndex, shapeId: 2, paragraphIndex: 0, start: 6, end: 11, text: 'world' },
      computed: { scope: 'text', action: 'changed', aspects: ['text'] },
    });
    changes = [text(1)];
    const r1 = await refresh.applyUpdate({
      submission,
      sequence: 1,
      mode: 'edit',
      bytes: deck(['One'], ['Hello world']),
    });
    if (!r1.ok) throw new Error(r1.code);
    expect(r1.changes.map((c) => [c.id, c.isNew])).toEqual([['text-1', true]]);
    changes = [
      { id: 'slide-0', location: { slideIndex: 0 }, computed: { scope: 'slide', action: 'added' } },
      text(2),
    ];
    const r2 = await refresh.applyUpdate({
      submission,
      sequence: 2,
      mode: 'edit',
      bytes: deck(['Intro'], ['One'], ['Hello world']),
    });
    if (!r2.ok) throw new Error(r2.code);
    expect(r2.changes.map((c) => [c.id, c.status, c.isNew])).toEqual([
      ['slide-0', 'available', true],
      ['text-1', 'available', false],
    ]);
    done();
  });

  test('computed changes passed to applyUpdate keep computed, targets, and descriptions', async () => {
    const { editor, done } = await mount(deck(['A'], ['B']));
    const refresh = createDocumentRefresh(editor);
    const submission = await refresh.capture();
    // What `comparePresentations()` returns for a layout change and a text change.
    const changes: ComputedRefreshChange[] = [
      {
        id: 'layout:/ppt/slideLayouts/slideLayout1.xml',
        computed: { scope: 'layout', action: 'changed', affectedSlides: [0, 1] },
      },
      {
        id: 'text:2',
        location: { slideIndex: 1, shapeId: 2, paragraphIndex: 0, start: 0, end: 2, text: 'B2' },
        computed: { scope: 'text', action: 'changed', aspects: ['text'] },
      },
    ];
    const r = await refresh.applyUpdate({
      submission,
      sequence: 1,
      bytes: deck(['A'], ['B2']),
      changes,
    });
    if (!r.ok) throw new Error(r.code);
    expect(r.changes.map((c) => [c.id, c.status, c.computed?.scope])).toEqual([
      ['layout:/ppt/slideLayouts/slideLayout1.xml', 'available', 'layout'],
      ['text:2', 'available', 'text'],
    ]);
    expect(r.changes[0]?.computed?.affectedSlides).toEqual([0, 1]);
    expect(describeRefreshChange(r.changes[0]!)).toBe('Changed a layout. Slides that use it: 2');
    expect(describeRefreshChange(r.changes[1]!)).toBe('Changed text on slide 2');
    expect(refresh.highlightChanges({ animation: false })).toBe(2);
    expect(Object.isFrozen(r.changes[0]?.computed)).toBe(true);
    done();
  });

  test('a malformed computed field is invalid-result', async () => {
    const { editor, done } = await mount(deck(['A']));
    const refresh = createDocumentRefresh(editor);
    const submission = await refresh.capture();
    const bad = [
      { scope: 'galaxy', action: 'changed' },
      { scope: 'text', action: 'changed', aspects: ['nope'] },
      { scope: 'layout', action: 'changed', affectedSlides: [-1] },
      { scope: 'slide', action: 'moved', previousSlideIndex: 1.5 },
      'text',
    ];
    for (const computed of bad) {
      const changes = [{ id: 'x', computed }] as unknown as ComputedRefreshChange[];
      const r = await refresh.applyUpdate({ submission, sequence: 1, bytes: deck(['B']), changes });
      expect(r.ok ? 'ok' : r.code).toBe('invalid-result');
    }
    const visible = [
      { id: 'v', computed: { scope: 'shape', action: 'changed', aspects: ['visibility'] } },
    ];
    const r = await refresh.applyUpdate({
      submission,
      sequence: 1,
      bytes: deck(['B']),
      changes: visible as unknown as ComputedRefreshChange[],
    });
    expect(r.ok).toBe(true);
    done();
  });
});

describe('document refresh review: input that runs when a result arrives', () => {
  /** A notes field as the adapters draw it: a draft until a press elsewhere commits it. */
  function notesField(shell: HTMLElement, editor: PptxEditorInstance) {
    const pane = doc.createElement('div');
    pane.className = 'pptx-notes';
    const field = doc.createElement('div');
    field.className = 'pptx-notes__text';
    field.tabIndex = 0;
    field.textContent = editor.snapshot().notes;
    pane.append(field);
    shell.append(pane);
    const onDown = (e: Event) => {
      if (field.contains(e.target as Node)) return;
      if (field.textContent !== editor.snapshot().notes)
        editor.exec({
          type: 'setNotes',
          index: editor.snapshot().currentSlide,
          text: field.textContent ?? '',
        });
    };
    doc.addEventListener('pointerdown', onDown, true);
    field.focus();
    return { field, off: () => doc.removeEventListener('pointerdown', onDown, true) };
  }

  test('capture commits a notes draft', async () => {
    const { editor, shell, done } = await mount(deck(['A']));
    const { field, off } = notesField(shell, editor);
    field.textContent = 'typed before capture';
    const refresh = createDocumentRefresh(editor);
    await refresh.capture();
    expect(editor.snapshot().notes).toBe('typed before capture');
    off();
    done();
  });

  test('notes typed after capture: the draft commits and the result refuses with local-edits', async () => {
    const { editor, shell, done } = await mount(deck(['A']));
    const { field, off } = notesField(shell, editor);
    const refresh = createDocumentRefresh(editor);
    const submission = await refresh.capture();
    field.textContent = 'typing across the update';
    const r = await refresh.applyUpdate({ submission, sequence: 1, bytes: deck(['Server']) });
    expect(r.ok ? 'ok' : r.code).toBe('local-edits');
    expect(editor.snapshot().notes).toBe('typing across the update');
    expect(nameOf(editor)).toBeDefined();
    off();
    done();
  });

  test('an IME composition after capture: the result waits, then refuses with local-edits', async () => {
    const { editor, host, done } = await mount(deck(['A']));
    const input = host.querySelector<HTMLElement>('.pptx-surface__input')!;
    const refresh = createDocumentRefresh(editor);
    const submission = await refresh.capture();
    input.classList.add('pptx-surface__input--composing');
    let settled = false;
    const pending = refresh
      .applyUpdate({ submission, sequence: 1, bytes: deck(['Server']) })
      .then((r) => ((settled = true), r));
    await tick(10);
    expect(settled).toBe(false);
    // The composition writes its text, then ends.
    expect(editor.exec({ type: 'duplicateSlide', index: 0 }).ok).toBe(true);
    input.classList.remove('pptx-surface__input--composing');
    input.dispatchEvent(new window.Event('compositionend') as unknown as Event);
    const r = await pending;
    expect(r.ok ? 'ok' : r.code).toBe('local-edits');
    done();
  });

  test('a composition that outlasts its limit refuses the result and fails the capture', async () => {
    REFRESH_INPUT_LIMITS.compositionMs = 20;
    const { editor, host, done } = await mount(deck(['A']));
    const input = host.querySelector<HTMLElement>('.pptx-surface__input')!;
    const refresh = createDocumentRefresh(editor);
    const submission = await refresh.capture();
    input.classList.add('pptx-surface__input--composing');
    const r = await refresh.applyUpdate({ submission, sequence: 1, bytes: deck(['Server']) });
    expect(r.ok ? 'ok' : r.code).toBe('local-edits');
    await expect(refresh.capture()).rejects.toMatchObject({ code: 'input-failed' });
    input.classList.remove('pptx-surface__input--composing');
    await refresh.capture();
    done();
  });

  test('a result waits for a pointer gesture on the slide to end', async () => {
    const { editor, host, done } = await mount(deck(['A']));
    const stage = host.querySelector<HTMLElement>('.pptx-surface__stage')!;
    const refresh = createDocumentRefresh(editor);
    const submission = await refresh.capture();
    const press = (type: string, target: EventTarget) =>
      target.dispatchEvent(
        new window.PointerEvent(type, { bubbles: true, pointerId: 7 }) as unknown as Event
      );
    press('pointerdown', stage);
    let settled = false;
    const pending = refresh
      .applyUpdate({ submission, sequence: 1, bytes: deck(['Server']) })
      .then((r) => ((settled = true), r));
    await tick(10);
    expect(settled).toBe(false);
    press('pointerup', doc);
    const r = await pending;
    expect(r.ok).toBe(true);
    // A drag that edits: the result refuses.
    const second = await refresh.capture();
    press('pointerdown', stage);
    const later = refresh.applyUpdate({ submission: second, sequence: 1, bytes: deck(['Again']) });
    await tick(5);
    expect(editor.exec({ type: 'duplicateSlide', index: 0 }).ok).toBe(true);
    press('pointerup', doc);
    const refused = await later;
    expect(refused.ok ? 'ok' : refused.code).toBe('local-edits');
    done();
  });
});

describe('document refresh review: selection and focus after a reload', () => {
  test('a text caret in a kept shape stays, clamped into its paragraph', async () => {
    const { editor, done } = await mount(deck(['Hello world', 'Other']));
    const id = nodeOf(editor, 2);
    const at = { paragraph: 0, offset: 9 };
    editor.setSelection({ kind: 'text', nodeId: id, anchor: at, focus: at });
    const refresh = createDocumentRefresh(editor);
    const submission = await refresh.capture();
    const r = await refresh.applyUpdate({ submission, sequence: 1, bytes: deck(['Hi', 'Other']) });
    expect(r.ok).toBe(true);
    const sel = editor.snapshot().selection;
    if (sel.kind !== 'text') throw new Error(`selection ${sel.kind}`);
    expect(sel.nodeId).toBe(id);
    expect(sel.focus.offset).toBeLessThanOrEqual(2);
    done();
  });

  test('a removed shape clears the selection and moves the focus off the hidden input', async () => {
    const { editor, host, done } = await mount(deck(['Hello', 'Other']));
    const id = nodeOf(editor, 3);
    const at = { paragraph: 0, offset: 2 };
    editor.setSelection({ kind: 'text', nodeId: id, anchor: at, focus: at });
    const input = host.querySelector<HTMLElement>('.pptx-surface__input')!;
    input.focus();
    expect(doc.activeElement).toBe(input);
    const refresh = createDocumentRefresh(editor);
    const submission = await refresh.capture();
    const r = await refresh.applyUpdate({ submission, sequence: 1, bytes: deck(['Hello']) });
    expect(r.ok).toBe(true);
    expect(editor.snapshot().selection.kind).toBe('none');
    expect(doc.activeElement).not.toBe(input);
    expect(doc.activeElement).toBe(host.querySelector('.pptx-surface'));
    done();
  });
});

describe('document refresh review: the change flash', () => {
  test('destroy tears the flash down; an adapter description comes back; colors load no URL', async () => {
    const { editor, rail, shell } = await mount(deck(['A'], ['B']));
    const item = rail.querySelector('[data-index="1"]')!;
    item.setAttribute('aria-description', 'Section start');
    const refresh = createDocumentRefresh(editor);
    for (const color of ['url(https://x.test/a.png)', 'image-set("a.png" 1x)', 'u\\72l(x)'])
      expect(() => refresh.highlightChanges({ color })).toThrow(TypeError);
    expect(() => refresh.highlightChanges({ borderColor: 'URL( x )' })).toThrow(TypeError);
    const submission = await refresh.capture();
    await refresh.applyUpdate({
      submission,
      sequence: 1,
      bytes: deck(['A'], ['B2']),
      changes: [{ id: 's', location: { slideIndex: 1 } }],
    });
    expect(refresh.highlightChanges({ animation: false, timeoutMs: null })).toBe(1);
    expect(item.getAttribute('aria-description')).toBe('Changed');
    refresh.clearHighlights({ animation: false });
    expect(item.getAttribute('aria-description')).toBe('Section start');
    expect(refresh.highlightChanges({ animation: false, timeoutMs: null })).toBe(1);
    editor.destroy();
    expect(shell.querySelectorAll('[data-refresh-highlight]').length).toBe(0);
    expect(item.getAttribute('aria-description')).toBe('Section start');
    // A thumbnail that mounts after destroy gets no flash.
    const li = doc.createElement('li');
    li.setAttribute('data-index', '1');
    li.innerHTML = '<div class="pptx-rail__thumb"></div>';
    rail.append(li);
    await tick(5);
    expect(shell.querySelectorAll('[data-refresh-highlight]').length).toBe(0);
    shell.remove();
  });

  test('the slide master view draws no canvas ring', async () => {
    const { editor, host, done } = await mount(deck(['A'], ['B'], ['C'], ['D']));
    const refresh = createDocumentRefresh(editor);
    const submission = await refresh.capture();
    // Every slide changed: an index in the master view names a slide too.
    await refresh.applyUpdate({
      submission,
      sequence: 1,
      bytes: deck(['A2'], ['B2'], ['C2'], ['D2']),
      changes: [0, 1, 2, 3].map((i) => ({ id: `s${i}`, location: { slideIndex: i } })),
    });
    expect(refresh.highlightChanges({ animation: false, timeoutMs: null })).toBe(4);
    expect(host.querySelectorAll('.pptx-refresh-ring').length).toBe(1);
    expect(editor.exec({ type: 'setPresentationView', view: 'slideMaster' }).ok).toBe(true);
    expect(host.querySelectorAll('.pptx-refresh-ring').length).toBe(0);
    done();
  });
});

describe('describeRefreshChange', () => {
  const change = (id: string, computed: RefreshChange['computed'], slideIndex?: number) =>
    ({
      id,
      resultId: 'r',
      isNew: true,
      status: 'available',
      computed,
      ...(slideIndex === undefined ? {} : { location: { slideIndex } }),
    }) as RefreshChange;

  test('removed notes and a truncated list have their own summaries', () => {
    expect(describeRefreshChange(change('n', { scope: 'notes', action: 'removed' }, 2))).toBe(
      'Removed the notes of slide 3'
    );
    expect(
      describeRefreshChange(change('truncated', { scope: 'package', action: 'changed' }))
    ).toBe('Other changes are not listed');
    expect(describeRefreshChange(change('package', { scope: 'package', action: 'changed' }))).toBe(
      'Changed the presentation settings'
    );
  });
});
