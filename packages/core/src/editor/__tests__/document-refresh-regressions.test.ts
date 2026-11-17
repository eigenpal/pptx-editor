// Document refresh regressions, second set: a handle drag that the overlay draws again, the
// slide on screen when slides come before it, a flash that continues through cumulative
// results, `capture()` during a quiet load, embedded fonts in edit mode, the cause of the next
// load, `cancel()` during an input wait, and the press that commits a notes draft.

import { afterAll, afterEach, describe, expect, test } from 'bun:test';
import { GlobalWindow } from 'happy-dom';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { buildPptx, para, run, shape } from '../../../../../scripts/fixtures/builder.ts';
import { refreshDeck, type RefreshEdits } from '../../../../../scripts/bench/refresh-deck.ts';
import {
  createDocumentRefresh,
  createPptxEditor,
  onEditorEvent,
  type EditorModule,
} from '../index.ts';
import { COMMIT_POINTER_ID, REFRESH_INPUT_LIMITS, gestureActive } from '../refresh-input.ts';

const window = new GlobalWindow();
const doc = window.document as unknown as Document;
afterAll(() => window.happyDOM.close());
const g = globalThis as { ResizeObserver?: unknown };
g.ResizeObserver ??= (window as unknown as { ResizeObserver: unknown }).ResizeObserver;
const LIMITS = { ...REFRESH_INPUT_LIMITS };
afterEach(() => Object.assign(REFRESH_INPUT_LIMITS, LIMITS));

const box = (id: number, text: string) =>
  shape({ id, x: 400_000, y: 400_000, cx: 4_000_000, cy: 800_000, text: para(run(text)) });
const deck = (...slides: string[]) =>
  buildPptx({ slides: slides.map((t) => ({ shapes: box(2, t) })) });
const tick = (ms = 0) => new Promise((r) => setTimeout(r, ms));
const MARK = 'data-refresh-highlight';

/** An editor in a page laid out as the adapters lay it out, with a rail of thumbnails. */
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
  /** Draw one rail item per slide, as the adapters do after each change. */
  const drawRail = () => {
    rail.replaceChildren();
    editor.snapshot().slides.forEach((s) => {
      const li = doc.createElement('li');
      li.setAttribute('data-index', String(s.index));
      const thumb = doc.createElement('div');
      thumb.className = 'pptx-rail__thumb';
      li.append(thumb);
      rail.append(li);
    });
  };
  drawRail();
  const done = () => {
    editor.destroy();
    shell.remove();
  };
  return { editor, host, shell, rail, drawRail, done };
}

const pointer = (type: string, init: Record<string, unknown> = {}) =>
  new window.PointerEvent(type, { bubbles: true, ...init }) as unknown as Event;

describe('document refresh regressions: input', () => {
  test('a handle drag stays a gesture after the overlay draws its handles again', async () => {
    const { editor, host, done } = await mount(deck('A'));
    const refresh = createDocumentRefresh(editor);
    editor.setSelection({ kind: 'shapes', nodeIds: [editor.snapshot().shapes![0]!.nodeId] });
    await tick(20);
    const handle = host.querySelector('[data-handle]')!;
    const submission = await refresh.capture();
    handle.dispatchEvent(pointer('pointerdown', { pointerId: 3, button: 0, buttons: 1 }));
    doc.dispatchEvent(pointer('pointermove', { pointerId: 3, buttons: 1, clientX: 30 }));
    await tick(20);
    // The overlay drew its handles again: the pressed handle left the page.
    expect(host.contains(handle)).toBe(false);
    expect(gestureActive(editor)).toBe(true);
    let settled = false;
    const pending = refresh
      .applyUpdate({ submission, sequence: 1, bytes: deck('Server') })
      .then((r) => ((settled = true), r));
    await tick(20);
    expect(settled).toBe(false);
    doc.dispatchEvent(pointer('pointerup', { pointerId: 3 }));
    const r = await pending;
    expect(r.ok).toBe(true);
    expect(gestureActive(editor)).toBe(false);
    done();
  });

  test('cancel() ends a wait for a gesture and a composition at once: cancelled', async () => {
    const { editor, host, done } = await mount(deck('A'));
    const refresh = createDocumentRefresh(editor);
    const stage = host.querySelector<HTMLElement>('.pptx-surface__stage')!;
    let submission = await refresh.capture();
    stage.dispatchEvent(pointer('pointerdown', { pointerId: 9 }));
    let t0 = performance.now();
    let pending = refresh.applyUpdate({ submission, sequence: 1, bytes: deck('Server') });
    await tick(10);
    refresh.cancel();
    let r = await pending;
    expect(r.ok ? 'ok' : r.code).toBe('cancelled');
    expect(performance.now() - t0).toBeLessThan(500);
    doc.dispatchEvent(pointer('pointerup', { pointerId: 9 }));
    // A composition in the slide's text input.
    submission = await refresh.capture();
    const input = host.querySelector<HTMLElement>('.pptx-surface__input')!;
    input.classList.add('pptx-surface__input--composing');
    t0 = performance.now();
    pending = refresh.applyUpdate({ submission, sequence: 1, bytes: deck('Server') });
    await tick(10);
    refresh.cancel();
    r = await pending;
    expect(r.ok ? 'ok' : r.code).toBe('cancelled');
    expect(performance.now() - t0).toBeLessThan(500);
    // A capture that waits for the composition ends as superseded.
    const capture = refresh.capture();
    await tick(10);
    refresh.cancel();
    await expect(capture).rejects.toMatchObject({ code: 'superseded' });
    input.classList.remove('pptx-surface__input--composing');
    expect(editor.snapshot().shapes?.length).toBe(1);
    done();
  });

  test('the notes commit press fires only for a real draft, with its own pointer id', async () => {
    const { editor, shell, done } = await mount(deck('A'));
    expect(editor.exec({ type: 'setNotes', index: 0, text: 'Line one\nLine two' }).ok).toBe(true);
    const pane = doc.createElement('div');
    pane.className = 'pptx-notes';
    const field = doc.createElement('div');
    field.className = 'pptx-notes__text';
    field.tabIndex = 0;
    pane.append(field);
    shell.append(pane);
    field.focus();
    const presses: number[] = [];
    const onDown = (e: Event) => presses.push((e as PointerEvent).pointerId);
    doc.addEventListener('pointerdown', onDown, true);
    const refresh = createDocumentRefresh(editor);
    // The field shows the committed notes in the form `setNotes` writes: no draft, no press.
    field.textContent = 'Line one\r\nLine two';
    const first = await refresh.capture();
    refresh.finish(first);
    expect(presses).toEqual([]);
    // A draft: one press, with an id that no real pointer has, and no gesture after it.
    field.textContent = 'Line one\nLine two, typed';
    const second = await refresh.capture();
    refresh.finish(second);
    expect(presses).toEqual([COMMIT_POINTER_ID]);
    expect(gestureActive(editor)).toBe(false);
    doc.removeEventListener('pointerdown', onDown, true);
    done();
  });
});

describe('document refresh regressions: the slide on screen', () => {
  for (const mode of ['reload', 'edit'] as const) {
    test(`${mode}: a slide inserted before the slide on screen keeps that slide and its caret`, async () => {
      const { editor, host, done } = await mount(refreshDeck(10));
      editor.goToSlide(5);
      const title = editor.snapshot().shapes!.find((s) => s.name === 'Title')!.nodeId;
      const at = { paragraph: 0, offset: 'Slide 6 title'.length };
      editor.setSelection({ kind: 'text', nodeId: title, anchor: at, focus: at });
      const refresh = createDocumentRefresh(editor);
      const submission = await refresh.capture();
      const bytes = refreshDeck(10, { insertAfter: 1 });
      const r = await refresh.applyUpdate({ submission, sequence: 1, bytes, mode });
      expect(r.ok).toBe(true);
      expect(editor.snapshot().currentSlide).toBe(6);
      const sel = editor.snapshot().selection;
      expect(sel).toMatchObject({ kind: 'text', nodeId: title, focus: at });
      // A typed key goes to the same title.
      const input = host.querySelector<HTMLTextAreaElement>('.pptx-surface__input')!;
      input.dispatchEvent(
        new window.InputEvent('beforeinput', {
          inputType: 'insertText',
          data: '!',
          bubbles: true,
          cancelable: true,
        }) as unknown as Event
      );
      const saved = unzipSync(editor.save()!);
      expect(strFromU8(saved['ppt/slides/slide6.xml']!)).toContain('Slide 6 title!');
      done();
    });
  }

  test('a removed slide on screen leaves the index, clamped, and clears the selection', async () => {
    const { editor, done } = await mount(deck('A', 'B', 'C'));
    editor.goToSlide(2);
    editor.setSelection({ kind: 'shapes', nodeIds: [editor.snapshot().shapes![0]!.nodeId] });
    const refresh = createDocumentRefresh(editor);
    const submission = await refresh.capture();
    const r = await refresh.applyUpdate({ submission, sequence: 1, bytes: deck('A', 'B') });
    expect(r.ok).toBe(true);
    expect(editor.snapshot().currentSlide).toBe(1);
    expect(editor.snapshot().selection.kind).toBe('none');
    done();
  });
});

describe('document refresh regressions: a flash through cumulative results', () => {
  /** Apply cumulative results to `edits`, each with a change on every edited slide. */
  async function results(mode: 'reload' | 'edit') {
    const m = await mount(refreshDeck(8));
    m.editor.goToSlide(2);
    const refresh = createDocumentRefresh(m.editor);
    const submission = await refresh.capture();
    let sequence = 0;
    const apply = async (edits: RefreshEdits, slides: number[]) => {
      const r = await refresh.applyUpdate({
        submission,
        sequence: ++sequence,
        mode,
        bytes: refreshDeck(8, edits),
        changes: slides.map((i) => ({ id: `s${i}`, location: { slideIndex: i, shapeId: 3 } })),
      });
      if (!r.ok) throw new Error(r.code);
      m.drawRail();
      return r;
    };
    return { ...m, refresh, apply };
  }

  for (const mode of ['reload', 'edit'] as const) {
    test(`${mode}: a kept place never leaves or enters again; a new place enters`, async () => {
      const { editor, rail, host, refresh, apply, done } = await results(mode);
      await apply({ text: [2] }, [2]);
      expect(refresh.highlightChanges({ timeoutMs: 60_000 })).toBe(1);
      await tick(40);
      const thumb = (i: number) => rail.querySelector(`[data-index="${i}"] .pptx-rail__thumb`);
      expect(thumb(2)?.getAttribute(MARK)).toBe('shown');
      const ring = () => host.querySelector('.pptx-surface__refresh > *');
      expect(ring()).not.toBeNull();
      // Watch the canvas mark of slide 3 through the next result.
      const lost: string[] = [];
      const off = editor.subscribe(() => {
        if (!ring()) lost.push('canvas');
      });
      const r = await apply({ text: [2, 5] }, [2, 5]);
      off();
      expect(r.changes.map((c) => [c.id, c.isNew])).toEqual([
        ['s2', false],
        ['s5', true],
      ]);
      expect(lost).toEqual([]);
      expect(refresh.snapshot().highlightsVisible).toBe(true);
      expect(refresh.highlightChanges({ includePrevious: true, timeoutMs: 60_000 })).toBe(2);
      expect(thumb(2)?.getAttribute(MARK)).toBe('shown');
      expect(thumb(5)?.getAttribute(MARK)).toBe('entering');
      expect(host.querySelector('.pptx-surface__refresh')?.getAttribute(MARK)).toBe('shown');
      expect(ring()?.hasAttribute(MARK)).toBe(false);
      await tick(40);
      expect(thumb(5)?.getAttribute(MARK)).toBe('shown');
      done();
    });
  }

  test('a flash follows its slide when a slide comes before it during the commit', async () => {
    const { editor, rail, refresh, apply, done } = await results('reload');
    await apply({ text: [4] }, [4]);
    expect(refresh.highlightChanges({ animation: false, timeoutMs: 60_000 })).toBe(1);
    const marked = () =>
      [...rail.querySelectorAll(`.pptx-rail__thumb[${MARK}]`)].map((el) =>
        el.parentElement!.getAttribute('data-index')
      );
    expect(marked()).toEqual(['4']);
    // The rail draws again on each change; during the commit the flash follows the part.
    const seen: (string | null)[][] = [];
    const off = editor.subscribe(() => seen.push(marked()));
    const r = await apply({ text: [4], insertAfter: 0 }, [5]);
    off();
    expect(r.changes[0]).toMatchObject({ id: 's5', isNew: true });
    for (const list of seen) expect(list).not.toContain('4');
    expect(marked()).toEqual(['5']);
    done();
  });
});

/** A deck with one embedded font part whose bytes are `font`. */
function withFont(bytes: Uint8Array, font: Uint8Array<ArrayBuffer>): Uint8Array<ArrayBuffer> {
  const files = unzipSync(bytes);
  const append = (part: string, end: string, xml: string) => {
    files[part] = strToU8(strFromU8(files[part]!).replace(end, xml + end));
  };
  append(
    'ppt/presentation.xml',
    '</p:presentation>',
    '<p:embeddedFontLst><p:embeddedFont><p:font typeface="Test Sans"/>' +
      '<p:regular r:id="rIdTestFont"/></p:embeddedFont></p:embeddedFontLst>'
  );
  append(
    'ppt/_rels/presentation.xml.rels',
    '</Relationships>',
    '<Relationship Id="rIdTestFont" ' +
      'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/font" ' +
      'Target="fonts/test.fntdata"/>'
  );
  append(
    '[Content_Types].xml',
    '</Types>',
    '<Override PartName="/ppt/fonts/test.fntdata" ContentType="application/x-fontdata"/>'
  );
  files['ppt/fonts/test.fntdata'] = font;
  return zipSync(files) as Uint8Array<ArrayBuffer>;
}

describe('document refresh regressions: replacement paths', () => {
  test('edit mode refuses changed embedded fonts; a reload opens the file whole', async () => {
    const { editor, done } = await mount(deck('A'));
    const refresh = createDocumentRefresh(editor);
    const submission = await refresh.capture();
    const fonts = withFont(deck('A'), new Uint8Array([1, 2, 3, 4]));
    const edit = await refresh.applyUpdate({ submission, sequence: 1, bytes: fonts, mode: 'edit' });
    expect(edit.ok ? 'ok' : [edit.code, edit.reason]).toEqual([
      'edit-unsupported',
      'unsupported-difference',
    ]);
    const direct = editor.exec({ type: 'replaceDocument', pptx: fonts });
    expect(direct.replaceDocument).toMatchObject({
      ok: false,
      code: 'editUnsupported',
      reason: 'unsupportedDifference',
    });
    const reload = await refresh.applyUpdate({ submission, sequence: 2, bytes: fonts });
    expect(reload.ok ? reload.mode : reload.code).toBe('reload');
    // Other font bytes are a font change too.
    const second = withFont(deck('A'), new Uint8Array([5, 6, 7, 8]));
    const again = await refresh.applyUpdate({
      submission,
      sequence: 3,
      bytes: second,
      mode: 'edit',
    });
    expect(again.ok ? 'ok' : again.code).toBe('edit-unsupported');
    done();
  });

  test('capture() during the quiet load of a reload rejects with busy', async () => {
    const { editor, done } = await mount(deck('A'));
    const refresh = createDocumentRefresh(editor);
    const submission = await refresh.capture();
    const load = editor.load.bind(editor);
    let during: Promise<unknown> = Promise.resolve();
    editor.load = (bytes, options) => {
      if (options?.quiet)
        during = refresh.capture().then(
          () => 'captured',
          (e: { code?: string }) => e.code
        );
      return load(bytes, options);
    };
    // Duplicate shape ids: only a load accepts the file.
    const bytes = buildPptx({ slides: [{ shapes: box(2, 'Z') + box(2, 'Y') }] });
    const r = await refresh.applyUpdate({ submission, sequence: 1, bytes });
    editor.load = load;
    expect(r.ok).toBe(true);
    expect(await during).toBe('busy');
    done();
  });

  test('a quiet load that fails names no cause for the next load', async () => {
    const { editor, done } = await mount(deck('A'));
    const refresh = createDocumentRefresh(editor);
    const submission = await refresh.capture();
    const load = editor.load.bind(editor);
    editor.load = async (bytes, options) =>
      options?.quiet ? { ok: false, reason: 'errors.unsupported' } : load(bytes, options);
    const bytes = buildPptx({ slides: [{ shapes: box(2, 'Z') + box(2, 'Y') }] });
    const r = await refresh.applyUpdate({ submission, sequence: 1, bytes });
    editor.load = load;
    expect(r.ok ? 'ok' : r.code).toBe('load-failed');
    const seen: unknown[] = [];
    const off = onEditorEvent(editor, 'change', (_s, c) => seen.push(c.source));
    expect((await editor.load(deck('Host'))).ok).toBe(true);
    off();
    expect(seen).toContain('load');
    expect(seen).not.toContain('refresh');
    done();
  });
});

describe('document refresh regressions: the exit of a flash', () => {
  /** An editor with one accepted result that changes shape 2 of slide 1. */
  async function flashed(edit = deck('B')) {
    const m = await mount(deck('A'));
    const refresh = createDocumentRefresh(m.editor);
    const submission = await refresh.capture();
    const r = await refresh.applyUpdate({
      submission,
      sequence: 1,
      mode: 'edit',
      bytes: edit,
      changes: [{ id: 'c1', location: { slideIndex: 0, shapeId: 2 } }],
    });
    if (!r.ok) throw new Error(r.code);
    m.drawRail();
    const thumb = () => m.rail.querySelector<HTMLElement>('[data-index="0"] .pptx-rail__thumb')!;
    const layer = () => m.host.querySelector<SVGSVGElement>('.pptx-surface__refresh');
    return { ...m, refresh, submission, thumb, layer };
  }

  test('clearHighlights() writes its exit duration and easing to every marked element', async () => {
    const { refresh, thumb, layer, done } = await flashed();
    expect(refresh.highlightChanges({ timeoutMs: 60_000 })).toBe(1);
    await tick(40);
    expect(thumb().style.getPropertyValue('--pptx-refresh-exit')).toBe('180ms');
    expect(layer()!.style.getPropertyValue('--pptx-refresh-exit-easing')).toBe(
      'var(--doc-motion-ease-in-out, ease-in-out)'
    );
    refresh.clearHighlights({ animation: { exitDurationMs: 600, easing: 'linear' } });
    for (const el of [thumb(), layer()!]) {
      expect(el.getAttribute(MARK)).toBe('exiting');
      expect(el.style.getPropertyValue('--pptx-refresh-exit')).toBe('600ms');
      expect(el.style.getPropertyValue('--pptx-refresh-exit-easing')).toBe('linear');
    }
    // The removal waits for the same duration.
    await tick(300);
    expect(thumb().getAttribute(MARK)).toBe('exiting');
    await tick(400);
    expect(thumb().hasAttribute(MARK)).toBe(false);
    expect(layer()).toBeNull();
    done();
  });

  test('a flash request during the exit turns kept places back to shown, with no entrance', async () => {
    const { refresh, thumb, layer, done } = await flashed();
    expect(refresh.highlightChanges({ timeoutMs: 60_000 })).toBe(1);
    await tick(40);
    refresh.clearHighlights({ animation: { exitDurationMs: 200 } });
    expect(layer()!.getAttribute(MARK)).toBe('exiting');
    expect(refresh.highlightChanges({ timeoutMs: 60_000 })).toBe(1);
    // No mark starts from the entering state: the layer never drops to its start opacity.
    expect(thumb().getAttribute(MARK)).toBe('shown');
    expect(layer()!.getAttribute(MARK)).toBe('shown');
    expect(layer()!.querySelector(`[${MARK}]`)).toBeNull();
    expect(refresh.snapshot().highlightsVisible).toBe(true);
    // The exit timer of the dismissal does not remove the new flash.
    await tick(300);
    expect(thumb().getAttribute(MARK)).toBe('shown');
    expect(layer()?.firstChild).not.toBeNull();
    // To replay the entrance, remove the flash at once first.
    refresh.clearHighlights({ animation: false });
    expect(refresh.highlightChanges({ timeoutMs: 60_000 })).toBe(1);
    expect(thumb().getAttribute(MARK)).toBe('entering');
    done();
  });

  test('a held canvas mark goes when the commit replaces its part whole', async () => {
    const { editor, refresh, submission, layer, done } = await flashed();
    expect(refresh.highlightChanges({ animation: false, timeoutMs: 60_000 })).toBe(1);
    expect(layer()?.firstChild).not.toBeNull();
    // A new namespace declaration on the slide root: the difference replaces the part whole,
    // and the parse gives its nodes new ids.
    const files = unzipSync(deck('C'));
    const name = 'ppt/slides/slide1.xml';
    files[name] = strToU8(
      strFromU8(files[name]!).replace('<p:sld ', '<p:sld xmlns:x14="urn:example:x14" ')
    );
    const seen: boolean[] = [];
    const off = editor.subscribe(() => seen.push(!!layer()?.firstChild));
    const r = await refresh.applyUpdate({
      submission,
      sequence: 2,
      mode: 'edit',
      bytes: zipSync(files),
      changes: [{ id: 'c1', location: { slideIndex: 0, shapeId: 2 } }],
    });
    off();
    expect(r.ok).toBe(true);
    // From the commit on, no frame draws a canvas mark at a node id of the replaced part.
    expect(seen.length).toBeGreaterThan(0);
    expect(seen).not.toContain(true);
    done();
  });
});
