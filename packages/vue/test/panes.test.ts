// The Vue panes (the slide rail, the slide sorter, the speaker notes, the zoom pill, the side
// panel, and the status bar's Notes and view parts) render the same markup, ARIA, and keys as
// their React twins (packages/react/src/panes.tsx, RailSections.tsx, SlideSorter.tsx,
// NotesPane.tsx, StatusBar.tsx). Each case renders both adapters on one deck and compares a
// canonical form of the DOM, then drives the Vue part with clicks and keys.

import { afterEach, describe, expect, spyOn, test } from 'bun:test';
import { createElement, type FunctionComponent, type ReactElement } from 'react';
import { act, cleanup, render } from '@testing-library/react';
import {
  createApp,
  defineComponent,
  h,
  nextTick,
  ref,
  shallowRef,
  type App,
  type Component,
} from 'vue';
import {
  createPptxEditor,
  MOUNT_FIRST_ITEMS,
  resetDevWarnings,
  type PptxEditorInstance,
} from '@pptx-editor.dev/core/editor';
import { buildPptx, para, run, shape } from '../../../scripts/fixtures/builder.ts';
import { buildSlideLayoutsDeck } from '../../../scripts/fixtures/slide-layouts.ts';
import { textDeck } from '../../../scripts/bench/text-deck.ts';
import { PptxEditorContext } from '../../react/src/context.ts';
import {
  NotesPane as RNotesPane,
  SidePanel,
  SlideRail as RSlideRail,
  ZoomPill,
  type NotesPaneProps as RNotesPaneProps,
  type SlideRailProps as RSlideRailProps,
} from '../../react/src/panes.tsx';
import { SlideSorter, ViewSwitch as RViewSwitch } from '../../react/src/SlideSorter.tsx';
import {
  PptxEditorNotesPane,
  PptxEditorRoot,
  PptxEditorSidePanel,
  PptxEditorSlideRail,
  PptxEditorSlideSorter,
  PptxEditorStatusBar,
  PptxEditorZoomPill,
  requestSectionRename,
  useNotesPane,
  type UseNotesPaneReturn,
} from '../src/index.ts';
import { providePptxEditor } from '../src/composables.ts';
import { ViewSwitch } from '../src/SlideSorter.tsx';
import { canonicalDom } from './canonical-dom.ts';

const IN = 914_400;

// The React panes take optional props (`props = {}`), which `createElement` cannot overload on.
const SlideRail = RSlideRail as FunctionComponent<RSlideRailProps>;
const NotesPane = RNotesPane as FunctionComponent<RNotesPaneProps>;

let app: App | null = null;
let host: HTMLElement | null = null;
afterEach(() => {
  cleanup();
  app?.unmount();
  host?.remove();
  app = null;
  host = null;
  document.body.replaceChildren();
});

async function mountVue(
  renderFn: () => unknown,
  editor?: PptxEditorInstance | null
): Promise<HTMLElement> {
  app?.unmount();
  host?.remove();
  host = document.createElement('div');
  document.body.append(host);
  app = createApp(
    defineComponent({
      setup() {
        if (editor !== undefined) providePptxEditor(shallowRef(editor));
        return renderFn;
      },
    }) as Component
  );
  app.mount(host);
  await nextTick();
  return host;
}

function mountReact(element: ReactElement, editor: PptxEditorInstance): HTMLElement {
  return render(createElement(PptxEditorContext.Provider, { value: editor }, element)).container;
}

/**
 * The canonical form of a pane host, with no slide renders (the engine paints the same thumbnail
 * in both; its SVG ids are per render).
 */
const canon = (root: Element): string =>
  canonicalDom(root, {
    normalizeStyle: true,
    skipChildren: (el) => el.classList.contains('pptx-rail__thumb'),
  });

async function same(react: ReactElement, vue: () => unknown, editor: PptxEditorInstance) {
  const r = canon(mountReact(react, editor));
  cleanup();
  const v = canon(await mountVue(vue, editor));
  expect(v).toEqual(r);
  return v;
}

async function sectionsDeck(): Promise<PptxEditorInstance> {
  const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await editor.load(buildSlideLayoutsDeck())).ok).toBe(true);
  return editor;
}

async function flatDeck(count = 3): Promise<PptxEditorInstance> {
  const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  const bytes = buildPptx({
    slides: Array.from({ length: count }, (_, i) => ({
      layout: 'blank' as const,
      shapes: shape({ id: 2, x: IN, y: IN, cx: 2 * IN, cy: IN, text: para(run(`S${i + 1}`)) }),
    })),
  });
  expect((await editor.load(bytes)).ok).toBe(true);
  return editor;
}

const key = (el: Element, k: string, init: KeyboardEventInit = {}): void => {
  el.dispatchEvent(new KeyboardEvent('keydown', { key: k, bubbles: true, ...init }));
};

const railItems = (root: ParentNode): HTMLElement[] => [
  ...root.querySelectorAll<HTMLElement>('.pptx-rail__item'),
];

describe('Vue panes match React', () => {
  test('the flat slide rail', async () => {
    const editor = await flatDeck();
    editor.selectSlide(1);
    await same(createElement(SlideRail), () => h(PptxEditorSlideRail), editor);
  });

  test('the slide rail with sections, one collapsed', async () => {
    const editor = await sectionsDeck();
    const details = editor.snapshot().sections![1]!.id;
    editor.exec({ type: 'setSectionsCollapsed', collapsed: true, sectionIds: [details] });
    await same(
      createElement(SlideRail, { className: 'host', paneWidth: 200 }),
      () => h(PptxEditorSlideRail, { class: 'host', paneWidth: 200 }),
      editor
    );
  });

  test('the closed rail strip, and no rail without its toggle', async () => {
    const editor = await flatDeck();
    await same(
      createElement(SlideRail, { defaultOpen: false, style: { width: 300, color: 'red' } }),
      () => h(PptxEditorSlideRail, { defaultOpen: false, style: { width: '300px', color: 'red' } }),
      editor
    );
    await same(
      createElement(SlideRail, { open: false, toggle: false }),
      () => h(PptxEditorSlideRail, { open: false, toggle: false }),
      editor
    );
  });

  test('the horizontal slide rail, flat, with sections, and closed', async () => {
    const flat = await flatDeck();
    flat.selectSlide(1);
    await same(
      createElement(SlideRail, { orientation: 'horizontal' }),
      () => h(PptxEditorSlideRail, { orientation: 'horizontal' }),
      flat
    );
    const grouped = await sectionsDeck();
    await same(
      createElement(SlideRail, { orientation: 'horizontal' }),
      () => h(PptxEditorSlideRail, { orientation: 'horizontal' }),
      grouped
    );
    await same(
      createElement(SlideRail, { orientation: 'horizontal', defaultOpen: false }),
      () => h(PptxEditorSlideRail, { orientation: 'horizontal', defaultOpen: false }),
      flat
    );
  });

  test('the slide sorter, flat and with sections', async () => {
    const flat = await flatDeck();
    await same(createElement(SlideSorter), () => h(PptxEditorSlideSorter), flat);
    const grouped = await sectionsDeck();
    grouped.goToSlide(4);
    await same(createElement(SlideSorter), () => h(PptxEditorSlideSorter), grouped);
  });

  test('the notes pane, open and closed', async () => {
    const editor = await flatDeck();
    editor.exec({ type: 'setNotes', index: 0, text: 'Say hello\nThen bye' });
    await same(createElement(NotesPane), () => h(PptxEditorNotesPane), editor);
    await same(
      createElement(NotesPane, { defaultOpen: false, showFormatButtons: false }),
      () => h(PptxEditorNotesPane, { defaultOpen: false, showFormatButtons: false }),
      editor
    );
  });

  test('the zoom pill, the view switch, and the side panel', async () => {
    const editor = await flatDeck();
    await same(createElement(ZoomPill), () => h(PptxEditorZoomPill), editor);
    await same(createElement(RViewSwitch), () => h(ViewSwitch), editor);
    await same(
      createElement(SidePanel, { onClose: () => {}, paneWidth: 280 }),
      () => h(PptxEditorSidePanel, { onClose: () => {}, paneWidth: 280 }),
      editor
    );
    // A selected shape fills the Format tab.
    const id = editor.session!.layout(0)!.items[0]!.element.nodeId;
    editor.setSelection({ kind: 'shapes', nodeIds: [id] });
    await same(createElement(SidePanel), () => h(PptxEditorSidePanel), editor);
  });

  test('the side panel tabs are one Tab stop, and the arrow keys move between them', async () => {
    const editor = await flatDeck();
    for (const mount of [
      async () => mountReact(createElement(SidePanel), editor),
      () => mountVue(() => h(PptxEditorSidePanel), editor),
    ]) {
      const el = await mount();
      const tabs = () => [...el.querySelectorAll<HTMLElement>('[role="tab"]')];
      expect(tabs().map((t) => t.tabIndex)).toEqual([0, -1]);
      tabs()[0]!.focus();
      key(tabs()[0]!, 'ArrowRight');
      await nextTick();
      await act(async () => undefined);
      expect(tabs().map((t) => t.getAttribute('aria-selected'))).toEqual(['false', 'true']);
      expect(tabs().map((t) => t.tabIndex)).toEqual([-1, 0]);
      expect(document.activeElement).toBe(tabs()[1]!);
      cleanup();
    }
  });

  test('a sorter view turns the zoom pill off with the reason', async () => {
    const editor = await flatDeck();
    editor.exec({ type: 'setPresentationView', view: 'slideSorter' });
    await same(createElement(ZoomPill), () => h(PptxEditorZoomPill), editor);
    // The sorter draws no slide at the zoom: the level names no percentage there.
    const level = document.querySelector<HTMLElement>('.pptx-zoom__level')!;
    expect(level.textContent!.includes('%')).toBe(false);
    expect(level.getAttribute('aria-label')).toBe('Zoom');
    expect(level.getAttribute('aria-disabled')).toBe('true');
  });
});

describe('the Vue slide rail', () => {
  test('arrow keys move the current slide; Delete and Ctrl+D act on the selection', async () => {
    const editor = await flatDeck();
    const el = await mountVue(() => h(PptxEditorSlideRail), editor);
    key(railItems(el)[0]!, 'ArrowDown');
    expect(editor.snapshot().currentSlide).toBe(1);
    await nextTick();
    expect(railItems(el)[1]!.getAttribute('aria-current')).toBe('true');
    key(railItems(el)[1]!, 'd', { ctrlKey: true });
    expect(editor.snapshot().slides.length).toBe(4);
    await nextTick();
    key(railItems(el)[2]!, 'Delete');
    expect(editor.snapshot().slides.length).toBe(3);
  });

  test('a horizontal rail steps with Left and Right, and Ctrl+Right moves a slide', async () => {
    const editor = await flatDeck();
    const first = editor.snapshot().slides[0]!.id;
    const el = await mountVue(() => h(PptxEditorSlideRail, { orientation: 'horizontal' }), editor);
    key(railItems(el)[0]!, 'ArrowDown');
    expect(editor.snapshot().currentSlide).toBe(0);
    key(railItems(el)[0]!, 'ArrowRight');
    expect(editor.snapshot().currentSlide).toBe(1);
    await nextTick();
    key(railItems(el)[1]!, 'ArrowLeft');
    expect(editor.snapshot().currentSlide).toBe(0);
    await nextTick();
    key(railItems(el)[0]!, 'ArrowRight', { ctrlKey: true });
    expect(editor.snapshot().slides[1]!.id).toBe(first);
  });

  test('a drag moves a slide as one moveSlides step', async () => {
    const editor = await flatDeck();
    const first = editor.snapshot().slides[0]!.id;
    const el = await mountVue(() => h(PptxEditorSlideRail), editor);
    railItems(el)[0]!.dispatchEvent(new Event('dragstart', { bubbles: true }));
    railItems(el)[2]!.dispatchEvent(new Event('dragover', { bubbles: true, cancelable: true }));
    await nextTick();
    // The pointer is past the middle of the last slide: the marker says after.
    expect(railItems(el)[2]!.dataset.drop).toBe('after');
    railItems(el)[2]!.dispatchEvent(new Event('drop', { bubbles: true, cancelable: true }));
    expect(
      editor
        .snapshot()
        .slides.map((s) => s.id)
        .indexOf(first)
    ).toBe(2);
    editor.exec({ type: 'undo' });
    expect(editor.snapshot().slides[0]!.id).toBe(first);
  });

  test('a click selects; Shift+click a range; the buttons act on the selection', async () => {
    const editor = await flatDeck();
    const el = await mountVue(() => h(PptxEditorSlideRail), editor);
    railItems(el)[0]!.click();
    railItems(el)[2]!.dispatchEvent(new MouseEvent('click', { bubbles: true, shiftKey: true }));
    expect(editor.snapshot().selectedSlides).toEqual([0, 1, 2]);
    await nextTick();
    const del = el.querySelector<HTMLElement>('[aria-label="Delete slide"]')!;
    expect(del.getAttribute('aria-disabled')).toBe('true');
    railItems(el)[1]!.click();
    await nextTick();
    el.querySelector<HTMLElement>('[aria-label="Move slide down"]')!.click();
    expect(editor.snapshot().currentSlide).toBe(2);
  });

  test('the toggle hides the rail to a strip and emits update:open; a host can control it', async () => {
    const editor = await flatDeck();
    const seen: boolean[] = [];
    const el = await mountVue(
      () => h(PptxEditorSlideRail, { 'onUpdate:open': (v: boolean) => seen.push(v) }),
      editor
    );
    el.querySelector<HTMLElement>('.pptx-rail__collapse')!.click();
    await nextTick();
    expect(seen).toEqual([false]);
    expect(el.querySelector('nav')!.dataset.open).toBe('false');
    // Controlled: a refused change keeps the rail open.
    const open = ref(true);
    const host2 = await mountVue(
      () => h(PptxEditorSlideRail, { open: open.value, 'onUpdate:open': () => {} }),
      editor
    );
    host2.querySelector<HTMLElement>('.pptx-rail__collapse')!.click();
    await nextTick();
    expect(host2.querySelector('nav')!.dataset.open).toBe('true');
    open.value = false;
    await nextTick();
    expect(host2.querySelector('nav')!.dataset.open).toBe('false');
  });

  test('sections collapse, rename in place on request, and Enter writes the name', async () => {
    const editor = await sectionsDeck();
    const el = await mountVue(() => h(PptxEditorSlideRail), editor);
    const details = editor.snapshot().sections![1]!;
    const header = (): HTMLElement =>
      el.querySelector<HTMLElement>(`[data-section="${details.id}"]`)!;
    header().querySelector<HTMLElement>('.pptx-rail__section-toggle')!.click();
    await nextTick();
    expect(header().getAttribute('aria-expanded')).toBe('false');
    expect(editor.snapshot().collapsedSections).toEqual([details.id]);
    // Right on a shut header opens it.
    key(header(), 'ArrowRight');
    await nextTick();
    expect(header().getAttribute('aria-expanded')).toBe('true');
    requestSectionRename(editor, details.id);
    await nextTick();
    await nextTick();
    const field = el.querySelector<HTMLInputElement>('.pptx-rail__section-name-field')!;
    expect(field !== null).toBe(true);
    field.value = 'Middle';
    field.dispatchEvent(new Event('input', { bubbles: true }));
    key(field, 'Enter');
    expect(editor.snapshot().sections![1]!.name).toBe('Middle');
  });

  test('a click on a section name selects its slides', async () => {
    const editor = await sectionsDeck();
    const el = await mountVue(() => h(PptxEditorSlideRail), editor);
    const details = editor.snapshot().sections![1]!;
    el.querySelector<HTMLElement>(
      `[data-section="${details.id}"] .pptx-rail__section-name`
    )!.click();
    const { selectedSlides } = editor.snapshot();
    expect(selectedSlides.length).toBe(details.slideCount);
    expect(selectedSlides[0]).toBe(details.firstSlide);
    // The first slide of the section shows, not the last.
    expect(editor.snapshot().currentSlide).toBe(details.firstSlide);
  });
});

describe('the Vue slide sorter', () => {
  test('Enter opens the current slide in the Normal view; arrows select', async () => {
    const editor = await flatDeck(4);
    editor.exec({ type: 'setPresentationView', view: 'slideSorter' });
    const el = await mountVue(() => h(PptxEditorSlideSorter), editor);
    const tiles = (): HTMLElement[] => [...el.querySelectorAll<HTMLElement>('.pptx-sorter__item')];
    key(tiles()[0]!, 'ArrowRight');
    expect(editor.snapshot().currentSlide).toBe(1);
    await nextTick();
    key(tiles()[1]!, 'Enter');
    expect(editor.snapshot().presentationView).toBe('normal');
  });

  test('shown from a menu, it focuses the current slide; focus a host put elsewhere stays', async () => {
    const editor = await flatDeck(4);
    editor.goToSlide(2);
    (document.activeElement as HTMLElement | null)?.blur();
    editor.exec({ type: 'setPresentationView', view: 'slideSorter' });
    const el = await mountVue(() => h(PptxEditorSlideSorter), editor);
    await new Promise((r) => setTimeout(r, 50));
    const tiles = [...el.querySelectorAll<HTMLElement>('.pptx-sorter__item')];
    expect(document.activeElement).toBe(tiles[2]!);
    // A host field keeps focus when the sorter appears.
    const host = document.createElement('input');
    document.body.append(host);
    host.focus();
    await mountVue(() => h(PptxEditorSlideSorter), editor);
    await new Promise((r) => setTimeout(r, 50));
    expect(document.activeElement).toBe(host);
    host.remove();
  });

  test('a double-click opens a slide', async () => {
    const editor = await flatDeck(4);
    editor.exec({ type: 'setPresentationView', view: 'slideSorter' });
    const el = await mountVue(() => h(PptxEditorSlideSorter), editor);
    el.querySelectorAll<HTMLElement>('.pptx-sorter__item')[2]!.dispatchEvent(
      new MouseEvent('dblclick', { bubbles: true })
    );
    expect(editor.snapshot().currentSlide).toBe(2);
    expect(editor.snapshot().presentationView).toBe('normal');
  });
});

describe('the Vue notes pane', () => {
  test('typed text commits as one setNotes on blur', async () => {
    const editor = await flatDeck();
    const el = await mountVue(() => h(PptxEditorNotesPane), editor);
    const field = el.querySelector<HTMLElement>('.pptx-notes__text')!;
    field.focus();
    field.textContent = 'Hello there';
    field.dispatchEvent(new Event('input', { bubbles: true }));
    field.blur();
    await nextTick();
    expect(editor.snapshot().notes).toBe('Hello there');
    expect(editor.snapshot().canUndo).toBe(true);
  });

  test('the collapse button closes the field and emits update:open', async () => {
    const editor = await flatDeck();
    const seen: boolean[] = [];
    const el = await mountVue(
      () => h(PptxEditorNotesPane, { 'onUpdate:open': (v: boolean) => seen.push(v) }),
      editor
    );
    el.querySelector<HTMLElement>('[aria-label="Hide speaker notes"]')!.click();
    await nextTick();
    expect(seen).toEqual([false]);
    expect(el.querySelector('.pptx-notes__text') === null).toBe(true);
    expect(el.querySelector('section')!.dataset.open).toBe('false');
  });

  test('the status bar Notes button and useNotesPane share the open state with the pane', async () => {
    let pane: UseNotesPaneReturn | null = null;
    const Probe = defineComponent({
      setup() {
        pane = useNotesPane();
        return () => null;
      },
    });
    const el = await mountVue(() =>
      h(PptxEditorRoot, null, () => [h(PptxEditorNotesPane), h(PptxEditorStatusBar), h(Probe)])
    );
    const button = el.querySelector<HTMLElement>('.pptx-statusbar__notes')!;
    expect(button.getAttribute('aria-pressed')).toBe('true');
    button.click();
    await nextTick();
    expect(el.querySelector('.pptx-notes')!.getAttribute('data-open')).toBe('false');
    expect(button.getAttribute('aria-pressed')).toBe('false');
    pane!.toggle();
    await nextTick();
    expect(el.querySelector('.pptx-notes')!.getAttribute('data-open')).toBe('true');
    // The bar has the slide number, the Notes button, the view switch, and the zoom pill.
    expect(
      [...el.querySelectorAll<HTMLElement>('.pptx-statusbar--bar [data-part]')].map(
        (p) => p.dataset.part
      )
    ).toEqual(['slideNumber', 'notes', 'views', 'zoom']);
  });

  test('a second notes pane in one root warns', async () => {
    resetDevWarnings();
    const warn = spyOn(console, 'warn').mockImplementation(() => {});
    try {
      await mountVue(() =>
        h(PptxEditorRoot, null, () => [h(PptxEditorNotesPane), h(PptxEditorNotesPane)])
      );
      expect(warn.mock.calls.some((c) => String(c[0]).includes('PptxEditorNotesPane'))).toBe(true);
    } finally {
      warn.mockRestore();
    }
  });
});

describe('the Vue zoom pill', () => {
  test('the level opens the zoom menu; a level sets the zoom', async () => {
    const editor = await flatDeck();
    const el = await mountVue(() => h(PptxEditorZoomPill), editor);
    el.querySelector<HTMLElement>('.pptx-zoom__level')!.click();
    await nextTick();
    const rows = [...document.querySelectorAll<HTMLElement>('[role="menuitemradio"]')];
    expect(rows.map((r) => r.textContent)).toEqual([
      'Fit to window',
      '50%',
      '75%',
      '100%',
      '150%',
      '200%',
    ]);
    // The menu opens on the current level, the checked row.
    const checked = rows.find((r) => r.getAttribute('aria-checked') === 'true');
    expect(checked !== undefined && document.activeElement === checked).toBe(true);
    rows[1]!.click();
    expect(editor.snapshot().zoom).toBe(0.5);
    await nextTick();
    expect(document.querySelector('[role="menuitemradio"]') === null).toBe(true);
  });

  test('React and Vue agree after the same zoom step', async () => {
    const editor = await flatDeck();
    act(() => editor.setZoom(1.5));
    await same(createElement(ZoomPill), () => h(PptxEditorZoomPill), editor);
  });
});

describe('the rails of a long deck that opens progressively', () => {
  const wait = (ms: number): Promise<void> => new Promise((r) => setTimeout(r, ms));

  test('React and Vue mount the same items, whole sections, and the same placeholders', async () => {
    const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    expect((await editor.load(textDeck(300))).ok).toBe(true);
    // Sections of 30 and 15 slides: item 40 falls in the second, so it mounts whole.
    expect(editor.exec({ type: 'addSection', index: 30 }).ok).toBe(true);
    expect(editor.exec({ type: 'addSection', index: 45 }).ok).toBe(true);
    editor.goToSlide(0);
    expect(editor.snapshot().pendingSlides ?? 0).toBeGreaterThan(0);
    // Both rails mount in the same task, while the deck still opens.
    const react = mountReact(createElement(SlideRail), editor);
    const vue = await mountVue(() => h(PptxEditorSlideRail), editor);
    const shape = (root: ParentNode) => ({
      items: railItems(root).length,
      headers: root.querySelectorAll('.pptx-rail__section').length,
      pending: root.querySelectorAll('.pptx-rail__thumb[data-pending]').length,
    });
    const first = shape(react);
    expect(first.items).toBe(45);
    expect(first.items).toBeGreaterThan(MOUNT_FIRST_ITEMS);
    expect(first.headers).toBe(2);
    expect(first.pending).toBeGreaterThan(0);
    expect(shape(vue)).toEqual(first);
    // When the open settles and the frames paint, both mount every item, none pending.
    for (let i = 0; i < 500 && (shape(react).items < 300 || shape(vue).items < 300); i += 1)
      await act(async () => {
        await wait(5);
        await nextTick();
      });
    for (let i = 0; i < 500 && (editor.snapshot().pendingSlides ?? 0) > 0; i += 1)
      await act(() => wait(5));
    await nextTick();
    const last = shape(react);
    expect(last).toEqual({ items: 300, headers: 3, pending: 0 });
    expect(shape(vue)).toEqual(last);
    editor.destroy();
  });
});
