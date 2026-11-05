// The Vue overlays against their React twins: the find bar, the edit and font notices, the link
// popup, and the slide show render the same markup for one editor, and behave the same (keys,
// requests through the root's action host, focus). After DOCX Editor's cross-adapter
// differential test (docx-editor/packages/vue/test/cross-adapter-differential.test.ts).

import { afterEach, describe, expect, jest, test } from 'bun:test';
import { createElement, type ReactNode } from 'react';
import { act, cleanup, render } from '@testing-library/react';
import {
  createApp,
  defineComponent,
  h,
  nextTick,
  provide,
  shallowRef,
  type App,
  type Component,
} from 'vue';
import {
  createPptxEditor,
  requestLinkPopup,
  runClipboardSlot,
  startSlideshowController,
  type PptxEditorInstance,
} from '@pptx-editor.dev/core/editor';
import { buildPptx, para, run, shape } from '../../../scripts/fixtures/builder.ts';
import { PptxEditorContext, SlideshowContext } from '../../react/src/context.ts';
import { fontBytes } from '../../core/src/editor/__tests__/font-test-helpers.ts';
import { FindBar as ReactFindBar } from '../../react/src/FindBar.tsx';
import { EditNotice as ReactEditNotice } from '../../react/src/EditNotice.tsx';
import { FontNotice as ReactFontNotice } from '../../react/src/FontNotice.tsx';
import { HyperlinkPopup as ReactHyperlinkPopup } from '../../react/src/HyperlinkPopup.tsx';
import { Slideshow as ReactSlideshow } from '../../react/src/Slideshow.tsx';
import {
  HyperlinkPopup,
  PptxEditorContent,
  PptxEditorEditNotice,
  PptxEditorFindBar,
  PptxEditorFontNotice,
  PptxEditorRoot,
  PptxEditorSlideshow,
  PptxEditorViewport,
  pptxSlideshowKey,
  rejectionMessage,
  requestFindBar,
  useEditorCommand,
  useSlideshow,
  useTranslation,
  type PptxEditorPopups,
  type UseEditorCommandReturn,
  type UseSlideshowReturn,
} from '../src/index.ts';
import { providePptxEditor } from '../src/composables.ts';
import { canonicalDom } from './canonical-dom.ts';

const IN = 914_400;

let app: App | null = null;
let vueHost: HTMLElement | null = null;
afterEach(() => {
  cleanup();
  app?.unmount();
  vueHost?.remove();
  app = null;
  vueHost = null;
});

function deck(): Uint8Array {
  return buildPptx({
    slides: [
      {
        layout: 'blank',
        shapes: shape({
          id: 2,
          name: 'T',
          x: IN,
          y: IN,
          cx: 6 * IN,
          cy: IN,
          text: para(run('one two one')),
        }),
      },
      {
        layout: 'blank',
        shapes: shape({ id: 2, x: IN, y: IN, cx: IN, cy: IN, text: para(run('one')) }),
      },
    ],
  });
}

async function open(): Promise<PptxEditorInstance> {
  const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await editor.load(deck())).ok).toBe(true);
  return editor;
}

/** Mount a Vue tree that provides `editor`, inside a `.pptx-editor` scope. */
async function mountVue(editor: PptxEditorInstance, render: () => unknown, setup?: () => void) {
  vueHost = document.createElement('div');
  vueHost.className = 'pptx-editor';
  document.body.append(vueHost);
  app = createApp(
    defineComponent({
      setup() {
        providePptxEditor(shallowRef(editor));
        setup?.();
        return render;
      },
    }) as Component
  );
  app.mount(vueHost);
  await nextTick();
  return vueHost;
}

function mountReact(editor: PptxEditorInstance, node: ReactNode): HTMLElement {
  const { container } = render(
    createElement(
      PptxEditorContext.Provider,
      { value: editor },
      createElement('div', { className: 'pptx-editor' }, node)
    )
  );
  return container.firstElementChild as HTMLElement;
}

/**
 * The canonical form of an overlay element. React writes an empty `class` on an SVG that Vue
 * leaves out, and the reverse.
 */
const canonical = (el: Element | null, skipChildren?: (el: Element) => boolean): string =>
  canonicalDom(el, { self: true, dropEmptyClass: true, skipChildren });

describe('PptxEditorFindBar', () => {
  test('renders the same bar as React when opened to replace, with the counter', async () => {
    const editor = await open();
    const reactScope = mountReact(editor, createElement(ReactFindBar, { className: 'acme' }));
    const vueScope = await mountVue(editor, () => h(PptxEditorFindBar, { class: 'acme' }));
    // Closed: both render the hidden search region only.
    expect(canonical(vueScope.querySelector('.pptx-find'))).toBe(
      canonical(reactScope.querySelector('.pptx-find'))
    );
    // One request names the editor: the React and the Vue bar of that editor both open.
    act(() => requestFindBar(editor, true));
    await nextTick();
    const type = (scope: Element, value: string) => {
      const field = scope.querySelector<HTMLInputElement>('input[aria-label="Find"]')!;
      act(() => {
        // React reads the value through its own setter tracker; set it the way it listens.
        const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')!.set!;
        setter.call(field, value);
        field.dispatchEvent(new Event('input', { bubbles: true }));
      });
    };
    type(reactScope, 'one');
    type(vueScope, 'one');
    await nextTick();
    const vueBar = vueScope.querySelector('.pptx-find');
    expect(vueBar?.classList.contains('acme')).toBe(true);
    // Find as you type: the first match is selected before Enter, in both adapters.
    expect(vueScope.querySelector('.pptx-find__count')?.textContent).toBe('1 of 3');
    expect(canonical(vueBar)).toBe(canonical(reactScope.querySelector('.pptx-find')));
  });

  test('Ctrl+F opens it, Enter steps, Escape closes and gives the focus back', async () => {
    const editor = await open();
    const scope = await mountVue(editor, () => [
      h('button', { type: 'button', id: 'before' }, 'before'),
      h(PptxEditorFindBar),
    ]);
    const before = scope.querySelector<HTMLButtonElement>('#before')!;
    before.focus();
    before.dispatchEvent(new KeyboardEvent('keydown', { key: 'f', ctrlKey: true, bubbles: true }));
    await nextTick();
    await nextTick();
    const field = scope.querySelector<HTMLInputElement>('input[aria-label="Find"]')!;
    expect(document.activeElement).toBe(field);
    expect(scope.querySelector('input[aria-label="Replace"]') === null).toBe(true);
    field.value = 'one';
    field.dispatchEvent(new Event('input'));
    await nextTick();
    expect(scope.querySelector('.pptx-find__count')?.textContent).toBe('1 of 3');
    expect(editor.snapshot().selection.kind).toBe('text');
    expect(document.activeElement).toBe(field);
    field.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
    await nextTick();
    expect(scope.querySelector('.pptx-find__count')?.textContent).toBe('2 of 3');
    field.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
    await nextTick();
    expect(scope.querySelector('.pptx-find')?.hasAttribute('hidden')).toBe(true);
    // Ctrl+H opens it with the replace row.
    before.focus();
    before.dispatchEvent(new KeyboardEvent('keydown', { key: 'h', ctrlKey: true, bubbles: true }));
    await nextTick();
    expect(scope.querySelector('input[aria-label="Replace"]') !== null).toBe(true);
  });

  test('the bar keeps its height free above the slide while it is mounted', async () => {
    const editor = await open();
    const scope = await mountVue(editor, () => h(PptxEditorFindBar));
    await nextTick();
    // Hidden, the bar reserves nothing; core's `reserveFindBarSpace` follows its height.
    expect(scope.style.getPropertyValue('--pptx-find-reserve')).toBe('0px');
    app?.unmount();
    app = null;
    expect(scope.style.getPropertyValue('--pptx-find-reserve')).toBe('');
  });
});

describe('the root hosts find, the link popup, and the slide show', () => {
  async function mountRoot(
    popups: PptxEditorPopups | undefined,
    children: () => unknown
  ): Promise<PptxEditorInstance> {
    vueHost = document.createElement('div');
    document.body.append(vueHost);
    let ready: PptxEditorInstance | null = null;
    app = createApp(
      defineComponent({
        setup: () => () =>
          h(PptxEditorRoot, { popups, onReady: (e: PptxEditorInstance) => (ready = e) }, () =>
            h(PptxEditorViewport, null, () => children())
          ),
      }) as Component
    );
    app.mount(vueHost);
    await nextTick();
    const editor = ready as PptxEditorInstance | null;
    expect(editor !== null).toBe(true);
    expect((await editor!.load(deck())).ok).toBe(true);
    await nextTick();
    return editor!;
  }

  test('Find enables only while a find bar is mounted, and runs it', async () => {
    let find: UseEditorCommandReturn | null = null;
    const Probe = defineComponent({
      setup() {
        find = useEditorCommand('find.find');
        return () => null;
      },
    });
    const shown = shallowRef(false);
    await mountRoot(undefined, () => [h(Probe), shown.value ? h(PptxEditorFindBar) : null]);
    expect(find!.isEnabled.value).toBe(false);
    expect(find!.disabledReason.value).toBe('find.noBar');
    shown.value = true;
    await nextTick();
    expect(find!.isEnabled.value).toBe(true);
    expect(find!.execute()).toBe(true);
    await nextTick();
    expect(document.querySelector('.pptx-find')?.hasAttribute('hidden')).toBe(false);
  });

  test('popups.find renders the host bar, and false turns Find off', async () => {
    const seen: boolean[] = [];
    let find: UseEditorCommandReturn | null = null;
    const Probe = defineComponent({
      setup() {
        find = useEditorCommand('find.replace');
        return () => null;
      },
    });
    const popups = shallowRef<PptxEditorPopups>({
      find: ({ replace, onClose }) => {
        seen.push(replace);
        return h('button', { class: 'acme-find', onClick: onClose });
      },
    });
    vueHost = document.createElement('div');
    document.body.append(vueHost);
    let ready: PptxEditorInstance | null = null;
    app = createApp(
      defineComponent({
        setup: () => () =>
          h(
            PptxEditorRoot,
            { popups: popups.value, onReady: (e: PptxEditorInstance) => (ready = e) },
            () => h(PptxEditorViewport, null, () => [h(Probe), h(PptxEditorFindBar)])
          ),
      }) as Component
    );
    app.mount(vueHost);
    await nextTick();
    const editor = ready as unknown as PptxEditorInstance;
    expect((await editor.load(deck())).ok).toBe(true);
    await nextTick();
    expect(find!.execute()).toBe(true);
    await nextTick();
    expect(seen.at(-1)).toBe(true);
    document.querySelector<HTMLButtonElement>('.acme-find')!.click();
    await nextTick();
    expect(document.querySelector('.acme-find') === null).toBe(true);
    popups.value = { find: false };
    await nextTick();
    expect(find!.isEnabled.value).toBe(false);
    expect(find!.disabledReason.value).toBe('find.unavailable');
    expect(document.querySelector('.pptx-find') === null).toBe(true);
  });

  test('Ctrl+K and the link command open the link popup with no toolbar on the page', async () => {
    let link: UseEditorCommandReturn | null = null;
    const Probe = defineComponent({
      setup() {
        link = useEditorCommand('text.link');
        return () => null;
      },
    });
    const editor = await mountRoot(undefined, () => [h(Probe), h(PptxEditorContent)]);
    const nodeId = editor.session!.layout(0)!.items[0]!.element.nodeId;
    const selectText = () =>
      editor.setSelection({
        kind: 'text',
        nodeId,
        anchor: { paragraph: 0, offset: 0 },
        focus: { paragraph: 0, offset: 3 },
      });
    selectText();
    await nextTick();
    // Ctrl+K on the slide text.
    const input = document.querySelector<HTMLElement>('.pptx-surface__input')!;
    input.focus();
    input.dispatchEvent(new KeyboardEvent('keydown', { key: 'k', ctrlKey: true, bubbles: true }));
    await nextTick();
    await nextTick();
    const field = document.querySelector<HTMLInputElement>('input[aria-label="Link"]')!;
    expect(field !== null).toBe(true);
    expect(document.activeElement).toBe(field);
    field.value = 'https://example.com/';
    field.dispatchEvent(new Event('input'));
    const enter = new KeyboardEvent('keydown', { key: 'Enter', bubbles: true, cancelable: true });
    field.dispatchEvent(enter);
    expect(enter.defaultPrevented).toBe(true);
    await nextTick();
    await nextTick();
    expect(document.querySelector('.pptx-link-popup') === null).toBe(true);
    expect(link!.value.value).toBe('https://example.com/');
    expect(document.activeElement).toBe(document.querySelector('.pptx-surface__input'));
    // The link command (Insert > Link runs the same host action) opens it again, with Remove.
    selectText();
    await nextTick();
    expect(link!.execute()).toBe(true);
    await nextTick();
    await nextTick();
    const remove = [...document.querySelectorAll<HTMLButtonElement>('.pptx-link button')].find(
      (b) => b.getAttribute('aria-label') === 'Remove link'
    );
    expect(remove).toBeDefined();
    remove!.click();
    await nextTick();
    await nextTick();
    expect(document.querySelector('.pptx-link-popup') === null).toBe(true);
    expect(link!.value.value ?? '').toBe('');
  });

  test('the link popup opens at the selected text, and Cancel closes it with no change', async () => {
    // The React twin is packages/react/src/__tests__/menu-parts.test.tsx.
    const editor = await mountRoot(undefined, () => [h(PptxEditorContent)]);
    const nodeId = editor.session!.layout(0)!.items[0]!.element.nodeId;
    editor.setSelection({
      kind: 'text',
      nodeId,
      anchor: { paragraph: 0, offset: 0 },
      focus: { paragraph: 0, offset: 2 },
    });
    await nextTick();
    const marks = [...document.querySelectorAll('.pptx-surface__text-selection')];
    expect(marks.length > 0).toBe(true);
    for (const mark of marks)
      mark.getBoundingClientRect = () =>
        ({ left: 100, top: 200, right: 180, bottom: 220, width: 80, height: 20 }) as DOMRect;
    requestLinkPopup(editor);
    await nextTick();
    await nextTick();
    const popup = document.querySelector<HTMLElement>('.pptx-link-popup')!;
    // Under the selected words, with the popover's 6px gap.
    expect(popup.style.left).toBe('100px');
    expect(popup.style.top).toBe('226px');
    const cancel = [...popup.querySelectorAll<HTMLButtonElement>('button')].find(
      (b) => b.textContent === 'Cancel'
    )!;
    cancel.click();
    await nextTick();
    await nextTick();
    expect(document.querySelector('.pptx-link-popup') === null).toBe(true);
    expect(editor.snapshot().format?.link ?? '').toBe('');
    expect(document.activeElement).toBe(document.querySelector('.pptx-surface__input'));
  });

  test('with no text selected, the link popup opens under the control that asked', async () => {
    const editor = await mountRoot(undefined, () => [h(PptxEditorContent)]);
    const button = document.createElement('button');
    document.body.append(button);
    button.getBoundingClientRect = () =>
      ({ left: 40, top: 10, right: 70, bottom: 30, width: 30, height: 20 }) as DOMRect;
    button.getClientRects = () => [button.getBoundingClientRect()] as unknown as DOMRectList;
    editor.selectAll();
    await nextTick();
    requestLinkPopup(editor, button);
    await nextTick();
    await nextTick();
    const popup = document.querySelector<HTMLElement>('.pptx-link-popup')!;
    expect(popup.style.left).toBe('40px');
    expect(popup.style.top).toBe('36px');
    const labels = [...popup.querySelectorAll('button')].map((b) => b.textContent);
    expect(labels).toContain('Cancel');
    expect(labels).not.toContain('Remove link');
    button.remove();
  });

  test('popups.hyperlink false turns the popup off; the API still sets links', async () => {
    let link: UseEditorCommandReturn | null = null;
    const Probe = defineComponent({
      setup() {
        link = useEditorCommand('text.link');
        return () => null;
      },
    });
    const editor = await mountRoot({ hyperlink: false }, () => [h(Probe)]);
    const nodeId = editor.session!.layout(0)!.items[0]!.element.nodeId;
    editor.setSelection({
      kind: 'text',
      nodeId,
      anchor: { paragraph: 0, offset: 0 },
      focus: { paragraph: 0, offset: 3 },
    });
    await nextTick();
    const warn = console.warn;
    console.warn = () => undefined;
    try {
      expect(link!.execute()).toBe(false);
    } finally {
      console.warn = warn;
    }
    await nextTick();
    expect(document.querySelector('.pptx-link-popup') === null).toBe(true);
    expect(link!.execute('https://example.com/')).toBe(true);
    await nextTick();
    expect(link!.value.value).toBe('https://example.com/');
  });

  test('useSlideshow starts the show, F5 starts it from the start, and Escape ends it', async () => {
    let show: UseSlideshowReturn | null = null;
    let present: UseEditorCommandReturn | null = null;
    const Probe = defineComponent({
      setup() {
        show = useSlideshow();
        present = useEditorCommand('view.playFromStart');
        return () => null;
      },
    });
    const editor = await mountRoot(undefined, () => [h(Probe), h(PptxEditorSlideshow)]);
    editor.goToSlide(1);
    await nextTick();
    show!.start('current');
    await nextTick();
    const screen = document.querySelector<HTMLElement>('.pptx-show');
    expect(screen !== null).toBe(true);
    expect(screen!.querySelector('.pptx-show__slide')?.getAttribute('data-slide-index')).toBe('1');
    expect(document.activeElement).toBe(screen);
    expect(screen!.querySelector('.pptx-show__slide svg') !== null).toBe(true);
    window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape' }));
    await nextTick();
    await nextTick();
    expect(document.querySelector('.pptx-show') === null).toBe(true);
    expect(show!.controller.value).toBeNull();
    // F5 while the focus is on nothing starts the show from the first slide.
    (document.activeElement as HTMLElement | null)?.blur();
    document.body.dispatchEvent(new KeyboardEvent('keydown', { key: 'F5', bubbles: true }));
    await nextTick();
    expect(document.querySelector('.pptx-show__slide')?.getAttribute('data-slide-index')).toBe('0');
    window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape' }));
    await nextTick();
    await nextTick();
    expect(document.querySelector('.pptx-show') === null).toBe(true);
    // The ribbon and menu rows run the same host action: Slide Show > From Beginning.
    expect(present!.isEnabled.value).toBe(true);
    expect(present!.execute()).toBe(true);
    await nextTick();
    expect(document.querySelector('.pptx-show__slide')?.getAttribute('data-slide-index')).toBe('0');
    window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape' }));
    await nextTick();
  });

  test('Escape gives focus back to the opener, or to the slide canvas when the opener left', async () => {
    let show: UseSlideshowReturn | null = null;
    const Probe = defineComponent({
      setup() {
        show = useSlideshow();
        return () => null;
      },
    });
    await mountRoot(undefined, () => [
      h('div', { class: 'pptx-surface', tabindex: 0, 'data-testid': 'canvas' }),
      h(Probe),
      h(PptxEditorSlideshow),
    ]);
    const opener = document.createElement('button');
    document.querySelector('.pptx-surface')!.after(opener);
    opener.focus();
    show!.start('start');
    await nextTick();
    window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape' }));
    await nextTick();
    await nextTick();
    expect(document.activeElement).toBe(opener);
    // A menu row starts the show and closes: it is gone when the show ends.
    opener.focus();
    show!.start('start');
    await nextTick();
    opener.remove();
    window.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape' }));
    await nextTick();
    await nextTick();
    await new Promise((r) => setTimeout(r, 50));
    expect(document.activeElement).toBe(document.querySelector('[data-testid="canvas"]'));
  });
});

describe('PptxEditorSlideshow', () => {
  test('renders the same screen as React for one show', async () => {
    const editor = await open();
    const controller = startSlideshowController(editor, 'start')!;
    const reactScope = mountReact(
      editor,
      createElement(
        SlideshowContext.Provider,
        { value: { controller, start: () => undefined, stop: () => undefined } },
        createElement(ReactSlideshow)
      )
    );
    // The painted slide is the engine's own; compare the screen around it.
    const paintedSlide = (el: Element) => el.classList.contains('pptx-show__slide');
    const reactMarkup = canonical(reactScope.querySelector('.pptx-show'), paintedSlide);
    cleanup();
    const handle: UseSlideshowReturn = {
      controller: shallowRef(controller),
      start: () => undefined,
      stop: () => undefined,
    };
    const scope = await mountVue(
      editor,
      () => h(PptxEditorSlideshow),
      () => provide(pptxSlideshowKey, handle)
    );
    const vueShow = scope.querySelector('.pptx-show');
    expect(vueShow !== null).toBe(true);
    expect(canonical(vueShow, paintedSlide)).toBe(reactMarkup);
    // A click on the slide steps to the next slide, as in React.
    (vueShow as HTMLElement).dispatchEvent(new MouseEvent('click', { bubbles: true, button: 0 }));
    await nextTick();
    expect(scope.querySelector('.pptx-show__position')?.textContent).toBe('Slide 2 of 2');
  });
});

describe('the notices', () => {
  test('PptxEditorEditNotice renders the same notice as React, and dismisses', async () => {
    const editor = await open();
    const reactScope = mountReact(editor, createElement(ReactEditNotice, { className: 'acme' }));
    const vueScope = await mountVue(editor, () => h(PptxEditorEditNotice, { class: 'acme' }));
    expect(vueScope.querySelector('.pptx-edit-notice') === null).toBe(true);
    act(() => void editor.reject('disabled.selectionChanged'));
    await nextTick();
    const vueNotice = vueScope.querySelector('.pptx-edit-notice');
    expect(vueNotice !== null).toBe(true);
    expect(canonical(vueNotice)).toBe(canonical(reactScope.querySelector('.pptx-edit-notice')));
    vueNotice!.querySelector('button')!.click();
    await nextTick();
    expect(vueScope.querySelector('.pptx-edit-notice') === null).toBe(true);
  });

  test('after a Delete, both notices say what went, and Undo brings it back', async () => {
    const editor = await open();
    expect(editor.exec({ type: 'addSlide' }).ok).toBe(true);
    const slides = editor.snapshot().slides.length;
    const reactScope = mountReact(editor, createElement(ReactEditNotice));
    const vueScope = await mountVue(editor, () => h(PptxEditorEditNotice));
    act(() => void editor.exec({ type: 'deleteSlide', index: 1 }));
    await nextTick();
    const vueNotice = vueScope.querySelector('.pptx-edit-notice');
    expect(vueNotice?.textContent).toContain('Slide deleted');
    expect(vueNotice?.getAttribute('data-tone')).toBe('info');
    expect(canonical(vueNotice)).toBe(canonical(reactScope.querySelector('.pptx-edit-notice')));
    expect(editor.snapshot().slides).toHaveLength(slides - 1);
    act(() => vueNotice!.querySelector<HTMLButtonElement>('.pptx-edit-notice__action')!.click());
    await nextTick();
    expect(editor.snapshot().slides).toHaveLength(slides);
    expect(vueScope.querySelector('.pptx-edit-notice') === null).toBe(true);
    expect(reactScope.querySelector('.pptx-edit-notice') === null).toBe(true);
  });

  test('the removal notice closes after ten seconds, but not while the pointer is on it', async () => {
    const editor = await open();
    const reactScope = mountReact(editor, createElement(ReactEditNotice));
    const vueScope = await mountVue(editor, () => h(PptxEditorEditNotice));
    jest.useFakeTimers();
    try {
      const id = editor.session!.layout(0)!.items[0]!.element.nodeId;
      act(() => editor.setSelection({ kind: 'shapes', nodeIds: [id] }));
      act(() => void editor.exec({ type: 'delete' }));
      await nextTick();
      for (const scope of [reactScope, vueScope])
        expect(scope.querySelector('.pptx-edit-notice')?.textContent).toContain('Shape deleted');
      // The pointer holds the React notice; the Vue one closes on time.
      act(() => {
        reactScope
          .querySelector('.pptx-edit-notice')!
          .dispatchEvent(new MouseEvent('mouseover', { bubbles: true }));
      });
      act(() => void jest.advanceTimersByTime(10_000));
      await nextTick();
      expect(vueScope.querySelector('.pptx-edit-notice') === null).toBe(true);
      expect(reactScope.querySelector('.pptx-edit-notice') !== null).toBe(true);
      act(() => {
        reactScope
          .querySelector('.pptx-edit-notice')!
          .dispatchEvent(new MouseEvent('mouseout', { bubbles: true }));
      });
      act(() => void jest.advanceTimersByTime(10_000));
      expect(reactScope.querySelector('.pptx-edit-notice') === null).toBe(true);
    } finally {
      jest.useRealTimers();
    }
  });

  test('Cut through the slide surface reports no removal', async () => {
    const editor = await open();
    const host = document.createElement('div');
    document.body.append(host);
    editor.attach(host);
    const id = editor.session!.layout(0)!.items[0]!.element.nodeId;
    editor.setSelection({ kind: 'shapes', nodeIds: [id] });
    const data = new Map<string, string>();
    const cut = new Event('cut', { bubbles: true, cancelable: true });
    Object.defineProperty(cut, 'clipboardData', {
      value: { setData: (type: string, value: string) => void data.set(type, value) },
    });
    host.querySelector('.pptx-surface')!.dispatchEvent(cut);
    expect(editor.session!.layout(0)!.items).toHaveLength(0);
    expect(data.size).toBeGreaterThan(0);
    expect(editor.snapshot().lastRemoval).toBeNull();
  });

  test('Cut from the menu reports no removal', async () => {
    const editor = await open();
    const id = editor.session!.layout(0)!.items[0]!.element.nodeId;
    editor.setSelection({ kind: 'shapes', nodeIds: [id] });
    runClipboardSlot(editor, 'edit.cut', null);
    expect(editor.session!.layout(0)!.items).toHaveLength(0);
    expect(editor.snapshot().lastRemoval).toBeNull();
  });

  test('after Undo under the pointer, the next removal notice still closes on time', async () => {
    const editor = await open();
    expect(editor.exec({ type: 'addSlide' }).ok).toBe(true);
    const reactScope = mountReact(editor, createElement(ReactEditNotice));
    const vueScope = await mountVue(editor, () => h(PptxEditorEditNotice));
    jest.useFakeTimers();
    try {
      act(() => void editor.exec({ type: 'deleteSlide', index: 1 }));
      await nextTick();
      const react = reactScope.querySelector('.pptx-edit-notice')!;
      const vue = vueScope.querySelector('.pptx-edit-notice')!;
      // The pointer is on both notices when Undo removes them: no leave event follows.
      act(() => void react.dispatchEvent(new MouseEvent('mouseover', { bubbles: true })));
      vue.dispatchEvent(new MouseEvent('mouseenter'));
      act(() => vue.querySelector<HTMLButtonElement>('.pptx-edit-notice__action')!.click());
      await nextTick();
      act(() => void editor.exec({ type: 'deleteSlide', index: 1 }));
      await nextTick();
      for (const scope of [reactScope, vueScope])
        expect(scope.querySelector('.pptx-edit-notice')?.textContent).toContain('Slide deleted');
      act(() => void jest.advanceTimersByTime(10_000));
      await nextTick();
      for (const scope of [reactScope, vueScope])
        expect(scope.querySelector('.pptx-edit-notice') === null).toBe(true);
    } finally {
      jest.useRealTimers();
    }
  });

  test('rejectionMessage reads the reason and its detail', async () => {
    let t: ReturnType<typeof useTranslation>['t'] | null = null;
    const editor = await open();
    await mountVue(
      editor,
      () => null,
      () => void (t = useTranslation().t)
    );
    expect(rejectionMessage(t!, { reason: 'Plain text' })).toBe('Plain text');
  });

  test('PptxEditorFontNotice lists the families with a substitute, as React does', async () => {
    const real = await open();
    const editor = Object.create(real) as PptxEditorInstance;
    // One snapshot object per engine snapshot, as the engine gives.
    let seen: unknown = null;
    let cached: ReturnType<PptxEditorInstance['snapshot']> | null = null;
    editor.snapshot = () => {
      const now = real.snapshot();
      if (now !== seen || !cached) {
        seen = now;
        cached = {
          ...now,
          substitutedFonts: [
            { family: 'Acme Display', substitute: 'Carlito', metricCompatible: false },
            { family: 'Calibri', substitute: 'Carlito', metricCompatible: true },
          ] as never,
        };
      }
      return cached;
    };
    const reactScope = mountReact(editor, createElement(ReactFontNotice));
    const vueScope = await mountVue(editor, () => h(PptxEditorFontNotice));
    const vueNotice = vueScope.querySelector('.pptx-font-notice');
    expect(vueNotice?.textContent).toContain('Acme Display');
    expect(vueNotice?.textContent).not.toContain('Calibri');
    expect(canonical(vueNotice)).toBe(canonical(reactScope.querySelector('.pptx-font-notice')));
    vueNotice!.querySelector<HTMLButtonElement>('button[aria-label]')!.click();
    await nextTick();
    expect(vueScope.querySelector('.pptx-font-notice') === null).toBe(true);
  });

  test('a newly opened document drops the outcome of an earlier local font request', async () => {
    const real = await open();
    const editor = Object.create(real) as PptxEditorInstance;
    let seen: unknown = null;
    let cached: ReturnType<PptxEditorInstance['snapshot']> | null = null;
    editor.snapshot = () => {
      const now = real.snapshot();
      if (now !== seen || !cached) {
        seen = now;
        cached = {
          ...now,
          substitutedFonts: [
            { family: 'Acme Display', substitute: 'Carlito', metricCompatible: false },
          ] as never,
        };
      }
      return cached;
    };
    const g = globalThis as { queryLocalFonts?: () => Promise<unknown[]> };
    const before = g.queryLocalFonts;
    g.queryLocalFonts = async () => [];
    try {
      const reactScope = mountReact(editor, createElement(ReactFontNotice));
      const vueScope = await mountVue(editor, () => h(PptxEditorFontNotice));
      const none = 'None of these fonts are installed on this computer.';
      for (const scope of [reactScope, vueScope]) {
        const ask = scope.querySelector<HTMLButtonElement>('.pptx-font-notice__action')!;
        await act(async () => ask.click());
        await nextTick();
        expect(scope.querySelector('.pptx-font-notice')?.textContent).toContain(none);
      }
      await act(async () => {
        expect((await real.load(deck())).ok).toBe(true);
      });
      await nextTick();
      for (const scope of [reactScope, vueScope]) {
        const text = scope.querySelector('.pptx-font-notice')?.textContent ?? '';
        expect(text).toContain('Acme Display');
        expect(text).not.toContain(none);
      }
    } finally {
      g.queryLocalFonts = before;
    }
  });

  test('a later change to the substituted families hides the outcome of a request', async () => {
    const real = await open();
    const editor = Object.create(real) as PptxEditorInstance;
    const acme = { family: 'Acme Display', substitute: 'Carlito', metricCompatible: false };
    let families = [acme];
    let seen: unknown = null;
    let cached: ReturnType<PptxEditorInstance['snapshot']> | null = null;
    editor.snapshot = () => {
      const now = real.snapshot();
      if (now !== seen || !cached) {
        seen = now;
        cached = { ...now, substitutedFonts: families as never };
      }
      return cached;
    };
    const g = globalThis as { queryLocalFonts?: () => Promise<unknown[]> };
    const before = g.queryLocalFonts;
    g.queryLocalFonts = async () => [];
    try {
      const reactScope = mountReact(editor, createElement(ReactFontNotice));
      const vueScope = await mountVue(editor, () => h(PptxEditorFontNotice));
      const none = 'None of these fonts are installed on this computer.';
      for (const scope of [reactScope, vueScope]) {
        const ask = scope.querySelector<HTMLButtonElement>('.pptx-font-notice__action')!;
        await act(async () => ask.click());
        await nextTick();
        expect(scope.querySelector('.pptx-font-notice')?.textContent).toContain(none);
      }
      // A host font that arrives later changes the list: the outcome no longer applies.
      families = [acme, { family: 'Beta Sans', substitute: 'Carlito', metricCompatible: false }];
      await act(async () => real.setMode('view'));
      await nextTick();
      for (const scope of [reactScope, vueScope]) {
        const text = scope.querySelector('.pptx-font-notice')?.textContent ?? '';
        expect(text).toContain('Beta Sans');
        expect(text).not.toContain(none);
      }
    } finally {
      g.queryLocalFonts = before;
    }
  });

  test('local fonts that resolve every family report that they are in use', async () => {
    const real = await open();
    const editor = Object.create(real) as PptxEditorInstance;
    // The family resolves once the local faces are set, before `setFonts` settles, as a real
    // font load does: the notice's family list changes while the request is still open.
    let added = false;
    editor.setFonts = (config) => {
      added = true;
      return real.setFonts(config);
    };
    let seen: unknown = null;
    let cached: ReturnType<PptxEditorInstance['snapshot']> | null = null;
    editor.snapshot = () => {
      const now = real.snapshot();
      if (now !== seen || !cached) {
        seen = now;
        cached = {
          ...now,
          substitutedFonts: (added
            ? []
            : [
                { family: 'Acme Display', substitute: 'Carlito', metricCompatible: false },
              ]) as never,
        };
      }
      return cached;
    };
    const bytes = fontBytes('LiberationSans-Regular.ttf');
    const g = globalThis as { queryLocalFonts?: () => Promise<unknown[]> };
    const before = g.queryLocalFonts;
    g.queryLocalFonts = async () => [
      {
        family: 'Acme Display',
        style: 'Regular',
        blob: async () => new Blob([bytes as Uint8Array<ArrayBuffer>]),
      },
    ];
    try {
      const mounts = [
        async () => mountReact(editor, createElement(ReactFontNotice)),
        () => mountVue(editor, () => h(PptxEditorFontNotice)),
      ];
      for (const mount of mounts) {
        added = false;
        cached = null;
        const scope = await mount();
        const ask = scope.querySelector<HTMLButtonElement>('.pptx-font-notice__action')!;
        await act(async () => ask.click());
        await act(async () => {
          await new Promise((r) => setTimeout(r, 0));
        });
        await nextTick();
        expect(scope.querySelector('.pptx-font-notice')?.textContent).toContain(
          'Fonts from this computer are now in use.'
        );
      }
    } finally {
      g.queryLocalFonts = before;
    }
  });

  test('a local font request that ends after another document opens reports nothing', async () => {
    const real = await open();
    const editor = Object.create(real) as PptxEditorInstance;
    let seen: unknown = null;
    let cached: ReturnType<PptxEditorInstance['snapshot']> | null = null;
    editor.snapshot = () => {
      const now = real.snapshot();
      if (now !== seen || !cached) {
        seen = now;
        cached = {
          ...now,
          substitutedFonts: [
            { family: 'Acme Display', substitute: 'Carlito', metricCompatible: false },
          ] as never,
        };
      }
      return cached;
    };
    const g = globalThis as { queryLocalFonts?: () => Promise<unknown[]> };
    const before = g.queryLocalFonts;
    const answers: ((fonts: unknown[]) => void)[] = [];
    g.queryLocalFonts = () => new Promise((resolve) => answers.push(resolve));
    try {
      const reactScope = mountReact(editor, createElement(ReactFontNotice));
      const vueScope = await mountVue(editor, () => h(PptxEditorFontNotice));
      for (const scope of [reactScope, vueScope])
        await act(async () =>
          scope.querySelector<HTMLButtonElement>('.pptx-font-notice__action')!.click()
        );
      await act(async () => {
        expect((await real.load(deck())).ok).toBe(true);
      });
      await act(async () => {
        for (const answer of answers) answer([]);
      });
      await nextTick();
      for (const scope of [reactScope, vueScope])
        expect(scope.querySelector('.pptx-font-notice')?.textContent).not.toContain(
          'None of these fonts are installed'
        );
    } finally {
      g.queryLocalFonts = before;
    }
  });
});

describe('HyperlinkPopup', () => {
  test('renders the same popup as React for a link', async () => {
    const editor = await open();
    const reactScope = mountReact(
      editor,
      createElement('button', { type: 'button', id: 'r-anchor', 'aria-label': 'Link' })
    );
    const reactAnchor = reactScope.querySelector<HTMLElement>('#r-anchor')!;
    render(
      createElement(
        PptxEditorContext.Provider,
        { value: editor },
        createElement(ReactHyperlinkPopup, {
          anchor: reactAnchor,
          url: 'https://example.com/',
          onClose: () => undefined,
          className: 'acme',
        })
      )
    );
    const reactPopup = canonical(reactScope.querySelector('.pptx-link-popup'));
    expect(reactPopup).not.toBe('');
    cleanup();
    const anchorScope = document.createElement('div');
    anchorScope.className = 'pptx-editor';
    const anchor = document.createElement('button');
    anchor.setAttribute('aria-label', 'Link');
    anchorScope.append(anchor);
    document.body.append(anchorScope);
    try {
      await mountVue(editor, () =>
        h(HyperlinkPopup, {
          anchor,
          url: 'https://example.com/',
          onClose: () => undefined,
          className: 'acme',
        })
      );
      const popup = anchorScope.querySelector('.pptx-link-popup');
      expect(popup !== null).toBe(true);
      expect(canonical(popup)).toBe(reactPopup);
    } finally {
      anchorScope.remove();
    }
  });
});
