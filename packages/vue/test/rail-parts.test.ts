// The Vue slide rail parts render the same markup as their React twins
// (packages/react/src/rail-parts.tsx), and drive the editor the same way.

import { afterEach, describe, expect, test } from 'bun:test';
import { createElement, type FunctionComponent, type ReactElement } from 'react';
import { cleanup, render } from '@testing-library/react';
import {
  createApp,
  createCommentVNode,
  defineComponent,
  h,
  nextTick,
  shallowRef,
  type App,
  type Component,
} from 'vue';
import { createPptxEditor, type PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import { buildPptx } from '../../../scripts/fixtures/builder.ts';
import { PptxEditorContext } from '../../react/src/context.ts';
import { SlideRail as RSlideRailNamespace } from '../../react/src/index.ts';
import type { SlideRailProps as RSlideRailProps } from '../../react/src/panes.tsx';
import { PptxEditorSlideRail } from '../src/index.ts';
import { providePptxEditor } from '../src/composables.ts';
import { canonicalDom } from './canonical-dom.ts';

const RSlideRail = RSlideRailNamespace as unknown as FunctionComponent<RSlideRailProps>;

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

async function mountVue(renderFn: () => unknown, editor: PptxEditorInstance) {
  host = document.createElement('div');
  document.body.append(host);
  app = createApp(
    defineComponent({
      setup() {
        providePptxEditor(shallowRef(editor));
        return renderFn;
      },
    }) as Component
  );
  app.mount(host);
  await nextTick();
  return host;
}

const canon = (root: Element): string =>
  canonicalDom(root, {
    normalizeStyle: true,
    skipChildren: (el) => el.classList.contains('pptx-rail__thumb'),
  });

async function same(react: ReactElement, vue: () => unknown, editor: PptxEditorInstance) {
  const r = canon(
    render(createElement(PptxEditorContext.Provider, { value: editor }, react)).container
  );
  cleanup();
  const v = canon(await mountVue(vue, editor));
  expect(v).toEqual(r);
}

async function deck(): Promise<PptxEditorInstance> {
  const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await editor.load(buildPptx({ slides: [{ shapes: '' }, { shapes: '' }] }))).ok).toBe(
    true
  );
  return editor;
}

describe('Vue slide rail parts match React', () => {
  test('hidden, className, and a host action', async () => {
    const editor = await deck();
    const R = RSlideRailNamespace;
    await same(
      createElement(
        RSlideRail,
        { orientation: 'horizontal' },
        createElement(R.MoveEarlier, { hidden: true }),
        createElement(R.MoveLater, { hidden: true }),
        createElement(R.Duplicate, { className: 'host-dup' }),
        createElement(R.Action, { label: 'Send' })
      ),
      () =>
        h(
          PptxEditorSlideRail,
          { orientation: 'horizontal' },
          {
            default: () => [
              h(PptxEditorSlideRail.MoveEarlier, { hidden: true }),
              // A bare `hidden` attribute in a template arrives as an empty string.
              h(PptxEditorSlideRail.MoveLater, { hidden: '' as unknown as boolean }),
              h(PptxEditorSlideRail.Duplicate, { className: 'host-dup' }),
              h(PptxEditorSlideRail.Action, { label: 'Send' }),
            ],
          }
        ),
      editor
    );
    expect(
      document
        .querySelector<HTMLElement>('.pptx-rail__actions')!
        .style.getPropertyValue('--pptx-rail-action-columns')
    ).toBe('3');
  });

  test('preset false renders only the children, and a part runs its command', async () => {
    const editor = await deck();
    const R = RSlideRailNamespace;
    await same(
      createElement(RSlideRail, { preset: false }, createElement(R.AddSlide)),
      () =>
        h(
          PptxEditorSlideRail,
          { preset: false },
          { default: () => [h(PptxEditorSlideRail.AddSlide)] }
        ),
      editor
    );
    document.querySelector<HTMLElement>('.pptx-rail__add')!.click();
    expect(editor.snapshot().slides.length).toBe(3);
  });

  test('a host icon replaces the packaged icon', async () => {
    const editor = await deck();
    await mountVue(
      () =>
        h(PptxEditorSlideRail, null, {
          default: () => [
            h(PptxEditorSlideRail.Delete, { icon: h('svg', { class: 'host-icon' }) }),
          ],
        }),
      editor
    );
    const del = document.querySelector<HTMLElement>('.pptx-rail__delete')!;
    expect(del.querySelector('.host-icon') !== null).toBe(true);
    expect(del.querySelector('.pptx-icon')).toBeNull();
  });

  test('asChild hands the part to the host element, as React', async () => {
    const editor = await deck();
    const R = RSlideRailNamespace;
    await same(
      createElement(
        RSlideRail,
        null,
        createElement(
          R.AddSlide,
          { asChild: true },
          createElement('button', { className: 'host-add' }, 'Add')
        ),
        createElement(
          R.MoveEarlier,
          { asChild: true },
          createElement('span', { className: 'host-earlier' }, 'Up')
        )
      ),
      () =>
        h(PptxEditorSlideRail, null, {
          default: () => [
            h(
              PptxEditorSlideRail.AddSlide,
              { asChild: true },
              {
                default: () => [h('button', { class: 'host-add' }, 'Add')],
              }
            ),
            h(
              PptxEditorSlideRail.MoveEarlier,
              { asChild: true },
              {
                default: () => [h('span', { class: 'host-earlier' }, 'Up')],
              }
            ),
          ],
        }),
      editor
    );
    const earlier = document.querySelector<HTMLElement>('.host-earlier')!;
    expect(earlier.getAttribute('aria-disabled')).toBe('true');
    expect(earlier.getAttribute('role')).toBe('button');
    const add = document.querySelector<HTMLElement>('.host-add')!;
    expect(add.getAttribute('aria-label')).toBe('New slide');
    add.click();
    expect(editor.snapshot().slides.length).toBe(3);
  });

  test('a v-if that is false counts as no action, as React null', async () => {
    const editor = await deck();
    const R = RSlideRailNamespace;
    await same(
      createElement(
        RSlideRail,
        { orientation: 'horizontal', preset: false },
        createElement(R.AddSlide),
        null,
        createElement(R.Delete)
      ),
      () =>
        h(
          PptxEditorSlideRail,
          { orientation: 'horizontal', preset: false },
          {
            default: () => [
              h(PptxEditorSlideRail.AddSlide),
              createCommentVNode('v-if', true),
              h(PptxEditorSlideRail.Delete),
            ],
          }
        ),
      editor
    );
    expect(columns()).toBe('1');
  });

  test('a horizontal rail with toggle false keeps two rows of three', async () => {
    const editor = await deck();
    await mountVue(
      () => h(PptxEditorSlideRail, { orientation: 'horizontal', toggle: false }),
      editor
    );
    expect(document.querySelector('.pptx-rail__collapse')).toBeNull();
    expect(columns()).toBe('3');
  });

  test('toggle false removes Collapse under preset false, as React', async () => {
    const editor = await deck();
    const R = RSlideRailNamespace;
    await same(
      createElement(
        RSlideRail,
        { preset: false, toggle: false },
        createElement(R.AddSlide),
        createElement(R.Collapse)
      ),
      () =>
        h(
          PptxEditorSlideRail,
          { preset: false, toggle: false },
          { default: () => [h(PptxEditorSlideRail.AddSlide), h(PptxEditorSlideRail.Collapse)] }
        ),
      editor
    );
    expect(document.querySelector('.pptx-rail__collapse')).toBeNull();
    expect(columns()).toBe('1');
  });

  test('Expand changes the folded strip button, as React', async () => {
    const editor = await deck();
    const R = RSlideRailNamespace;
    await same(
      createElement(
        RSlideRail,
        { defaultOpen: false },
        createElement(R.Expand, { className: 'host-expand' })
      ),
      () =>
        h(
          PptxEditorSlideRail,
          { defaultOpen: false },
          { default: () => [h(PptxEditorSlideRail.Expand, { className: 'host-expand' })] }
        ),
      editor
    );
    const expand = document.querySelector<HTMLElement>('.host-expand')!;
    expect(expand.getAttribute('aria-label')).toBe('Show thumbnails');
    expand.click();
    await nextTick();
    expect(document.querySelector('.pptx-rail')!.getAttribute('data-open')).toBe('true');
    expect(document.querySelector('.host-expand')).toBeNull();
  });

  test('the rail keeps the current slide in view when it resizes', async () => {
    const observers: { target: Element | null; run: () => void; off: boolean }[] = [];
    const saved = globalThis.ResizeObserver;
    globalThis.ResizeObserver = class {
      readonly entry: { target: Element | null; run: () => void; off: boolean };
      constructor(run: () => void) {
        this.entry = { target: null, run, off: false };
        observers.push(this.entry);
      }
      observe(target: Element) {
        this.entry.target = target;
      }
      unobserve() {}
      disconnect() {
        this.entry.off = true;
      }
    } as unknown as typeof ResizeObserver;
    try {
      const editor = await deck();
      await mountVue(() => h(PptxEditorSlideRail, { orientation: 'horizontal' }), editor);
      const list = document.querySelector<HTMLElement>('.pptx-rail__list')!;
      const watcher = observers.find((o) => o.target === list && !o.off)!;
      expect(watcher).toBeDefined();
      editor.goToSlide(1);
      await nextTick();
      const style = document.createElement('style');
      style.textContent = '.pptx-rail__list { overflow-x: auto; }';
      document.head.append(style);
      const item = document.querySelector<HTMLElement>('.pptx-rail__item[data-index="1"]')!;
      const scrolled: number[] = [];
      Object.defineProperties(list, {
        scrollWidth: { configurable: true, value: 600 },
        clientWidth: { configurable: true, value: 200 },
        getBoundingClientRect: {
          configurable: true,
          value: () => ({ top: 0, left: 0, width: 200, height: 100 }) as DOMRect,
        },
        scrollBy: {
          configurable: true,
          value: (to: ScrollToOptions) => scrolled.push(to.left ?? 0),
        },
      });
      Object.defineProperty(item, 'getBoundingClientRect', {
        configurable: true,
        value: () => ({ top: 0, left: 300, width: 100, height: 80 }) as DOMRect,
      });
      watcher.run();
      style.remove();
      expect(scrolled).toEqual([200]);
      app?.unmount();
      app = null;
      expect(watcher.off).toBe(true);
    } finally {
      globalThis.ResizeObserver = saved;
    }
  });
});

function columns(): string {
  return document
    .querySelector<HTMLElement>('.pptx-rail__actions')!
    .style.getPropertyValue('--pptx-rail-action-columns');
}

describe('the slide menu of a row rail (React and Vue)', () => {
  const rows = (): string[] =>
    ['slide.moveUp', 'slide.moveDown'].map(
      (slot) => document.querySelector(`[data-slot="${slot}"]`)?.textContent ?? ''
    );
  const openMenu = (): void =>
    void document
      .querySelector<HTMLElement>('.pptx-rail__item[data-index="0"]')!
      .dispatchEvent(
        new MouseEvent('contextmenu', { bubbles: true, cancelable: true, clientX: 5, clientY: 5 })
      );

  const buttons = (): string[] =>
    [...document.querySelectorAll<HTMLElement>('.pptx-rail__actions .pptx-rail__move')].map(
      (b) => b.getAttribute('aria-label') ?? ''
    );

  for (const [orientation, direction, earlier, later, earlierButton] of [
    ['horizontal', 'ltr', /Move left.*←/, /Move right.*→/, 'Move slide left'],
    ['horizontal', 'rtl', /Move right.*→/, /Move left.*←/, 'Move slide right'],
    ['vertical', 'ltr', /Move up.*↑/, /Move down.*↓/, 'Move slide up'],
    ['vertical', 'rtl', /Move up.*↑/, /Move down.*↓/, 'Move slide up'],
  ] as const) {
    test(`a ${orientation} ${direction} rail names its move rows and shortcuts by its direction`, async () => {
      const editor = await deck();
      render(
        createElement(
          'div',
          { style: { direction } },
          createElement(
            PptxEditorContext.Provider,
            { value: editor },
            createElement(RSlideRail, { orientation })
          )
        )
      );
      await new Promise((r) => setTimeout(r, 0));
      const reactButtons = buttons();
      expect(reactButtons[0]).toBe(earlierButton);
      openMenu();
      await new Promise((r) => setTimeout(r, 0));
      const react = rows();
      expect(react[0]).toMatch(earlier);
      expect(react[1]).toMatch(later);
      cleanup();
      document.body.replaceChildren();
      await mountVue(
        () => h('div', { style: { direction } }, [h(PptxEditorSlideRail, { orientation })]),
        editor
      );
      await nextTick();
      expect(buttons()).toEqual(reactButtons);
      openMenu();
      await nextTick();
      expect(rows()).toEqual(react);
    });
  }
});

describe('a narrow rail folds its secondary actions into one menu (Vue)', () => {
  // The React twin is packages/react/src/__tests__/rail-parts.test.tsx.
  let restore: (() => void) | null = null;
  /** Every `.pptx-rail` measures `width` pixels wide. */
  const railWidth = (width: number): void => {
    const saved = Object.getOwnPropertyDescriptor(HTMLElement.prototype, 'clientWidth');
    Object.defineProperty(HTMLElement.prototype, 'clientWidth', {
      configurable: true,
      get(this: HTMLElement) {
        return this.classList.contains('pptx-rail') ? width : 0;
      },
    });
    restore = () => {
      if (saved) Object.defineProperty(HTMLElement.prototype, 'clientWidth', saved);
      else delete (HTMLElement.prototype as { clientWidth?: number }).clientWidth;
    };
  };
  afterEach(() => {
    restore?.();
    restore = null;
  });
  const actionClasses = () =>
    [...document.querySelectorAll<HTMLElement>('.pptx-rail__actions > button')].map(
      (b) => b.className.replace('pptx-toolbar__button ', '').split(' ')[0]
    );
  const rowLabels = (): string[] =>
    [...document.querySelectorAll('.pptx-contextmenu [role="menuitem"]')].map(
      (row) => row.querySelector('.pptx-menu__label')?.textContent ?? ''
    );

  test('at 390px, New slide and More stay; the menu holds the rest and runs them', async () => {
    railWidth(390);
    const editor = await deck();
    await mountVue(() => h(PptxEditorSlideRail, { orientation: 'horizontal' }), editor);
    await nextTick();
    expect(actionClasses()).toEqual(['pptx-rail__add', 'pptx-rail__more', 'pptx-rail__collapse']);
    const more = document.querySelector<HTMLElement>('.pptx-rail__more')!;
    expect(more.getAttribute('aria-label')).toBe('More slide actions');
    more.click();
    await nextTick();
    expect(rowLabels()).toEqual([
      'Add slide with layout',
      'Duplicate slide',
      'Move slide left',
      'Move slide right',
      'Delete slide',
    ]);
    expect(more.getAttribute('aria-expanded')).toBe('true');
    [...document.querySelectorAll<HTMLElement>('.pptx-contextmenu [role="menuitem"]')]
      .find((row) => row.textContent?.includes('Duplicate slide'))!
      .click();
    await nextTick();
    expect(editor.snapshot().slides.length).toBe(3);
    expect(document.querySelector('.pptx-contextmenu')).toBeNull();
  });

  test('a hidden part stays hidden, and with every secondary part hidden no More shows', async () => {
    railWidth(390);
    const editor = await deck();
    await mountVue(
      () =>
        h(PptxEditorSlideRail, { orientation: 'horizontal' }, () => [
          h(PptxEditorSlideRail.Duplicate, { hidden: true }),
        ]),
      editor
    );
    await nextTick();
    document.querySelector<HTMLElement>('.pptx-rail__more')!.click();
    await nextTick();
    expect(rowLabels()).not.toContain('Duplicate slide');
    expect(rowLabels().length).toBe(4);
    app?.unmount();
    host?.remove();
    await mountVue(
      () =>
        h(PptxEditorSlideRail, { orientation: 'horizontal' }, () => [
          h(PptxEditorSlideRail.ChooseLayout, { hidden: true }),
          h(PptxEditorSlideRail.Duplicate, { hidden: true }),
          h(PptxEditorSlideRail.MoveEarlier, { hidden: true }),
          h(PptxEditorSlideRail.MoveLater, { hidden: true }),
          h(PptxEditorSlideRail.Delete, { hidden: true }),
        ]),
      editor
    );
    await nextTick();
    expect(actionClasses()).toEqual(['pptx-rail__add', 'pptx-rail__collapse']);
  });

  test('a wide rail and a vertical rail keep every button', async () => {
    railWidth(800);
    const editor = await deck();
    await mountVue(() => h(PptxEditorSlideRail, { orientation: 'horizontal' }), editor);
    await nextTick();
    expect(document.querySelector('.pptx-rail__more')).toBeNull();
    expect(actionClasses().length).toBe(7);
    app?.unmount();
    host?.remove();
    restore?.();
    railWidth(200);
    await mountVue(() => h(PptxEditorSlideRail), editor);
    await nextTick();
    expect(document.querySelector('.pptx-rail__more')).toBeNull();
  });
});
