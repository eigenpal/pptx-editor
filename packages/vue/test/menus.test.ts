// The Vue menu bar, canvas context menu, slide menu, and section menu render the same markup,
// ARIA, and keys as their React twins (packages/react/src/menu-parts.tsx,
// context-menu-parts.tsx, ContextMenu.tsx, CanvasContextMenu.tsx, RailContextMenu.tsx), with
// DOCX Editor's menu customization (docs/site/content/customization/menus.mdx). Each parity
// case renders both adapters on the same deck and compares a canonical form of the DOM.

import { afterEach, describe, expect, spyOn, test } from 'bun:test';
import { createElement, type ReactElement } from 'react';
import { act, cleanup, render } from '@testing-library/react';
import { createApp, defineComponent, h, nextTick, shallowRef, type App, type Component } from 'vue';
import {
  createPptxEditor,
  resetDevWarnings,
  type PptxEditorInstance,
} from '@pptx-editor.dev/core/editor';
import { buildPptx, para, run, shape } from '../../../scripts/fixtures/builder.ts';
import { buildSlideLayoutsDeck } from '../../../scripts/fixtures/slide-layouts.ts';
import { PptxEditorContext } from '../../react/src/context.ts';
import { PptxEditorMenu as RMenu } from '../../react/src/index.ts';
import { ContextMenu as RContextMenu, MenuRow as RMenuRow } from '../../react/src/ContextMenu.tsx';
import { RailContextMenu as RRailContextMenu } from '../../react/src/RailContextMenu.tsx';
import { SectionContextMenu as RSectionContextMenu } from '../../react/src/RailSections.tsx';
import {
  ContextMenu,
  MenuRow,
  PptxEditorContent,
  PptxEditorContextMenu,
  PptxEditorMenu,
  PptxEditorRoot,
  RailContextMenu,
  SectionContextMenu,
  useContextMenuRequest,
  useContextMenuTarget,
  type PptxEditorPopups,
} from '../src/index.ts';
import { providePptxEditor } from '../src/composables.ts';
import { canonicalDom } from './canonical-dom.ts';
import { withHalos } from '../../core/src/editor/surface-halo.ts';

const IN = 914_400;
const PX = 9_525;

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
  await nextTick();
  return host;
}

function unmountVue(): void {
  app?.unmount();
  host?.remove();
  app = null;
  host = null;
}

function mountReact(element: ReactElement, editor: PptxEditorInstance | null): HTMLElement {
  return render(createElement(PptxEditorContext.Provider, { value: editor }, element)).container;
}

/** The canonical form of a menu element itself. */
const canon = (root: Element | null): string => canonicalDom(root, { self: true });

async function deck(): Promise<PptxEditorInstance> {
  const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false, zoom: 1 });
  const bytes = buildPptx({
    slides: [
      {
        layout: 'blank',
        shapes: shape({ id: 2, x: IN, y: IN, cx: 2 * IN, cy: IN, text: para(run('Hi')) }),
      },
      { layout: 'blank', shapes: '' },
      { layout: 'blank', shapes: '' },
    ],
  });
  expect((await editor.load(bytes)).ok).toBe(true);
  return editor;
}

const bar = () => document.querySelector('.pptx-menubar');
const panel = () => document.querySelector<HTMLElement>('.pptx-menubar__panel');
const trigger = (name: string) =>
  [...document.querySelectorAll<HTMLElement>('.pptx-menubar__trigger')].find(
    (b) => b.textContent === name
  )!;
const triggers = () =>
  [...document.querySelectorAll('.pptx-menubar__trigger')].map((b) => b.textContent);
const rows = (el: Element | null) =>
  [...(el?.querySelectorAll('[data-slot]') ?? [])].map((r) => r.getAttribute('data-slot'));

const MENUS = ['File', 'Edit', 'View', 'Insert', 'Format', 'Arrange', 'Slide Show'];

describe('PptxEditorMenu', () => {
  test('the bar and every menu match React', async () => {
    const editor = await deck();
    const nodeId = editor.session!.layout(0)!.items[0]!.element.nodeId;
    editor.setSelection({ kind: 'shapes', nodeIds: [nodeId] });
    const react: Record<string, string> = {};
    mountReact(createElement(RMenu), editor);
    react.bar = canon(bar());
    for (const name of MENUS) {
      act(() => trigger(name).click());
      react[name] = canon(panel());
      act(() => trigger(name).click());
    }
    cleanup();
    await mountVue(() => h(PptxEditorMenu), editor);
    expect(canon(bar())).toBe(react.bar!);
    for (const name of MENUS) {
      trigger(name).click();
      await nextTick();
      expect(canon(panel())).toBe(react[name]!);
      trigger(name).click();
      await nextTick();
    }
  });

  test('a row runs its slot and closes the menu', async () => {
    const editor = await deck();
    const nodeId = editor.session!.layout(0)!.items[0]!.element.nodeId;
    editor.setSelection({ kind: 'shapes', nodeIds: [nodeId] });
    await mountVue(() => h(PptxEditorMenu), editor);
    trigger('Edit').click();
    await nextTick();
    panel()!.querySelector<HTMLElement>('[data-slot="more.duplicate"]')!.click();
    await nextTick();
    expect(editor.session!.layout(0)!.items.length).toBe(2);
    expect(panel() === null).toBe(true);
  });

  test('the pointer moves the focused row, so one row looks active', async () => {
    const editor = await deck();
    await mountVue(() => h(PptxEditorMenu), editor);
    trigger('Insert').focus();
    trigger('Insert').click();
    await nextTick();
    await nextTick();
    const items = [...panel()!.querySelectorAll<HTMLElement>('[role^="menuitem"]')];
    expect(document.activeElement === items[0]).toBe(true);
    const label = items[3]!.querySelector('.pptx-menu__label')!;
    // The first move only records the pointer, so a menu that opens under a resting pointer
    // keeps its focus; a real move focuses the row under it.
    const move = (x: number) =>
      label.dispatchEvent(new MouseEvent('mousemove', { bubbles: true, clientX: x, clientY: 10 }));
    move(10);
    expect(document.activeElement === items[0]).toBe(true);
    move(10);
    expect(document.activeElement === items[0]).toBe(true);
    move(14);
    expect(document.activeElement === items[3]).toBe(true);
  });

  test('named menus change rows in place, hide menus, and add host rows', async () => {
    const editor = await deck();
    let opened = 0;
    await mountVue(
      () =>
        h(PptxEditorMenu, { class: 'acme-bar' }, () => [
          h(PptxEditorMenu.File, null, () => [
            h(PptxEditorMenu.Row, {
              label: 'Open…',
              icon: 'folder_open',
              onSelect: () => (opened += 1),
            }),
          ]),
          h(PptxEditorMenu.Edit, null, () => [
            h(PptxEditorMenu.Item, { slotId: 'history.redo', hidden: true }),
            h(PptxEditorMenu.Item, { slotId: 'history.undo', class: 'acme-undo' }),
          ]),
          h(PptxEditorMenu.Arrange, { hidden: true }),
          h(PptxEditorMenu.Entry, { id: 'help', label: 'Help', class: 'acme-help' }, () => [
            h(PptxEditorMenu.Group, { label: 'Support' }, () => [
              h(PptxEditorMenu.Row, { label: 'Contact' }),
            ]),
          ]),
        ]),
      editor
    );
    expect(document.querySelector('.pptx-menubar.acme-bar') !== null).toBe(true);
    expect(triggers()).not.toContain('Arrange');
    expect(triggers()[triggers().length - 1]).toBe('Help');
    trigger('Edit').click();
    await nextTick();
    expect(rows(panel())).not.toContain('history.redo');
    expect(panel()!.querySelector('[data-slot="history.undo"]')!.classList).toContain('acme-undo');
    trigger('Edit').click();
    await nextTick();
    trigger('File').click();
    await nextTick();
    expect(rows(panel())).toEqual(['file.pageSetup']);
    [...panel()!.querySelectorAll<HTMLElement>('[role="menuitem"]')]
      .find((r) => r.textContent?.includes('Open…'))!
      .click();
    await nextTick();
    expect(opened).toBe(1);
    expect(panel() === null).toBe(true);
    trigger('Help').click();
    await nextTick();
    expect(panel()!.classList).toContain('acme-help');
    expect(panel()!.querySelector('[role="group"][aria-label="Support"]') !== null).toBe(true);
  });

  test('preset false renders only the host menus, and a named menu keeps its rows', async () => {
    const editor = await deck();
    await mountVue(
      () =>
        h(PptxEditorMenu, { preset: false }, () => [
          h(PptxEditorMenu.Entry, { id: 'mine', label: 'Mine' }, () => [
            h(PptxEditorMenu.Item, { slotId: 'text.bold' }),
          ]),
          h(PptxEditorMenu.Edit, { preset: false }, () => [
            h(PptxEditorMenu.Item, { slotId: 'history.undo' }),
          ]),
        ]),
      editor
    );
    expect(triggers()).toEqual(['Mine', 'Edit']);
    trigger('Edit').click();
    await nextTick();
    expect(rows(panel())).toEqual(['history.undo']);
  });

  test('the bar is one Tab stop; the arrows move and ArrowDown opens', async () => {
    const editor = await deck();
    await mountVue(() => h(PptxEditorMenu), editor);
    const tabbable = () =>
      [...document.querySelectorAll<HTMLElement>('.pptx-menubar__trigger')].filter(
        (b) => b.tabIndex === 0
      );
    expect(tabbable().map((b) => b.textContent)).toEqual(['File']);
    trigger('File').focus();
    trigger('File').dispatchEvent(
      new KeyboardEvent('keydown', { key: 'ArrowRight', bubbles: true })
    );
    await nextTick();
    expect(document.activeElement).toBe(trigger('Edit'));
    expect(tabbable().map((b) => b.textContent)).toEqual(['Edit']);
    trigger('Edit').dispatchEvent(
      new KeyboardEvent('keydown', { key: 'ArrowDown', bubbles: true })
    );
    await nextTick();
    await nextTick();
    expect(trigger('Edit').getAttribute('aria-expanded')).toBe('true');
    // ArrowRight in the open panel moves to the next menu, open.
    panel()!.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowRight', bubbles: true }));
    await nextTick();
    expect(trigger('View').getAttribute('aria-expanded')).toBe('true');
    expect(panel()!.getAttribute('data-menu')).toBe('view');
  });

  test('Escape in a pointer-opened menu focuses its menubar item, not the earlier focus', async () => {
    const editor = await deck();
    await mountVue(() => h(PptxEditorMenu), editor);
    const before = document.createElement('button');
    document.body.append(before);
    before.focus();
    trigger('Format').click();
    await nextTick();
    await nextTick();
    expect(panel() !== null).toBe(true);
    document.activeElement!.dispatchEvent(
      new KeyboardEvent('keydown', { key: 'Escape', bubbles: true })
    );
    await nextTick();
    await nextTick();
    expect(panel() === null).toBe(true);
    expect(document.activeElement).toBe(trigger('Format'));
  });

  test('misused menu parts warn', async () => {
    resetDevWarnings();
    const warn = spyOn(console, 'warn').mockImplementation(() => {});
    try {
      await mountVue(() => [
        h(PptxEditorMenu.Entry, { id: 'x' }),
        h(PptxEditorContextMenu),
        h(PptxEditorMenu, null, () => [
          h(PptxEditorMenu.Edit, null, () => [
            h(PptxEditorMenu.Item, { slotId: 'text.bold', hidden: true }),
          ]),
        ]),
      ]);
      trigger('Edit').click();
      await nextTick();
      const said = warn.mock.calls.map((c) => String(c[0])).join('\n');
      expect(said).toContain('PptxEditorMenu.Entry is outside PptxEditorMenu');
      expect(said).toContain('PptxEditorContextMenu is outside PptxEditorRoot');
      expect(said).toContain('or render it from the popups.contextMenu entry');
      expect(said).toContain(
        'PptxEditorMenu.Edit > PptxEditorMenu.Item slotId="text.bold" hidden matches no row'
      );
    } finally {
      warn.mockRestore();
    }
  });
});

describe('ContextMenu', () => {
  test('matches React, flips inside the window, and Escape closes it', async () => {
    const rowsOf = (R: boolean) =>
      R
        ? [
            createElement(RMenuRow, { key: 'a', label: 'Cut', icon: 'content_cut' }),
            createElement(RMenuRow, { key: 'b', label: 'Copy', disabled: true, reason: 'Nothing' }),
          ]
        : [
            h(MenuRow, { label: 'Cut', icon: 'content_cut' }),
            h(MenuRow, { label: 'Copy', disabled: true, reason: 'Nothing' }),
          ];
    mountReact(
      createElement(RContextMenu, {
        at: { x: 20, y: 30 },
        onClose: () => {},
        label: 'Actions',
        className: 'acme',
        children: rowsOf(true) as ReactElement[],
      }),
      null
    );
    const react = canon(document.querySelector('.pptx-contextmenu'));
    cleanup();
    let closed = 0;
    await mountVue(
      () =>
        h(
          ContextMenu,
          { at: { x: 20, y: 30 }, onClose: () => (closed += 1), label: 'Actions', class: 'acme' },
          () => rowsOf(false)
        ),
      null
    );
    const menu = document.querySelector<HTMLElement>('.pptx-contextmenu')!;
    expect(canon(menu)).toBe(react);
    expect(menu.style.left).toBe('20px');
    menu.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
    expect(closed).toBeGreaterThan(0);
  });
});

/** Mount a Vue root with a canvas and load a deck into it. */
async function vueCanvas(popups?: PptxEditorPopups, extra?: () => unknown) {
  let editor: PptxEditorInstance | null = null;
  await mountVue(() =>
    h(
      PptxEditorRoot,
      { ...(popups ? { popups } : {}), onReady: (e: PptxEditorInstance) => (editor = e) },
      () => [h(PptxEditorContent), extra?.()]
    )
  );
  const r = await editor!.load(
    buildPptx({ slides: [{ shapes: shape({ id: 2, x: IN, y: IN, cx: 2 * IN, cy: IN }) }] })
  );
  expect(r.ok).toBe(true);
  await nextTick();
  const stage = document.querySelector('.pptx-surface__stage') as HTMLElement;
  stage.getBoundingClientRect = () =>
    ({
      left: 0,
      top: 0,
      right: 1280,
      bottom: 720,
      width: 1280,
      height: 720,
      x: 0,
      y: 0,
    }) as DOMRect;
  return { editor: editor!, stage };
}

/** Right-click at (x, y) inches. Returns true when a handler took the event (default prevented). */
async function rightClick(stage: HTMLElement, x: number, y: number): Promise<boolean> {
  const taken = !stage.dispatchEvent(
    new MouseEvent('contextmenu', {
      bubbles: true,
      cancelable: true,
      clientX: (x * IN) / PX,
      clientY: (y * IN) / PX,
    })
  );
  await nextTick();
  await nextTick();
  return taken;
}

const menuSlots = () =>
  [...document.querySelectorAll('.pptx-contextmenu [data-slot]')].map((r) =>
    r.getAttribute('data-slot')
  );

describe('selection marks on any slide background', () => {
  test('the Vue canvas draws a wider halo under the selection outline', async () => {
    const { editor } = await vueCanvas();
    const nodeId = editor.session!.layout(0)!.items[0]!.element.nodeId;
    editor.setSelection({ kind: 'shapes', nodeIds: [nodeId] });
    await nextTick();
    const outline = document.querySelector('.pptx-surface__overlay .pptx-surface__outline')!;
    const halo = outline.previousElementSibling!;
    expect(halo.getAttribute('class')).toBe('pptx-surface__halo pptx-surface__halo--outline');
    expect(Number(halo.getAttribute('stroke-width'))).toBe(
      3 * Number(outline.getAttribute('stroke-width'))
    );
  });

  test('the caret and the marquee get a halo too; handles do not', () => {
    const ns = 'http://www.w3.org/2000/svg';
    const mark = (cls: string, width: string) => {
      const el = document.createElementNS(ns, cls.includes('caret') ? 'line' : 'rect');
      el.setAttribute('class', cls);
      el.setAttribute('stroke-width', width);
      return el;
    };
    const parts = withHalos([
      mark('pptx-surface__caret', '2'),
      mark('pptx-surface__marquee', '1'),
      mark('pptx-surface__handle pptx-surface__handle--nw', '1'),
    ]);
    expect(parts.map((p) => p.getAttribute('class'))).toEqual([
      'pptx-surface__halo pptx-surface__halo--caret',
      'pptx-surface__caret',
      'pptx-surface__halo pptx-surface__halo--marquee',
      'pptx-surface__marquee',
      'pptx-surface__handle pptx-surface__handle--nw',
    ]);
    expect(parts[0]!.getAttribute('stroke-width')).toBe('5');
  });
});

describe('context menu dismissal', () => {
  test('a pointer press outside, a wheel outside, a blur, and a resize each close it', async () => {
    const outside = document.createElement('div');
    document.body.append(outside);
    const closes: string[] = [];
    const ways: [string, () => void][] = [
      // A press the slide surface prevents fires `pointerdown` and no `mousedown`.
      [
        'pointerdown',
        () => void outside.dispatchEvent(new Event('pointerdown', { bubbles: true })),
      ],
      ['wheel', () => void outside.dispatchEvent(new Event('wheel', { bubbles: true }))],
      ['blur', () => void window.dispatchEvent(new Event('blur'))],
      ['resize', () => void window.dispatchEvent(new Event('resize'))],
    ];
    for (const [name, run] of ways) {
      await mountVue(
        () =>
          h(
            ContextMenu,
            { at: { x: 0, y: 0 }, onClose: () => closes.push(name), label: 'Actions' },
            () => [h(MenuRow, { label: 'One', onSelect: () => undefined })]
          ),
        null
      );
      // A press or a wheel inside the menu keeps it.
      const menu = document.querySelector<HTMLElement>('[role="menu"]')!;
      menu.dispatchEvent(new Event('pointerdown', { bubbles: true }));
      menu.dispatchEvent(new Event('wheel', { bubbles: true }));
      expect(closes.length).toBe(ways.findIndex(([n]) => n === name));
      run();
      expect(closes.at(-1)).toBe(name);
      unmountVue();
    }
    outside.remove();
  });

  test('on the canvas: a second right press leaves one menu, and a left press closes it', async () => {
    const { stage } = await vueCanvas();
    await rightClick(stage, 1.5, 1.5);
    stage.dispatchEvent(new Event('pointerdown', { bubbles: true }));
    await nextTick();
    await rightClick(stage, 5, 5);
    expect(document.querySelectorAll('.pptx-contextmenu').length).toBe(1);
    stage.dispatchEvent(new Event('pointerdown', { bubbles: true }));
    await nextTick();
    await nextTick();
    expect(document.querySelectorAll('.pptx-contextmenu').length).toBe(0);
  });
});

describe('PptxEditorContextMenu and popups.contextMenu', () => {
  test('right-click on a shape opens the shape menu; Duplicate runs and closes it', async () => {
    const { editor, stage } = await vueCanvas();
    await rightClick(stage, 1.5, 1.5);
    const menu = document.querySelector('[role="menu"]')!;
    expect(menu.getAttribute('aria-label')).toBe('Shape actions');
    const duplicate = menu.querySelector<HTMLElement>('[data-slot="more.duplicate"]')!;
    expect(duplicate.getAttribute('aria-disabled')).toBeNull();
    // Without the pro comments pane, the menu has no New comment row.
    expect(menu.querySelector('[data-slot="review.addComment"]')).toBeNull();
    duplicate.click();
    await nextTick();
    expect(document.querySelector('.pptx-contextmenu') === null).toBe(true);
    expect(editor.session!.layout(0)!.items.length).toBe(2);
  });

  test('empty slide area offers Paste, Format Background, and New slide', async () => {
    const { editor, stage } = await vueCanvas();
    await rightClick(stage, 10, 6);
    expect(menuSlots()).toEqual(['edit.paste', 'slide.background', 'slide.add']);
    document.querySelector<HTMLElement>('.pptx-contextmenu [data-slot="slide.add"]')!.click();
    expect(editor.snapshot().slides.length).toBe(2);
  });

  test('Shift+F10 opens the menu from the keyboard with focus on the first row', async () => {
    const { stage } = await vueCanvas();
    const canvas = document.querySelector<HTMLElement>('.pptx-canvas')!;
    void stage;
    canvas.dispatchEvent(
      new KeyboardEvent('keydown', { key: 'F10', shiftKey: true, bubbles: true })
    );
    await nextTick();
    await nextTick();
    const menu = document.querySelector('.pptx-contextmenu')!;
    expect(menu.getAttribute('aria-label')).toBe('Slide actions');
    expect(document.activeElement).toBe(menu.querySelector('[role="menuitem"]'));
  });

  test('a renderer changes rows in place and adds host rows', async () => {
    let asked = '';
    const AskRow = defineComponent({
      setup() {
        const target = useContextMenuRequest();
        const element = useContextMenuTarget();
        return () =>
          h(PptxEditorContextMenu.Item, {
            label: 'Ask legal',
            onSelect: () => (asked = element === target?.owner ? (target?.kind ?? '') : 'wrong'),
          });
      },
    });
    const { stage } = await vueCanvas({
      contextMenu: (props) =>
        h(PptxEditorContextMenu, { ...props, class: 'acme-menu' }, () => [
          h(PptxEditorContextMenu.Slot, { slotId: 'format.painter', hidden: true }),
          h(AskRow),
        ]),
    });
    await rightClick(stage, 1.5, 1.5);
    const menu = document.querySelector<HTMLElement>('.pptx-contextmenu.acme-menu')!;
    expect(menu !== null).toBe(true);
    expect(menu.querySelector('[data-slot="format.painter"]') === null).toBe(true);
    expect(menu.querySelector('[data-slot="arrange.order.front"]') !== null).toBe(true);
    [...menu.querySelectorAll<HTMLElement>('[role="menuitem"]')]
      .find((r) => r.textContent?.includes('Ask legal'))!
      .click();
    await nextTick();
    expect(asked).toBe('shapes');
    expect(document.querySelector('.pptx-contextmenu') === null).toBe(true);
  });

  test('false shows no menu, and a hidden New slide leaves no second row', async () => {
    const off = await vueCanvas({ contextMenu: false });
    // The canvas took the right-click, as with a menu, and opened nothing.
    expect(await rightClick(off.stage, 1.5, 1.5)).toBe(true);
    expect(document.querySelectorAll('.pptx-contextmenu').length).toBe(0);
    const { stage } = await vueCanvas({
      contextMenu: (props) =>
        h(PptxEditorContextMenu, props, () => [
          h(PptxEditorContextMenu.Slot, { slotId: 'slide.add', hidden: true }),
        ]),
    });
    // The same flow with a renderer opens the menu.
    expect(await rightClick(stage, 10, 6)).toBe(true);
    expect(menuSlots()).toEqual(['edit.paste', 'slide.background']);
  });
});

describe('a PptxEditorContextMenu placed in the root', () => {
  // The React twin is packages/react/src/__tests__/menu-parts.test.tsx.
  test('it is the canvas menu, with its rows, and openChange hears each open and close', async () => {
    const opens: boolean[] = [];
    let asked = 0;
    const warn = spyOn(console, 'warn').mockImplementation(() => {});
    try {
      const { stage } = await vueCanvas(undefined, () =>
        h(
          PptxEditorContextMenu,
          { className: 'placed', onOpenChange: (open: boolean) => opens.push(open) },
          () => [
            h(PptxEditorContextMenu.Slot, { slotId: 'format.painter', hidden: true }),
            h(PptxEditorContextMenu.Item, { label: 'Ask legal', onSelect: () => (asked += 1) }),
          ]
        )
      );
      expect(document.querySelector('.pptx-contextmenu') === null).toBe(true);
      await rightClick(stage, 1.5, 1.5);
      const menu = document.querySelector<HTMLElement>('.pptx-contextmenu.placed')!;
      expect(menu !== null).toBe(true);
      expect(menu.querySelector('[data-slot="format.painter"]') === null).toBe(true);
      expect(menu.querySelector('[data-slot="arrange.order.front"]') !== null).toBe(true);
      expect(opens).toEqual([true]);
      [...menu.querySelectorAll<HTMLElement>('[role="menuitem"]')]
        .find((r) => r.textContent?.includes('Ask legal'))!
        .click();
      await nextTick();
      expect(asked).toBe(1);
      expect(document.querySelector('.pptx-contextmenu') === null).toBe(true);
      expect(opens).toEqual([true, false]);
      const said = warn.mock.calls.map((c) => String(c[0])).join('\n');
      expect(said).not.toContain('PptxEditorContextMenu is outside');
    } finally {
      warn.mockRestore();
    }
  });

  test('disabled shows no menu, and turning it on closes an open menu', async () => {
    const opens: boolean[] = [];
    const disabled = shallowRef(false);
    const { stage } = await vueCanvas(undefined, () =>
      h(PptxEditorContextMenu, {
        disabled: disabled.value,
        onOpenChange: (open: boolean) => opens.push(open),
      })
    );
    const count = () => document.querySelectorAll('.pptx-contextmenu').length;
    await rightClick(stage, 1.5, 1.5);
    expect(count()).toBe(1);
    disabled.value = true;
    await nextTick();
    await nextTick();
    expect(count()).toBe(0);
    expect(opens).toEqual([true, false]);
    await rightClick(stage, 1.5, 1.5);
    expect(count()).toBe(0);
    disabled.value = false;
    await nextTick();
    expect(count()).toBe(0);
    await rightClick(stage, 1.5, 1.5);
    expect(count()).toBe(1);
  });

  test('an explicit popups.contextMenu entry wins over a placed menu', async () => {
    const placedMenu = () => h(PptxEditorContextMenu, { className: 'placed' });
    const { stage } = await vueCanvas(
      { contextMenu: (props) => h(PptxEditorContextMenu, { ...props, className: 'explicit' }) },
      placedMenu
    );
    await rightClick(stage, 1.5, 1.5);
    expect(document.querySelector('.pptx-contextmenu.explicit') !== null).toBe(true);
    expect(document.querySelector('.pptx-contextmenu.placed') === null).toBe(true);
    const off = await vueCanvas({ contextMenu: false }, placedMenu);
    await rightClick(off.stage, 1.5, 1.5);
    expect(document.querySelector('.pptx-contextmenu') === null).toBe(true);
  });
});

describe('RailContextMenu and SectionContextMenu', () => {
  test('the slide menu matches React and acts on the selected slides', async () => {
    const editor = await deck();
    editor.selectSlide(0);
    editor.selectSlide(1, 'toggle');
    const menu = { at: { x: 20, y: 30 }, index: 1, keyboard: false, owner: null };
    mountReact(
      createElement(RRailContextMenu, { menu, onClose: () => {}, onBackground: () => {} }),
      editor
    );
    const react = canon(document.querySelector('.pptx-contextmenu'));
    cleanup();
    let closed = 0;
    await mountVue(
      () => h(RailContextMenu, { menu, onClose: () => (closed += 1), onBackground: () => {} }),
      editor
    );
    const el = document.querySelector<HTMLElement>('.pptx-contextmenu')!;
    expect(canon(el)).toBe(react);
    el.querySelector<HTMLElement>('[data-slot="slide.duplicate"]')!.click();
    expect(closed).toBe(1);
    expect(editor.snapshot().slides.length).toBe(5);
  });

  test('the section menu matches React; Remove all removes every section', async () => {
    const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    expect((await editor.load(buildSlideLayoutsDeck())).ok).toBe(true);
    const section = editor.snapshot().sections![0]!;
    const menu = { at: { x: 5, y: 5 }, keyboard: true, owner: null };
    const noop = (): void => {};
    mountReact(
      createElement(RSectionContextMenu, {
        section,
        menu,
        onClose: noop,
        onRename: noop,
        onRemoved: noop,
      }),
      editor
    );
    const react = canon(document.querySelector('.pptx-contextmenu'));
    cleanup();
    let removed = 0;
    await mountVue(
      () =>
        h(SectionContextMenu, {
          section,
          menu,
          onClose: noop,
          onRename: noop,
          onRemoved: () => (removed += 1),
        }),
      editor
    );
    const el = document.querySelector<HTMLElement>('.pptx-contextmenu')!;
    expect(canon(el)).toBe(react);
    // A keyboard open puts focus on the first row.
    expect(document.activeElement).toBe(el.querySelector('[role="menuitem"]'));
    el.querySelector<HTMLElement>('[data-slot="section.removeAll"]')!.click();
    expect(editor.snapshot().sections ?? []).toEqual([]);
    expect(removed).toBe(1);
    unmountVue();
  });
});
