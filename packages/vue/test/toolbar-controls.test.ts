// The Vue toolbar's controls against the React twin: the same canonical DOM for every
// control on one deck (pickers, steppers, value menus, color splits, galleries, and the
// adapter actions), then the behavior of each control kind, the More menu, and the
// customization ladder (icon, asChild, class). After DOCX Editor's cross-adapter test.

import { afterEach, describe, expect, spyOn, test } from 'bun:test';
import { createElement, type FunctionComponent } from 'react';
import { act, cleanup } from '@testing-library/react';
import { h, nextTick, type App, type Component } from 'vue';
import { resetDevWarnings, type PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import { PptxEditorRibbon as ReactRibbon, Toolbar } from '../../react/src/index.ts';
import { PptxEditorRibbon, PptxEditorToolbar } from '../src/index.ts';
import {
  canonical,
  fixWidths,
  mountVue,
  openDeck,
  openPictureDeck,
  openTableDeck,
  renderReact,
  selectShape,
  selectText,
} from './toolbar-dom.ts';

const T = PptxEditorToolbar as unknown as Record<string, Component & { pptxSlot?: string }>;

const apps: [App, HTMLElement][] = [];
let restore: (() => void) | null = null;
afterEach(() => {
  cleanup();
  for (const [app, host] of apps.splice(0)) {
    app.unmount();
    host.remove();
  }
  restore?.();
  restore = null;
  for (const el of document.querySelectorAll('.pptx-menu')) el.remove();
});

async function vue(editor: PptxEditorInstance, render: () => unknown): Promise<HTMLElement> {
  const { host, app } = await mountVue(editor, render);
  apps.push([app, host]);
  return host;
}

const bar = (root: Element): Element => root.querySelector('.pptx-toolbar')!;

const markup = (root: Element): string => canonical(bar(root));

describe('the React and Vue toolbars draw the same controls', () => {
  test('with nothing selected', async () => {
    restore = fixWidths(() => 100_000);
    const editor = await openDeck();
    const react = renderReact(editor, createElement(Toolbar));
    const host = await vue(editor, () => h(PptxEditorToolbar));
    expect(markup(host)).toBe(markup(react));
  });

  test('with a text box selected', async () => {
    restore = fixWidths(() => 100_000);
    const editor = await openDeck();
    selectShape(editor);
    const react = renderReact(editor, createElement(Toolbar));
    const host = await vue(editor, () => h(PptxEditorToolbar));
    expect(markup(host)).toBe(markup(react));
  });
});

describe('the React and Vue toolbars draw the same contextual groups and menus', () => {
  test('with a picture selected', async () => {
    restore = fixWidths(() => 100_000);
    const editor = await openPictureDeck();
    const react = renderReact(editor, createElement(Toolbar));
    const host = await vue(editor, () => h(PptxEditorToolbar));
    expect(host.querySelector('[data-group="image"]') !== null).toBe(true);
    expect(markup(host)).toBe(markup(react));
  });

  test('with the caret in a table', async () => {
    restore = fixWidths(() => 100_000);
    const editor = await openTableDeck();
    const react = renderReact(editor, createElement(Toolbar));
    const host = await vue(editor, () => h(PptxEditorToolbar));
    expect(host.querySelector('[data-group="table"]') !== null).toBe(true);
    expect(markup(host)).toBe(markup(react));
  });

  test('on a narrow bar: the same groups leave for More, and the open panel matches', async () => {
    restore = fixWidths(() => 60);
    const editor = await openDeck();
    // A selected shape fills the bar: the empty selection's short bar fits even this width.
    selectShape(editor);
    const react = renderReact(editor, createElement(Toolbar));
    const reactMore = react.querySelector<HTMLElement>('[aria-label="More"]')!;
    expect(reactMore !== null).toBe(true);
    const host = await vue(editor, () => h(PptxEditorToolbar));
    await nextTick();
    expect(markup(host)).toBe(markup(react));
    act(() => reactMore.click());
    const reactPanel = canonical(document.querySelector('.pptx-menu[role="dialog"]')!);
    act(() => reactMore.click());
    expect(document.querySelector('.pptx-menu[role="dialog"]') === null).toBe(true);
    host.querySelector<HTMLElement>('[aria-label="More"]')!.click();
    await nextTick();
    expect(canonical(document.querySelector('.pptx-menu[role="dialog"]')!)).toBe(reactPanel);
  });

  /** Open a control's menu in React, read it, close it, then the same in Vue. */
  async function sameMenu(slot: string, trigger: (el: Element) => HTMLElement): Promise<void> {
    restore = fixWidths(() => 100_000);
    const editor = await openDeck();
    // The caret in the text box: the text bar has every one of these menus.
    selectText(editor);
    const react = renderReact(editor, createElement(Toolbar));
    const reactTrigger = trigger(react.querySelector(`[data-slot="${slot}"]`)!);
    act(() => reactTrigger.click());
    const reactMenu = canonical(document.querySelector('.pptx-menu')!);
    act(() => reactTrigger.click());
    cleanup();
    const host = await vue(editor, () => h(PptxEditorToolbar));
    trigger(host.querySelector(`[data-slot="${slot}"]`)!).click();
    await nextTick();
    await nextTick();
    expect(canonical(document.querySelector('.pptx-menu')!)).toBe(reactMenu);
  }

  test('the font family list', async () => {
    await sameMenu('font.family', (el) => el as HTMLElement);
  });

  test('a value menu with a dialog row (line spacing)', async () => {
    await sameMenu('list.lineSpacing', (el) => el as HTMLElement);
  });

  test('the font color menu', async () => {
    // The arrow half opens the swatches; the main half applies the last color.
    await sameMenu('text.color', (el) => el.parentElement!.querySelector('.pptx-toolbar__step')!);
  });
});

describe('the customization ladder draws the same in both adapters', () => {
  test('class, icon, asChild, hidden, and a host group with priority and pinned', async () => {
    restore = fixWidths(() => 100_000);
    const editor = await openDeck();
    selectShape(editor);
    const R = Toolbar as unknown as Record<string, FunctionComponent<Record<string, unknown>>>;
    const react = renderReact(
      editor,
      createElement(
        Toolbar,
        null,
        createElement(Toolbar.Slot, {
          slotId: 'text.italic',
          icon: createElement('i', { className: 'acme-italic' }),
          className: 'acme-slot',
        }),
        createElement(
          R.Slot!,
          { slotId: 'history.undo', asChild: true },
          createElement('span', { className: 'acme-undo' }, 'Undo')
        ),
        createElement(R.Bold!, null, createElement('b', { className: 'acme-bold' })),
        createElement(R.FontFamily!, { className: 'acme-family' }),
        createElement(R.FontColor!, { icon: createElement('u', { className: 'acme-color' }) }),
        createElement(R.Strike!, { hidden: true }),
        createElement(
          Toolbar.Group,
          { id: 'acme', label: 'Acme', priority: 175, pinned: true, className: 'acme-group' },
          createElement(Toolbar.Action, { label: 'Ask' })
        )
      )
    );
    const host = await vue(editor, () =>
      h(PptxEditorToolbar, null, () => [
        h(T.Slot!, {
          slotId: 'text.italic',
          icon: h('i', { class: 'acme-italic' }),
          class: 'acme-slot',
        }),
        h(T.Slot!, { slotId: 'history.undo', asChild: true }, () => [
          h('span', { class: 'acme-undo' }, 'Undo'),
        ]),
        h(T.Bold!, null, () => [h('b', { class: 'acme-bold' })]),
        h(T.FontFamily!, { class: 'acme-family' }),
        h(T.FontColor!, { icon: h('u', { class: 'acme-color' }) }),
        h(T.Strike!, { hidden: true }),
        h(
          T.Group!,
          { id: 'acme', label: 'Acme', priority: 175, pinned: true, class: 'acme-group' },
          () => [h(PptxEditorToolbar.Action, { label: 'Ask' })]
        ),
      ])
    );
    expect(
      host.querySelector('span.acme-undo[role="button"][data-slot="history.undo"]') !== null
    ).toBe(true);
    expect(host.querySelector('[data-slot="text.italic"].acme-slot .acme-italic') !== null).toBe(
      true
    );
    expect(host.querySelector('[data-slot="text.bold"] .acme-bold') !== null).toBe(true);
    expect(host.querySelector('[data-slot="font.family"].acme-family') !== null).toBe(true);
    expect(host.querySelector('[data-slot="text.color"] .acme-color') !== null).toBe(true);
    expect(host.querySelector('[data-slot="text.strike"]') === null).toBe(true);
    expect(markup(host)).toBe(markup(react));
  });
});

/** Mount the preset Vue toolbar over a deck with its text box selected. */
async function selectedBar(
  width = 100_000
): Promise<{ host: HTMLElement; editor: PptxEditorInstance }> {
  restore = fixWidths(() => width);
  const editor = await openDeck();
  selectShape(editor);
  const host = await vue(editor, () => h(PptxEditorToolbar));
  return { host, editor };
}

const flush = async (): Promise<void> => {
  await nextTick();
  await nextTick();
};

const undoOn = (host: Element): boolean =>
  host.querySelector('[data-slot="history.undo"]')!.getAttribute('aria-disabled') !== 'true';

function type(el: HTMLInputElement, value: string): void {
  el.value = value;
  el.dispatchEvent(new Event('input', { bubbles: true }));
}

const key = (el: Element, k: string): void =>
  void el.dispatchEvent(new KeyboardEvent('keydown', { key: k, bubbles: true, cancelable: true }));

describe('Vue toolbar controls run their commands', () => {
  test('the font family list filters and applies the first match on Enter', async () => {
    const { host, editor } = await selectedBar();
    const families = editor.fontFamilies();
    expect(families.length).toBeGreaterThan(1);
    const target = families[families.length - 1]!;
    const trigger = host.querySelector<HTMLElement>('[data-slot="font.family"]')!;
    trigger.click();
    await flush();
    const search = document.querySelector<HTMLInputElement>('.pptx-font-picker__search')!;
    type(search, target);
    await flush();
    const rows = [...document.querySelectorAll('.pptx-font-picker__list [role="option"]')];
    expect(rows.map((r) => r.textContent)).toContain(target);
    expect(rows.length).toBeLessThan(families.length);
    key(search, 'Enter');
    await flush();
    expect(document.querySelector('.pptx-font-picker') === null).toBe(true);
    expect(trigger.querySelector('.pptx-toolbar__family-label')!.textContent).toBe(
      rows[0]!.textContent
    );
  });

  test('the size box applies a typed size on Enter, and the steps change it', async () => {
    const { host } = await selectedBar();
    const box = host.querySelector<HTMLInputElement>('.pptx-toolbar__size')!;
    type(box, '20');
    key(box, 'Enter');
    await flush();
    expect(box.value).toBe('20');
    box.parentElement!.querySelectorAll<HTMLElement>('.pptx-toolbar__step')[1]!.click();
    await flush();
    expect(Number(box.value)).toBeGreaterThan(20);
    expect(undoOn(host)).toBe(true);
  });

  test('a value menu checks the current choice and applies a pick', async () => {
    const { host } = await selectedBar();
    const trigger = host.querySelector<HTMLElement>('[data-slot="paragraph.align"]')!;
    trigger.click();
    await flush();
    const rows = [...document.querySelectorAll<HTMLElement>('.pptx-menu [role="menuitemradio"]')];
    expect(rows.filter((r) => r.getAttribute('aria-checked') === 'true').length).toBe(1);
    const pick = rows.find((r) => r.getAttribute('aria-checked') !== 'true')!;
    const label = pick.textContent;
    pick.click();
    await flush();
    expect(document.querySelector('.pptx-menu') === null).toBe(true);
    trigger.click();
    await flush();
    const checked = document.querySelector('.pptx-menu [aria-checked="true"]')!;
    expect(checked.textContent).toBe(label);
  });

  test('the color split applies its default color, and the arrow opens the swatches', async () => {
    const { host } = await selectedBar();
    expect(undoOn(host)).toBe(false);
    const main = host.querySelector<HTMLElement>('[data-slot="text.color"]')!;
    // Dark Red from the Standard Colors, core's default for text.
    const swatch = main.querySelector<HTMLElement>('.pptx-toolbar__swatch')!;
    expect(swatch.getAttribute('style')).toContain('#C00000');
    main.click();
    await flush();
    expect(undoOn(host)).toBe(true);
    main.parentElement!.querySelector<HTMLElement>('.pptx-toolbar__step')!.click();
    await flush();
    expect(document.querySelectorAll('.pptx-menu .pptx-swatch').length).toBeGreaterThan(0);
  });

  test('Insert Picture opens a file input that takes pictures only', async () => {
    const { host } = await selectedBar();
    const button = host.querySelector<HTMLElement>('[data-slot="insert.image"]')!;
    const input = button.parentElement!.querySelector<HTMLInputElement>('input[type="file"]')!;
    expect(input.hidden).toBe(true);
    expect(input.accept).toContain('image/');
    const clicked = spyOn(input, 'click').mockImplementation(() => {});
    button.click();
    expect(clicked).toHaveBeenCalledTimes(1);
  });
});

describe('the Vue More menu', () => {
  test('groups that do not fit move into More; ArrowDown opens it at its first control', async () => {
    const { host } = await selectedBar(60);
    const toolbar = bar(host);
    expect(toolbar.getAttribute('data-overflow')).toBe('menu');
    const more = host.querySelector<HTMLElement>('[aria-label="More"]')!;
    expect(more.getAttribute('aria-haspopup')).toBe('dialog');
    more.focus();
    key(more, 'ArrowDown');
    await flush();
    const panel = document.querySelector<HTMLElement>('.pptx-menu[role="dialog"]')!;
    expect(more.getAttribute('aria-controls')).toBe(panel.id);
    expect(
      panel.querySelector('[data-group="more"] .pptx-toolbar__overflow-heading') !== null
    ).toBe(true);
    expect(panel.contains(document.activeElement)).toBe(true);
  });

  test('priority and pinned keep host groups on the bar; overflowContent fills the menu', async () => {
    restore = fixWidths(() => 20);
    const editor = await openDeck();
    const host = await vue(editor, () =>
      h(PptxEditorToolbar, null, () => [
        h(T.Group!, { id: 'pin', label: 'Pin', pinned: true }, () => [
          h(PptxEditorToolbar.Action, { label: 'Pinned' }),
        ]),
        h(
          T.Group!,
          { id: 'find', overflowContent: () => h('span', { 'data-testid': 'extra' }) },
          () => []
        ),
      ])
    );
    const ids = [...bar(host).querySelectorAll(':scope > [data-group]')].map((g) =>
      g.getAttribute('data-group')
    );
    expect(ids).toContain('pin');
    expect(ids).not.toContain('find');
    host.querySelector<HTMLElement>('[aria-label="More"]')!.click();
    await flush();
    const find = document.querySelector('.pptx-toolbar__overflow [data-group="find"]')!;
    // A built-in group keeps its controls before its overflowContent, as in React.
    expect(find.querySelectorAll('button[data-slot]').length).toBeGreaterThan(0);
    expect(find.querySelector('[data-testid="extra"]') !== null).toBe(true);
  });

  test('overflow false wraps the bar, and turning it off brings collapsed groups back', async () => {
    restore = fixWidths(() => 60);
    const editor = await openDeck();
    selectShape(editor);
    const state = { overflow: true };
    const { reactive } = await import('vue');
    const live = reactive(state);
    const host = await vue(editor, () => h(PptxEditorToolbar, { overflow: live.overflow }));
    const moreButtons = () => host.querySelectorAll('button[aria-label="More"]').length;
    expect(moreButtons()).toBe(1);
    live.overflow = false;
    await flush();
    // The built-in More group keeps its label; only the overflow button goes away.
    expect(moreButtons()).toBe(0);
    expect(bar(host).classList.contains('pptx-toolbar--wrap')).toBe(true);
    // `more` is among the first shape groups to leave the bar; it is back now.
    expect(bar(host).querySelector(':scope > [data-group="more"]') !== null).toBe(true);
  });
});

describe('the Vue customization ladder', () => {
  test('asChild: a host span runs the command on Enter and carries the state', async () => {
    restore = fixWidths(() => 100_000);
    const editor = await openDeck();
    selectShape(editor);
    const host = await vue(editor, () =>
      h(PptxEditorToolbar, null, () => [
        h(T.Slot!, { slotId: 'text.bold', asChild: true }, () => [
          h('span', { class: 'acme-bold' }, 'B'),
        ]),
      ])
    );
    const bold = host.querySelector<HTMLElement>('span.acme-bold')!;
    expect(bold.getAttribute('role')).toBe('button');
    expect(bold.getAttribute('aria-pressed')).toBe('false');
    key(bold, 'Enter');
    await flush();
    expect(host.querySelector('span.acme-bold')!.getAttribute('aria-pressed')).toBe('true');
  });

  test('a picker refuses icon and asChild with a warning that names the part', async () => {
    resetDevWarnings();
    const warn = spyOn(console, 'warn').mockImplementation(() => {});
    try {
      restore = fixWidths(() => 100_000);
      const editor = await openDeck();
      selectShape(editor);
      const host = await vue(editor, () =>
        h(PptxEditorToolbar, null, () => [
          h(T.Slot!, { slotId: 'text.highlight', asChild: true }, () => [h('span')]),
          h(T.FontFamily!, { icon: h('i') }),
        ])
      );
      expect(host.querySelector('button[data-slot="text.highlight"]') !== null).toBe(true);
      const text = warn.mock.calls.map((c: unknown[]) => String(c[0])).join('\n');
      expect(text).toContain('"text.highlight" renders a picker');
      expect(text).toContain('PptxEditorToolbar.FontFamily');
    } finally {
      warn.mockRestore();
    }
  });

  test('a dialog control opens its registry dialog under the root', async () => {
    restore = fixWidths(() => 100_000);
    const { PptxEditorRoot } = await import('../src/index.ts');
    const { createApp, defineComponent } = await import('vue');
    const { buildPptx } = await import('../../../scripts/fixtures/builder.ts');
    let editor: PptxEditorInstance | null = null;
    const host = document.createElement('div');
    document.body.append(host);
    const app = createApp(
      defineComponent({
        setup: () => () =>
          h(PptxEditorRoot, { onReady: (e: PptxEditorInstance) => (editor = e) }, () => [
            h(PptxEditorToolbar, { preset: false }, () => [
              h(T.Slot!, { slotId: 'file.pageSetup' }),
            ]),
          ]),
      }) as Component
    );
    app.mount(host);
    apps.push([app, host]);
    await nextTick();
    await editor!.load(buildPptx({ slides: [{ shapes: '' }] }));
    await flush();
    const button = host.querySelector<HTMLElement>('[data-slot="file.pageSetup"]')!;
    expect(button.getAttribute('aria-haspopup')).toBe('dialog');
    button.click();
    await flush();
    expect(document.querySelector('[data-pptx-dialog="pageSetup"]') !== null).toBe(true);
  });
});

describe('the ribbon tab panels', () => {
  const panel = (root: Element): Element =>
    root.querySelector('.pptx-ribbon__panel .pptx-toolbar')!;

  test('the Home panel measures for More and draws the same groups as React', async () => {
    restore = fixWidths(() => 400);
    const editor = await openDeck();
    selectShape(editor);
    const react = renderReact(editor, createElement(ReactRibbon));
    const host = await vue(editor, () => h(PptxEditorRibbon));
    await nextTick();
    const vuePanel = panel(host);
    expect(vuePanel.getAttribute('data-overflow')).toBe('menu');
    expect(vuePanel.classList.contains('pptx-toolbar--wrap')).toBe(false);
    expect(vuePanel.querySelector(':scope > [aria-label="More"]') !== null).toBe(true);
    expect(canonical(vuePanel)).toBe(canonical(panel(react)));
    // The Quick Access Toolbar wraps in both: it never collapses.
    const qat = (root: Element) => root.querySelector('.pptx-ribbon__qat .pptx-toolbar')!;
    expect(canonical(qat(host))).toBe(canonical(qat(react)));
  });

  test('overflow false wraps the panel in both adapters', async () => {
    restore = fixWidths(() => 400);
    const editor = await openDeck();
    const react = renderReact(editor, createElement(ReactRibbon, { overflow: false }));
    const host = await vue(editor, () => h(PptxEditorRibbon, { overflow: false }));
    expect(panel(host).getAttribute('data-overflow')).toBe('wrap');
    expect(canonical(panel(host))).toBe(canonical(panel(react)));
  });
});
