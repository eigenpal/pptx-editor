// The chrome review fixes, on both adapters over one deck: the custom color of every color
// menu, the current color and line spacing marked in their menus, the font size box that says
// why it refuses a size, the link popup's announced error, the dialogs' Close button and first
// field, and a menu row's description. Each scenario runs on the React and the Vue part, and
// the two open color menus have the same canonical DOM.

import { afterEach, describe, expect, test } from 'bun:test';
import { createElement } from 'react';
import { act, cleanup } from '@testing-library/react';
import { h, nextTick, type App } from 'vue';
import type { PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import { serializeXmlPart, xmlPartOf } from '@pptx-editor.dev/core/store';
import {
  BackgroundDialog as RBackgroundDialog,
  HyperlinkPopup as RHyperlinkPopup,
  Toolbar,
} from '../../react/src/index.ts';
import { MenuItem as RMenuItem } from '../../react/src/chrome.tsx';
import { BackgroundDialog, HyperlinkPopup, MenuItem, PptxEditorToolbar } from '../src/index.ts';
import {
  canonical,
  fixWidths,
  mountVue,
  openDeck,
  renderReact,
  selectShape,
  selectText,
} from './toolbar-dom.ts';

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
  for (const el of document.querySelectorAll('.pptx-menu, dialog')) el.remove();
});

type Adapter = 'react' | 'vue';
const ADAPTERS: readonly Adapter[] = ['react', 'vue'];

/** Mount the React or the Vue part over `editor`; the returned function settles a change. */
async function mount(
  adapter: Adapter,
  editor: PptxEditorInstance,
  react: () => ReturnType<typeof createElement>,
  vue: () => unknown
): Promise<{ root: HTMLElement; settle: () => Promise<void> }> {
  if (adapter === 'react') {
    let root!: HTMLElement;
    act(() => void (root = renderReact(editor, react())));
    return { root, settle: async () => void act(() => undefined) };
  }
  const { host, app } = await mountVue(editor, vue);
  apps.push([app, host]);
  return {
    root: host,
    settle: async () => {
      await nextTick();
      await nextTick();
    },
  };
}

/** Run `fn` inside React's `act` (a no-op wrapper for Vue). */
const run = (fn: () => void): void => void act(fn);

const slideXml = (e: PptxEditorInstance): string =>
  serializeXmlPart(xmlPartOf(e.session!.store.package, e.session!.info.slides[0]!.partName)!);

/** Type into an input as a user does: set the value and send `input`. */
function typeInto(input: HTMLInputElement, text: string): void {
  const set = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')!.set!;
  run(() => {
    set.call(input, text);
    input.dispatchEvent(new Event('input', { bubbles: true }));
  });
}

const toolbar = (adapter: Adapter, editor: PptxEditorInstance) =>
  mount(
    adapter,
    editor,
    () => createElement(Toolbar),
    () => h(PptxEditorToolbar)
  );

/** The arrow of a split color button, by its name. */
const colorArrow = (root: Element, name: string): HTMLElement =>
  [...root.querySelectorAll<HTMLElement>('[aria-haspopup="menu"]')].find(
    (b) => b.getAttribute('aria-label') === name
  )!;

describe('the custom color of a color menu', () => {
  for (const adapter of ADAPTERS)
    test(`${adapter}: a hex color applies on Enter and on Apply, and the menu marks the current color`, async () => {
      restore = fixWidths(() => 100_000);
      const editor = await openDeck();
      selectShape(editor);
      const { root, settle } = await toolbar(adapter, editor);
      // Fill color: the hex field and Enter.
      run(() => colorArrow(root, 'Fill color').click());
      await settle();
      let menu = document.querySelector<HTMLElement>('.pptx-menu')!;
      const hex = menu.querySelector<HTMLInputElement>('.pptx-custom-color__hex')!;
      expect(hex.getAttribute('aria-label')).toBe('Custom color, hex code');
      expect(menu.querySelector('input[type="color"]') !== null).toBe(true);
      typeInto(hex, '#33g66cc');
      await settle();
      // Only hex digits stay, in upper case.
      expect(hex.value).toBe('3366CC');
      run(() => hex.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true })));
      await settle();
      expect(slideXml(editor)).toContain('3366CC');
      expect(document.querySelector('.pptx-menu') === null).toBe(true);
      // Reopened: the current fill is the custom color, so the field starts with it.
      run(() => colorArrow(root, 'Fill color').click());
      await settle();
      menu = document.querySelector<HTMLElement>('.pptx-menu')!;
      expect(menu.querySelector<HTMLInputElement>('.pptx-custom-color__hex')!.value).toBe('3366CC');
      run(() => colorArrow(root, 'Fill color').click());
      await settle();
      // A standard color is marked as current in the menu (`aria-current`).
      run(() => void editor.exec({ type: 'setShapeFill', color: '#C00000' }));
      await settle();
      run(() => colorArrow(root, 'Fill color').click());
      await settle();
      menu = document.querySelector<HTMLElement>('.pptx-menu')!;
      const current = [...menu.querySelectorAll('[aria-current="true"]')];
      expect(current.map((e) => e.getAttribute('aria-label'))).toEqual(['Dark Red']);
      // The hex field and the selector show the current color, not a fixed example.
      expect(menu.querySelector('.pptx-custom-color__hex')!.getAttribute('placeholder')).toBe(
        'C00000'
      );
      expect(menu.querySelector<HTMLInputElement>('input[type="color"]')!.value).toBe('#c00000');
      // Apply is a row: the arrow keys reach the fields and Apply.
      const apply = menu.querySelector<HTMLButtonElement>('.pptx-custom-color__apply')!;
      expect(apply.getAttribute('role')).toBe('menuitem');
      typeInto(menu.querySelector<HTMLInputElement>('.pptx-custom-color__hex')!, '12ab34');
      await settle();
      run(() => apply.click());
      await settle();
      expect(slideXml(editor)).toContain('12AB34');
    });

  test('the open color menus of React and Vue have the same DOM', async () => {
    restore = fixWidths(() => 100_000);
    const editor = await openDeck();
    selectShape(editor);
    const react = await toolbar('react', editor);
    run(() => colorArrow(react.root, 'Fill color').click());
    const reactMenu = canonical(document.querySelector('.pptx-menu')!);
    run(() => colorArrow(react.root, 'Fill color').click());
    cleanup();
    for (const el of document.querySelectorAll('.pptx-menu')) el.remove();
    const vue = await toolbar('vue', editor);
    colorArrow(vue.root, 'Fill color').click();
    await vue.settle();
    expect(canonical(document.querySelector('.pptx-menu')!)).toBe(reactMenu);
  });
});

describe('the Line spacing menu', () => {
  for (const adapter of ADAPTERS)
    test(`${adapter}: the current spacing is the checked row`, async () => {
      restore = fixWidths(() => 100_000);
      const editor = await openDeck();
      selectShape(editor);
      run(() => void editor.exec({ type: 'setLineSpacing', multiple: 1.5 }));
      // Line spacing is on the text bar.
      selectText(editor);
      const { root, settle } = await toolbar(adapter, editor);
      run(() => root.querySelector<HTMLElement>('[data-slot="list.lineSpacing"]')!.click());
      await settle();
      const rows = [...document.querySelectorAll('.pptx-menu [role="menuitemradio"]')];
      expect(rows.length).toBe(4);
      const checked = rows.filter((r) => r.getAttribute('aria-checked') === 'true');
      expect(checked.map((r) => r.textContent)).toEqual(['1.5']);
    });
});

describe('the font size box', () => {
  for (const adapter of ADAPTERS)
    test(`${adapter}: a size out of range is refused with the range, announced`, async () => {
      restore = fixWidths(() => 100_000);
      const editor = await openDeck();
      selectShape(editor);
      const before = slideXml(editor);
      const { root, settle } = await toolbar(adapter, editor);
      const box = root.querySelector<HTMLInputElement>('.pptx-toolbar__size')!;
      run(() => box.focus());
      typeInto(box, '5000');
      run(() => box.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true })));
      await settle();
      const alert = root.querySelector<HTMLElement>('.pptx-toolbar__size-error')!;
      expect(alert.getAttribute('role')).toBe('alert');
      expect(alert.textContent).toBe('Font size must be a number from 1 to 4000 pt.');
      expect(box.getAttribute('aria-invalid')).toBe('true');
      expect(box.getAttribute('aria-describedby')).toBe(alert.id);
      expect(slideXml(editor)).toBe(before);
      // A new entry clears the message.
      typeInto(box, '2');
      await settle();
      expect(root.querySelectorAll('.pptx-toolbar__size-error').length).toBe(0);
    });
});

describe('the link popup', () => {
  for (const adapter of ADAPTERS)
    test(`${adapter}: a refused link is announced, and the field names it`, async () => {
      const editor = await openDeck();
      selectShape(editor);
      const anchor = document.body.appendChild(document.createElement('button'));
      const props = { anchor, url: '', onClose: () => undefined };
      const { settle } = await mount(
        adapter,
        editor,
        () => createElement(RHyperlinkPopup, props),
        () => h(HyperlinkPopup, props)
      );
      const input = document.querySelector<HTMLInputElement>('.pptx-link-popup input')!;
      expect(input.hasAttribute('aria-invalid')).toBe(false);
      typeInto(input, 'javascript:alert(1)');
      run(() => input.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true })));
      await settle();
      const error = document.querySelector<HTMLElement>('.pptx-link__error')!;
      expect(error.getAttribute('role')).toBe('alert');
      expect(input.getAttribute('aria-invalid')).toBe('true');
      expect(input.getAttribute('aria-describedby')).toBe(error.id);
      anchor.remove();
    });
});

describe('Format Background', () => {
  for (const adapter of ADAPTERS)
    test(`${adapter}: a Close button in the header, and the first field takes focus`, async () => {
      const editor = await openDeck();
      const props = { open: true, onClose: () => undefined };
      const { settle } = await mount(
        adapter,
        editor,
        () => createElement(RBackgroundDialog, props),
        () => h(BackgroundDialog, props)
      );
      await settle();
      const dialog = document.querySelector<HTMLElement>('[data-pptx-dialog="background"]')!;
      const close = dialog.querySelector<HTMLElement>('[data-pptx-part="header"] button')!;
      expect(close.getAttribute('aria-label')).toBe('Close');
      // The first field is the color grid: its swatch, not OK and not Close.
      expect(document.activeElement?.classList.contains('pptx-swatch')).toBe(true);
    });
});

describe('a menu row with a description', () => {
  for (const adapter of ADAPTERS)
    test(`${adapter}: the name is the label, the hint is its description`, async () => {
      const editor = await openDeck();
      const props = { description: 'Download the edited presentation', onSelect: () => undefined };
      const { root } = await mount(
        adapter,
        editor,
        () => createElement(RMenuItem, { ...props, children: 'Export' }),
        () => h(MenuItem, props, () => 'Export')
      );
      const row = root.querySelector('[role="menuitem"]')!;
      expect(row.getAttribute('aria-description')).toBe('Download the edited presentation');
      expect(row.textContent).toBe('Export');
    });
});
