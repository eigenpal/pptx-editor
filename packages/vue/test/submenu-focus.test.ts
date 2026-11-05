// ArrowRight on a submenu row moves focus to the panel's first row in both adapters, and only
// once the panel is placed and visible: a panel with `visibility: hidden` cannot take focus in a
// browser, so a focus call while it is hidden leaves focus on the parent row.

import { afterEach, describe, expect, spyOn, test } from 'bun:test';
import { createElement } from 'react';
import { act, cleanup, render } from '@testing-library/react';
import { createApp, h, nextTick, type App } from 'vue';
import { MenuRow as RMenuRow, MenuSubmenu as RMenuSubmenu } from '../../react/src/ContextMenu.tsx';
import { MenuRow, MenuSubmenu } from '../src/index.ts';

let app: App | null = null;
afterEach(() => {
  cleanup();
  app?.unmount();
  app = null;
  document.body.replaceChildren();
});

interface FocusRecord {
  readonly label: string;
  readonly hiddenPanel: boolean;
}

/** Record each focus call: the element's text, and whether its panel was hidden then. */
function recordFocus(): { calls: FocusRecord[]; restore: () => void } {
  const calls: FocusRecord[] = [];
  const spy = spyOn(HTMLElement.prototype, 'focus').mockImplementation(function (
    this: HTMLElement
  ) {
    const panel = this.closest<HTMLElement>('.pptx-menu--submenu');
    calls.push({
      label: this.textContent?.trim() ?? '',
      hiddenPanel: panel?.style.visibility === 'hidden',
    });
  });
  return { calls, restore: () => spy.mockRestore() };
}

const pressArrowRight = (row: HTMLElement): void => {
  row.dispatchEvent(new KeyboardEvent('keydown', { key: 'ArrowRight', bubbles: true }));
};

const submenuRow = (): HTMLElement => {
  const row = document.querySelector<HTMLElement>('[aria-haspopup="menu"]');
  if (!row) throw new Error('no submenu row');
  return row;
};

describe('submenu keyboard focus', () => {
  test('React and Vue focus the first submenu row after the panel shows', async () => {
    const react = recordFocus();
    try {
      render(
        createElement(
          'div',
          { className: 'pptx-menu', role: 'menu' },
          createElement(RMenuSubmenu, {
            label: 'Find',
            children: [
              createElement(RMenuRow, { key: 'next', label: 'Find next' }),
              createElement(RMenuRow, { key: 'replace', label: 'Replace' }),
            ],
          })
        )
      );
      await act(async () => {
        pressArrowRight(submenuRow());
      });
    } finally {
      react.restore();
    }
    cleanup();

    const vue = recordFocus();
    try {
      const host = document.createElement('div');
      document.body.append(host);
      app = createApp({
        render: () =>
          h('div', { class: 'pptx-menu', role: 'menu' }, [
            h(MenuSubmenu, { label: 'Find' }, () => [
              h(MenuRow, { label: 'Find next' }),
              h(MenuRow, { label: 'Replace' }),
            ]),
          ]),
      });
      app.mount(host);
      await nextTick();
      pressArrowRight(submenuRow());
      for (let i = 0; i < 4; i++) await nextTick();
    } finally {
      vue.restore();
    }

    expect(react.calls).toEqual([{ label: 'Find next', hiddenPanel: false }]);
    expect(vue.calls).toEqual(react.calls);
  });
});
