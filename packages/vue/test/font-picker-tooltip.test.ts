// The Vue twins of the React tooltip and font picker tests: the root mounts the tooltip bubble,
// a hover opens it inside the editor scope with a disabled control's reason on its own line, no
// toolbar or status bar control keeps a native `title`; the font picker groups theme fonts, the
// fonts in this presentation, and all fonts, draws each name in its own face, marks the current
// family, scrolls to it, and applies a pick. After DOCX Editor's `Tooltip` and FontFamily tests.

import { afterEach, describe, expect, test } from 'bun:test';
import { createApp, h, nextTick, type App } from 'vue';
import {
  TOOLTIP_DELAY_MS,
  createPptxEditor,
  type PptxEditorInstance,
} from '@pptx-editor.dev/core/editor';
import { serializeXmlPart, xmlPartOf } from '@pptx-editor.dev/core/store';
import { buildPptx, para, run, shape } from '../../../scripts/fixtures/builder.ts';
import {
  PptxEditorRoot,
  PptxEditorStatusBar,
  PptxEditorToolbar,
  PptxEditorViewport,
} from '../src/index.ts';
import { mountVue } from './toolbar-dom.ts';

const IN = 914_400;
const wait = (ms: number) => new Promise<void>((r) => setTimeout(r, ms));
const over = (el: Element) =>
  void el.dispatchEvent(new MouseEvent('pointerover', { bubbles: true }));

const apps: [App, HTMLElement][] = [];
afterEach(() => {
  for (const [app, host] of apps.splice(0)) {
    app.unmount();
    host.remove();
  }
  for (const el of document.querySelectorAll('.pptx-menu')) el.remove();
});

async function mountRoot(colorMode?: 'dark'): Promise<HTMLElement> {
  const host = document.createElement('div');
  document.body.append(host);
  const app = createApp({
    render: () =>
      h(PptxEditorRoot, colorMode ? { colorMode } : {}, {
        default: () => h(PptxEditorViewport, () => [h(PptxEditorToolbar), h(PptxEditorStatusBar)]),
      }),
  });
  app.mount(host);
  apps.push([app, host]);
  await nextTick();
  await nextTick();
  return host;
}

describe('the chrome tooltip (Vue)', () => {
  test('a hover opens the bubble in the editor scope, with the reason on its own line', async () => {
    const host = await mountRoot();
    const undo = host.querySelector<HTMLElement>('[data-slot="history.undo"]')!;
    expect(undo.getAttribute('aria-disabled')).toBe('true');
    over(undo);
    await nextTick();
    expect(document.querySelectorAll('.pptx-tooltip').length).toBe(0);
    await wait(TOOLTIP_DELAY_MS + 40);
    await nextTick();
    const bubble = document.querySelector<HTMLElement>('.pptx-tooltip')!;
    expect(bubble.parentElement).toBe(host.querySelector('.pptx-editor'));
    expect(bubble.getAttribute('role')).toBe('tooltip');
    expect(bubble.getAttribute('aria-hidden')).toBe('true');
    expect(bubble.getAttribute('data-side')).toBe('bottom');
    expect(bubble.querySelector('.pptx-tooltip__label')!.textContent).toBe(
      undo.getAttribute('aria-label')!
    );
    const reason = bubble.querySelector('.pptx-tooltip__reason')!.textContent!;
    expect(reason.length).toBeGreaterThan(0);
    expect(bubble.textContent).not.toContain('—');
    expect(document.getElementById(undo.getAttribute('aria-describedby')!)!.textContent).toBe(
      reason
    );
    over(host.querySelector('.pptx-editor')!);
    await nextTick();
    expect(document.querySelectorAll('.pptx-tooltip').length).toBe(0);
  });

  test('no toolbar or status bar control keeps a native title', async () => {
    const host = await mountRoot();
    expect(host.querySelectorAll('[data-pptx-tooltip]').length).toBeGreaterThan(10);
    expect(host.querySelectorAll('.pptx-toolbar [title], .pptx-statusbar [title]').length).toBe(0);
  });

  test('in dark mode the bubble sits in the dark scope', async () => {
    const host = await mountRoot('dark');
    over(host.querySelector('[data-slot="history.redo"]')!);
    await wait(TOOLTIP_DELAY_MS + 40);
    await nextTick();
    expect(document.querySelector('.pptx-tooltip')!.closest('.pptx-editor.dark') !== null).toBe(
      true
    );
  });
});

async function fontDeck(latin = 'Courier New'): Promise<PptxEditorInstance> {
  const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  const bytes = buildPptx({
    majorFont: 'Georgia',
    minorFont: 'Verdana',
    slides: [
      {
        layout: 'blank',
        shapes: shape({
          id: 2,
          x: IN,
          y: IN,
          cx: 2 * IN,
          cy: IN,
          text: para(run('Hi', '', `<a:latin typeface="${latin}"/>`)),
        }),
      },
    ],
  });
  expect((await editor.load(bytes)).ok).toBe(true);
  const id = editor.session!.layout(0)!.items[0]!.element.nodeId;
  editor.setSelection({ kind: 'shapes', nodeIds: [id] });
  return editor;
}

async function bar(editor: PptxEditorInstance): Promise<HTMLElement> {
  const { host, app } = await mountVue(editor, () => h(PptxEditorToolbar));
  apps.push([app, host]);
  return host;
}

const headings = () =>
  [...document.querySelectorAll('.pptx-font-picker__heading')].map((el) => el.textContent);
const options = () => [
  ...document.querySelectorAll<HTMLElement>('.pptx-font-picker__list [role="option"]'),
];
const family = (row: Element) => row.querySelector('.pptx-font-picker__face')!.textContent;
// The marked rows, as `family:theme`, with their check marks.
const marked = () =>
  options()
    .filter((o) => o.getAttribute('aria-selected') === 'true')
    .filter((o) => o.querySelector('.pptx-menu__check') !== null)
    .map((o) => `${family(o)}:${o.dataset.theme ?? ''}`);
// The first run's `a:latin` typeface, as written in the slide part.
const latinOf = (editor: PptxEditorInstance): string | undefined => {
  const part = xmlPartOf(editor.session!.store.package, editor.session!.info.slides[0]!.partName)!;
  return /<a:latin typeface="([^"]*)"/.exec(serializeXmlPart(part))?.[1];
};

describe('the font family picker (Vue)', () => {
  test('groups theme fonts, the deck families, and all fonts, each in its own face', async () => {
    const editor = await fontDeck();
    const host = await bar(editor);
    const trigger = host.querySelector<HTMLElement>('[data-slot="font.family"]')!;
    expect(trigger.textContent).toContain('Courier New');
    trigger.click();
    await nextTick();
    expect(headings()).toEqual(['Theme fonts', 'Fonts in this presentation', 'All fonts']);
    const [headingsRow, bodyRow] = options();
    expect(family(headingsRow!)).toBe('Georgia');
    expect(headingsRow!.querySelector('.pptx-font-picker__role')!.textContent).toBe('Headings');
    expect(bodyRow!.querySelector('.pptx-font-picker__role')!.textContent).toBe('Body');
    const face = headingsRow!.querySelector<HTMLElement>('.pptx-font-picker__face')!;
    expect(face.style.getPropertyValue('--pptx-font-face')).toBe('"Georgia"');
    const current = options().filter((o) => o.getAttribute('aria-selected') === 'true');
    expect(current.length).toBeGreaterThan(1);
    expect(current.every((o) => family(o) === 'Courier New')).toBe(true);
    expect(current.every((o) => o.querySelector('.pptx-menu__check') !== null)).toBe(true);
    // The Headings row writes the theme reference (`+mj-lt`), the trigger shows the theme's
    // family, and the reopened list marks the Headings row only.
    headingsRow!.click();
    await nextTick();
    expect(document.querySelectorAll('.pptx-font-picker').length).toBe(0);
    expect(trigger.textContent).toContain('Georgia');
    expect(latinOf(editor)).toBe('+mj-lt');
    trigger.click();
    await nextTick();
    expect(marked()).toEqual(['Georgia:major']);
  });

  test('a literal family with the theme font name marks the literal row, not the theme row', async () => {
    const editor = await fontDeck('Georgia');
    const host = await bar(editor);
    const trigger = host.querySelector<HTMLElement>('[data-slot="font.family"]')!;
    expect(trigger.textContent).toContain('Georgia');
    trigger.click();
    await nextTick();
    expect(marked()).toEqual(['Georgia:']);
    // The Body row writes `+mn-lt`; then only the Body row is marked.
    options()
      .find((o) => o.dataset.theme === 'minor')!
      .click();
    await nextTick();
    expect(latinOf(editor)).toBe('+mn-lt');
    expect(trigger.textContent).toContain('Verdana');
    trigger.click();
    await nextTick();
    expect(marked()).toEqual(['Verdana:minor']);
  });

  test('the search filters every group, and Enter applies the first match', async () => {
    const host = await bar(await fontDeck());
    const trigger = host.querySelector<HTMLElement>('[data-slot="font.family"]')!;
    trigger.click();
    await nextTick();
    const search = document.querySelector<HTMLInputElement>('.pptx-font-picker__search')!;
    search.value = 'verd';
    search.dispatchEvent(new Event('input', { bubbles: true }));
    await nextTick();
    expect(headings()).toEqual(['Theme fonts', 'All fonts']);
    expect(options().every((o) => /verd/i.test(family(o) ?? ''))).toBe(true);
    search.dispatchEvent(new KeyboardEvent('keydown', { key: 'Enter', bubbles: true }));
    await nextTick();
    expect(trigger.textContent).toContain('Verdana');
  });

  test('the list scrolls to the current family when it opens', async () => {
    const host = await bar(await fontDeck());
    const rect = Element.prototype.getBoundingClientRect;
    Element.prototype.getBoundingClientRect = function (this: Element) {
      const top = this.getAttribute('aria-selected') === 'true' ? 900 : 0;
      const height = this.classList.contains('pptx-font-picker__list') ? 300 : 30;
      return { top, bottom: top + height, height, left: 0, right: 100, width: 100 } as DOMRect;
    };
    try {
      host.querySelector<HTMLElement>('[data-slot="font.family"]')!.click();
      await nextTick();
      await nextTick();
      const list = document.querySelector<HTMLElement>('.pptx-font-picker__list')!;
      expect(list.scrollTop).toBeGreaterThan(0);
    } finally {
      Element.prototype.getBoundingClientRect = rect;
    }
  });
});
