// The `PptxEditorToolbar.FontFamily` compound in the Vue adapter: the twin of the React test
// (packages/react/src/__tests__/font-family-parts.test.tsx), after DOCX Editor's `FontFamily`
// Trigger / Content / Item parts.

import { afterEach, describe, expect, spyOn, test } from 'bun:test';
import { createApp, h, nextTick, type App } from 'vue';
import {
  createPptxEditor,
  resetDevWarnings,
  type PptxEditorInstance,
} from '@pptx-editor.dev/core/editor';
import { serializeXmlPart, xmlPartOf } from '@pptx-editor.dev/core/store';
import { buildPptx, para, run, shape } from '../../../scripts/fixtures/builder.ts';
import { PptxEditorToolbar } from '../src/index.ts';
import { mountVue } from './toolbar-dom.ts';

const IN = 914_400;
const FontFamily = PptxEditorToolbar.FontFamily;

const apps: [App, HTMLElement][] = [];
afterEach(() => {
  for (const [app, host] of apps.splice(0)) {
    app.unmount();
    host.remove();
  }
  for (const el of document.querySelectorAll('.pptx-menu')) el.remove();
});

async function open(minorFont = 'Verdana'): Promise<PptxEditorInstance> {
  const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  const bytes = buildPptx({
    majorFont: 'Georgia',
    minorFont,
    slides: [
      {
        layout: 'blank',
        shapes: shape({
          id: 2,
          x: IN,
          y: IN,
          cx: 2 * IN,
          cy: IN,
          text: para(run('Hi', '', '<a:latin typeface="Courier New"/>')),
        }),
      },
    ],
  });
  expect((await editor.load(bytes)).ok).toBe(true);
  const id = editor.session!.layout(0)!.items[0]!.element.nodeId;
  editor.setSelection({ kind: 'shapes', nodeIds: [id] });
  return editor;
}

const latinOf = (editor: PptxEditorInstance): string | undefined => {
  const part = xmlPartOf(editor.session!.store.package, editor.session!.info.slides[0]!.partName)!;
  return /<a:latin typeface="([^"]*)"/.exec(serializeXmlPart(part))?.[1];
};

describe('PptxEditorToolbar.FontFamily parts (Vue)', () => {
  test('a host composes Trigger, Content, and Item around the shared state', async () => {
    const editor = await open();
    const { host, app } = await mountVue(editor, () =>
      h(PptxEditorToolbar, null, () => [
        h(FontFamily, { className: 'acme-family' }, () => [
          h(FontFamily.Trigger, { class: 'acme-trigger' }, () => 'Font'),
          h(FontFamily.Content, { className: 'acme-list' }, () => [
            h(FontFamily.Item, { value: '+mj-lt' }),
            h(FontFamily.Item, { value: 'Arial', className: 'acme-item' }, () => 'Plain Arial'),
            h(FontFamily.Item, { value: 'Courier New' }),
            h(FontFamily.Item, { value: 'bad"name' }),
          ]),
        ]),
      ])
    );
    apps.push([app, host]);
    await nextTick();
    const trigger = host.querySelector<HTMLElement>('[data-slot="font.family"]')!;
    expect(trigger.textContent).toBe('Font');
    expect(trigger.className).toContain('acme-trigger');
    expect(trigger.className).toContain('acme-family');
    expect(trigger.getAttribute('aria-expanded')).toBe('false');
    trigger.click();
    await nextTick();
    expect(trigger.getAttribute('aria-expanded')).toBe('true');
    expect(document.querySelectorAll('.pptx-font-picker__search').length).toBe(0);
    expect(document.querySelectorAll('.acme-list').length).toBe(1);
    const rows = [...document.querySelectorAll<HTMLElement>('[role="option"]')];
    expect(rows.map((r) => r.textContent)).toEqual([
      'GeorgiaHeadings',
      'Plain Arial',
      'Courier New',
    ]);
    expect(rows[0]!.dataset.theme).toBe('major');
    expect(rows[1]!.className).toContain('acme-item');
    expect(rows.map((r) => r.getAttribute('aria-selected'))).toEqual(['false', 'false', 'true']);
    rows[0]!.click();
    await nextTick();
    expect(latinOf(editor)).toBe('+mj-lt');
    expect(trigger.getAttribute('aria-expanded')).toBe('false');
    editor.destroy();
  });

  test('with no slot content the default trigger and grouped list render', async () => {
    const editor = await open();
    const { host, app } = await mountVue(editor, () =>
      h(PptxEditorToolbar, null, () => [h(FontFamily, { className: 'acme-family' })])
    );
    apps.push([app, host]);
    const trigger = host.querySelector<HTMLElement>('[data-slot="font.family"]')!;
    expect(trigger.className).toContain('acme-family');
    expect(trigger.textContent).toContain('Courier New');
    trigger.click();
    await nextTick();
    expect(document.querySelectorAll('.pptx-font-picker__search').length).toBe(1);
    expect(document.querySelectorAll('.pptx-font-picker__group').length).toBe(3);
    editor.destroy();
  });

  test('a host list takes focus at its current option, and the arrow keys move', async () => {
    resetDevWarnings();
    const warn = spyOn(console, 'warn').mockImplementation(() => undefined);
    // The theme's body font is empty: its item shows the role, disabled.
    const editor = await open('');
    const { host, app } = await mountVue(editor, () =>
      h(PptxEditorToolbar, null, () => [
        h(FontFamily, null, () => [
          h(FontFamily.Trigger),
          h(FontFamily.Content, null, () => [
            h(FontFamily.Item, { value: '+mn-lt' }),
            h(FontFamily.Item, { value: 'Arial' }),
            h(FontFamily.Item, { value: 'Courier New' }),
            h(FontFamily.Item, { value: 'bad"name' }),
          ]),
        ]),
      ])
    );
    apps.push([app, host]);
    await nextTick();
    const trigger = host.querySelector<HTMLElement>('[data-slot="font.family"]')!;
    trigger.focus();
    trigger.click();
    await nextTick();
    await nextTick();
    const rows = [...document.querySelectorAll<HTMLButtonElement>('[role="option"]')];
    expect(rows.map((r) => r.textContent)).toEqual(['Body', 'Arial', 'Courier New']);
    expect(rows.map((r) => r.disabled)).toEqual([true, false, false]);
    expect(rows[0]!.getAttribute('aria-selected')).toBe('false');
    expect(warn.mock.calls.some((c) => String(c[0]).includes('value "bad"name"'))).toBe(true);
    // The current option takes focus; the arrow keys skip the disabled row and wrap.
    expect(document.activeElement).toBe(rows[2]!);
    const key = (el: HTMLElement, k: string): boolean =>
      el.dispatchEvent(new KeyboardEvent('keydown', { key: k, bubbles: true, cancelable: true }));
    key(rows[2]!, 'ArrowDown');
    expect(document.activeElement).toBe(rows[1]!);
    key(rows[1]!, 'ArrowUp');
    expect(document.activeElement).toBe(rows[2]!);
    rows[0]!.click();
    await nextTick();
    expect(latinOf(editor)).toBe('Courier New');
    warn.mockRestore();
    editor.destroy();
  });

  test('a part outside the compound renders nothing and warns', async () => {
    resetDevWarnings();
    const warn = spyOn(console, 'warn').mockImplementation(() => undefined);
    const host = document.createElement('div');
    document.body.append(host);
    const app = createApp({ render: () => h(FontFamily.Item, { value: 'Arial' }) });
    app.mount(host);
    apps.push([app, host]);
    await nextTick();
    expect(host.querySelectorAll('[role="option"]').length).toBe(0);
    expect(
      warn.mock.calls.some((c) =>
        String(c[0]).includes(
          'PptxEditorToolbar.FontFamily.Item is used outside PptxEditorToolbar.FontFamily'
        )
      )
    ).toBe(true);
    warn.mockRestore();
  });
});
