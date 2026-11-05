// The font family picker in the React adapter: theme fonts (Headings, Body), the fonts in this
// presentation, then all fonts; each name in its own face; the current family marked and
// scrolled into view when the list opens; a pick applies and closes. After DOCX Editor's
// FontFamily compound test.

import { afterEach, describe, expect, test } from 'bun:test';
import { act, cleanup, fireEvent, render } from '@testing-library/react';
import { createPptxEditor, type PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import { serializeXmlPart, xmlPartOf } from '@pptx-editor.dev/core/store';
import { buildPptx, para, run, shape } from '../../../../scripts/fixtures/builder.ts';
import { PptxEditorContext } from '../context.ts';
import { Toolbar } from '../index.ts';

const IN = 914_400;

async function open(latin = 'Courier New'): Promise<PptxEditorInstance> {
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

afterEach(() => {
  cleanup();
  for (const el of document.querySelectorAll('.pptx-menu')) el.remove();
});

const headings = () =>
  [...document.querySelectorAll('.pptx-font-picker__heading')].map((h) => h.textContent);
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

describe('the font family picker', () => {
  test('groups theme fonts, the deck families, and all fonts, each in its own face', async () => {
    const editor = await open();
    const { container } = render(
      <PptxEditorContext.Provider value={editor}>
        <Toolbar />
      </PptxEditorContext.Provider>
    );
    const trigger = container.querySelector<HTMLElement>('[data-slot="font.family"]')!;
    expect(trigger.textContent).toContain('Courier New');
    act(() => void fireEvent.click(trigger));
    expect(headings()).toEqual(['Theme fonts', 'Fonts in this presentation', 'All fonts']);
    // Each group is a labeled group of the list box.
    for (const g of document.querySelectorAll('.pptx-font-picker__group')) {
      expect(g.getAttribute('role')).toBe('group');
      expect(document.getElementById(g.getAttribute('aria-labelledby')!)!.textContent).toBe(
        g.querySelector('.pptx-font-picker__heading')!.textContent
      );
    }
    const [headingsRow, bodyRow] = options();
    expect(family(headingsRow!)).toBe('Georgia');
    expect(headingsRow!.querySelector('.pptx-font-picker__role')!.textContent).toBe('Headings');
    expect(family(bodyRow!)).toBe('Verdana');
    expect(bodyRow!.querySelector('.pptx-font-picker__role')!.textContent).toBe('Body');
    // The name is drawn in its own face.
    const face = headingsRow!.querySelector<HTMLElement>('.pptx-font-picker__face')!;
    expect(face.style.getPropertyValue('--pptx-font-face')).toBe('"Georgia"');
    // The current family is marked in every group that lists it, with a check mark.
    const current = options().filter((o) => o.getAttribute('aria-selected') === 'true');
    expect(current.length).toBeGreaterThan(1);
    expect(current.every((o) => family(o) === 'Courier New')).toBe(true);
    expect(current.every((o) => o.querySelector('.pptx-menu__check') !== null)).toBe(true);
    // A pick applies and closes. The Headings row writes the theme reference (`+mj-lt`), the
    // trigger shows the theme's family, and the reopened list marks the Headings row only.
    act(() => void fireEvent.click(headingsRow!));
    expect(document.querySelectorAll('.pptx-font-picker').length).toBe(0);
    expect(trigger.textContent).toContain('Georgia');
    expect(latinOf(editor)).toBe('+mj-lt');
    act(() => void fireEvent.click(trigger));
    expect(marked()).toEqual(['Georgia:major']);
  });

  test('a literal family with the theme font name marks the literal row, not the theme row', async () => {
    const editor = await open('Georgia');
    const { container } = render(
      <PptxEditorContext.Provider value={editor}>
        <Toolbar />
      </PptxEditorContext.Provider>
    );
    const trigger = container.querySelector<HTMLElement>('[data-slot="font.family"]')!;
    expect(trigger.textContent).toContain('Georgia');
    act(() => void fireEvent.click(trigger));
    expect(marked()).toEqual(['Georgia:']);
    // The Body row writes `+mn-lt`; then only the Body row is marked.
    const body = options().find((o) => o.dataset.theme === 'minor')!;
    act(() => void fireEvent.click(body));
    expect(latinOf(editor)).toBe('+mn-lt');
    expect(trigger.textContent).toContain('Verdana');
    act(() => void fireEvent.click(trigger));
    expect(marked()).toEqual(['Verdana:minor']);
  });

  test('the search filters every group, and Enter applies the first match', async () => {
    const editor = await open();
    const { container } = render(
      <PptxEditorContext.Provider value={editor}>
        <Toolbar />
      </PptxEditorContext.Provider>
    );
    const trigger = container.querySelector<HTMLElement>('[data-slot="font.family"]')!;
    act(() => void fireEvent.click(trigger));
    const search = document.querySelector<HTMLInputElement>('.pptx-font-picker__search')!;
    expect(document.activeElement).toBe(search);
    act(() => void fireEvent.change(search, { target: { value: 'verd' } }));
    expect(headings()).toEqual(['Theme fonts', 'All fonts']);
    expect(options().every((o) => /verd/i.test(family(o) ?? ''))).toBe(true);
    // Down moves from the field to the first row.
    act(() => void fireEvent.keyDown(search, { key: 'ArrowDown' }));
    expect(document.activeElement).toBe(options()[0]!);
    act(() => void fireEvent.keyDown(search, { key: 'Enter' }));
    expect(trigger.textContent).toContain('Verdana');
  });

  test('the list scrolls to the current family when it opens', async () => {
    const editor = await open();
    const rect = Element.prototype.getBoundingClientRect;
    // happy-dom has no layout: the list is 300 pixels tall, and the current row is far below.
    Element.prototype.getBoundingClientRect = function (this: Element) {
      const top = this.getAttribute('aria-selected') === 'true' ? 900 : 0;
      const height = this.classList.contains('pptx-font-picker__list') ? 300 : 30;
      return { top, bottom: top + height, height, left: 0, right: 100, width: 100 } as DOMRect;
    };
    try {
      const { container } = render(
        <PptxEditorContext.Provider value={editor}>
          <Toolbar />
        </PptxEditorContext.Provider>
      );
      act(() => void fireEvent.click(container.querySelector('[data-slot="font.family"]')!));
      const list = document.querySelector<HTMLElement>('.pptx-font-picker__list')!;
      expect(list.scrollTop).toBeGreaterThan(0);
    } finally {
      Element.prototype.getBoundingClientRect = rect;
    }
  });
});
