// Accessibility of the editor chrome: slide items named by hidden text with the thumbnail art
// hidden from assistive technology (the rail and the slide sorter), layout tiles named by the
// layout, the viewport's "Presentation editor" landmark region, and the forced-colors and
// target-size rules of the stylesheet. The Vue twin is packages/vue/test/chrome-a11y.test.ts.

import { afterEach, describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { act, cleanup, fireEvent, render, screen } from '@testing-library/react';
import { createPptxEditor, type PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import { buildPptx, para, run, shape } from '../../../../scripts/fixtures/builder.ts';
import { buildSlideLayoutsDeck } from '../../../../scripts/fixtures/slide-layouts.ts';
import { PptxEditorContext } from '../context.ts';
import { PptxEditor } from '../index.ts';
import { SlideRail } from '../panes.tsx';
import { PptxEditorRoot, PptxEditorViewport } from '../PptxEditorRoot.tsx';
import { DialogProvider } from '../popups.tsx';
import { SlideSorter } from '../SlideSorter.tsx';

afterEach(() => {
  cleanup();
  document.body.replaceChildren();
});

const IN = 914_400;
const deck = () =>
  buildPptx({
    slides: [1, 2, 3].map((n) => ({
      layout: 'blank' as const,
      hidden: n === 2,
      shapes: shape({ id: 2, x: IN, y: IN, cx: 2 * IN, cy: IN, text: para(run(`Text ${n}`)) }),
    })),
  });

async function open(bytes = deck()): Promise<PptxEditorInstance> {
  const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await editor.load(bytes)).ok).toBe(true);
  return editor;
}

const tree = (editor: PptxEditorInstance, ui: React.ReactNode) => (
  <PptxEditorContext.Provider value={editor}>
    <DialogProvider>{ui}</DialogProvider>
  </PptxEditorContext.Provider>
);

/** The slide items name themselves by hidden text and hide the number and the art. */
function expectNamedItems(selector: string): void {
  const items = [...document.querySelectorAll<HTMLElement>(selector)];
  expect(items.length).toBe(3);
  for (const item of items) {
    expect(item.hasAttribute('aria-label')).toBe(false);
    expect(item.querySelector('.pptx-rail__thumb')?.getAttribute('aria-hidden')).toBe('true');
  }
  expect(screen.getByRole('option', { name: 'Slide 1' }) === items[0]).toBe(true);
  expect(screen.getByRole('option', { name: 'Slide 2 (hidden)' }) === items[1]).toBe(true);
  expect(screen.getByRole('option', { name: 'Slide 3' }) === items[2]).toBe(true);
}

describe('slide items', () => {
  test('rail items take their name from hidden text, and the thumbnail art is hidden', async () => {
    const editor = await open();
    render(tree(editor, <SlideRail />));
    expectNamedItems('.pptx-rail__item');
    expect(document.querySelector('.pptx-rail__number')?.getAttribute('aria-hidden')).toBe('true');
  });

  test('slide sorter tiles take their name from hidden text, and the art is hidden', async () => {
    const editor = await open();
    render(tree(editor, <SlideSorter />));
    expectNamedItems('.pptx-sorter__item');
    expect(document.querySelector('.pptx-sorter__number')?.getAttribute('aria-hidden')).toBe(
      'true'
    );
  });

  test('a layout tile is named by its layout, not by the placeholder prompts', async () => {
    const editor = await open(buildSlideLayoutsDeck());
    render(tree(editor, <SlideRail />));
    fireEvent.click(screen.getByRole('button', { name: 'Add slide with layout' }));
    const tile = screen.getByRole('menuitem', { name: 'Title and Content' });
    expect(tile.querySelector('.pptx-layout-gallery__thumb')?.getAttribute('aria-hidden')).toBe(
      'true'
    );
  });
});

describe('the editor landmark', () => {
  test('the viewport is a region named "Presentation editor"', () => {
    render(
      <PptxEditorRoot>
        <PptxEditorViewport />
      </PptxEditorRoot>
    );
    const region = screen.getByRole('region', { name: 'Presentation editor' });
    expect(region.classList.contains('pptx-editor')).toBe(true);
  });

  test('label names the region with text or a translation key, and false removes it', () => {
    const { rerender } = render(
      <PptxEditorRoot>
        <PptxEditorViewport label="Deck preview" />
      </PptxEditorRoot>
    );
    expect(screen.getByRole('region', { name: 'Deck preview' }) !== null).toBe(true);
    rerender(
      <PptxEditorRoot>
        <PptxEditorViewport label="rail.ariaLabel" />
      </PptxEditorRoot>
    );
    expect(screen.getByRole('region', { name: 'Slides' }) !== null).toBe(true);
    rerender(
      <PptxEditorRoot>
        <PptxEditorViewport label={false} />
      </PptxEditorRoot>
    );
    const scope = document.querySelector('.pptx-editor')!;
    expect(scope.hasAttribute('role') || scope.hasAttribute('aria-label')).toBe(false);
  });

  test('the default editor puts the toolbar, the rail, the canvas, and the notes in it', async () => {
    let editor: PptxEditorInstance | null = null;
    render(<PptxEditor onReady={(e) => (editor = e)} />);
    for (let i = 0; i < 300 && !editor; i += 1)
      await act(() => new Promise((r) => setTimeout(r, 10)));
    await act(async () => void (await editor!.load(deck())));
    const region = screen.getByRole('region', { name: 'Presentation editor' });
    for (const part of ['.pptx-default__dock', '.pptx-canvas']) {
      expect(region.contains(region.querySelector(part))).toBe(true);
    }
    expect(region.contains(screen.getByRole('navigation', { name: 'Slides' }))).toBe(true);
    expect(region.contains(screen.getByRole('region', { name: 'Speaker notes' }))).toBe(true);
  });
});

describe('the stylesheet', () => {
  const css = readFileSync(
    new URL('../../../core/src/styles/editor.css', import.meta.url),
    'utf8'
  ).replace(/\/\*[\s\S]*?\*\//g, '');
  const forced = css.slice(css.indexOf('@media (forced-colors: active)'));

  test('forced colors keep selection, pressed, focus, and disabled states visible', () => {
    expect(forced.length > 0 && css.includes('@media (forced-colors: active)')).toBe(true);
    for (const rule of [
      ".pptx-toolbar__button:is([aria-pressed='true'], [data-active])",
      ".pptx-view-switch .pptx-view-switch__option[aria-checked='true']",
      ".pptx-rail__item:is([aria-selected='true'], [aria-current='true']) .pptx-rail__thumb::after",
      ".pptx-sorter__item:is([aria-selected='true'], [aria-current='true']) .pptx-rail__thumb::after",
      '.pptx-editor :focus-visible',
      ".pptx-menu__item[aria-disabled='true']",
      '.pptx-surface__handle',
    ]) {
      expect(forced).toContain(rule);
    }
    for (const color of ['Highlight', 'HighlightText', 'GrayText', 'CanvasText']) {
      expect(forced).toContain(color);
    }
  });

  test('a color swatch is a 24px target', () => {
    expect(css).toMatch(/\.pptx-swatch \{\s*width: 24px;\s*height: 24px;/);
    expect(css).toContain('grid-template-columns: repeat(10, 24px);');
  });
});
