// The slide panes after the UX review: a quick drag in the rail or the sorter reorders before
// React renders again, the zoom trigger says whether its menu is open, the side panel landmark
// has a name, and focus goes to the slide when the sorter closes with the rail hidden.

import { afterEach, describe, expect, test } from 'bun:test';
import { act, cleanup, fireEvent, render, screen } from '@testing-library/react';
import { createPptxEditor, type PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import { buildPptx, para, run, shape } from '../../../../scripts/fixtures/builder.ts';
import { PptxEditorContext } from '../context.ts';
import { SidePanel, SlideRail, ZoomPill } from '../panes.tsx';
import { DialogProvider } from '../popups.tsx';
import { DefaultWorkspace, SlideSorter } from '../SlideSorter.tsx';

afterEach(() => {
  cleanup();
  document.body.replaceChildren();
});

const IN = 914_400;
const deck = () =>
  buildPptx({
    slides: [1, 2, 3].map((i) => ({
      layout: 'blank' as const,
      shapes: shape({ id: 2, x: IN, y: IN, cx: 2 * IN, cy: IN, text: para(run(`S${i}`)) }),
    })),
  });

async function open(): Promise<PptxEditorInstance> {
  const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await editor.load(deck())).ok).toBe(true);
  return editor;
}

const tree = (editor: PptxEditorInstance, ui: React.ReactNode) => (
  <PptxEditorContext.Provider value={editor}>
    <DialogProvider>{ui}</DialogProvider>
  </PptxEditorContext.Provider>
);

/** Drag item 0 onto item 2 with no React render between the events, as a quick drag does. */
function quickDrag(items: () => HTMLElement[]): boolean {
  let over = false;
  act(() => {
    // Nested in one act, the updates wait: each handler sees the render before the drag.
    fireEvent.dragStart(items()[0]!, { dataTransfer: { effectAllowed: '' } });
    over = !fireEvent.dragOver(items()[2]!, { clientX: 1000, clientY: 1000 });
    fireEvent.drop(items()[2]!);
  });
  return over;
}

const ids = (editor: PptxEditorInstance) => editor.snapshot().slides.map((s) => s.id);

describe('a quick drag', () => {
  test('in the rail: dragover takes the drop before a render, and the drop moves the slide', async () => {
    const editor = await open();
    const first = ids(editor)[0]!;
    render(tree(editor, <SlideRail />));
    const items = () => [...document.querySelectorAll<HTMLElement>('.pptx-rail__item')];
    // A dragover that does not call preventDefault() refuses the drop.
    expect(quickDrag(items)).toBe(true);
    expect(ids(editor).indexOf(first)).toBe(2);
  });

  test('in the sorter: the same', async () => {
    const editor = await open();
    const first = ids(editor)[0]!;
    render(tree(editor, <SlideSorter />));
    const items = () => [...document.querySelectorAll<HTMLElement>('.pptx-sorter__item')];
    expect(quickDrag(items)).toBe(true);
    expect(ids(editor).indexOf(first)).toBe(2);
  });
});

describe('chrome names and states', () => {
  test('the zoom trigger says whether its menu is open', async () => {
    const editor = await open();
    render(tree(editor, <ZoomPill />));
    const trigger = document.querySelector<HTMLElement>('.pptx-zoom__level')!;
    expect(trigger.getAttribute('aria-haspopup')).toBe('menu');
    expect(trigger.getAttribute('aria-expanded')).toBe('false');
    fireEvent.click(trigger);
    expect(trigger.getAttribute('aria-expanded')).toBe('true');
    fireEvent.click(trigger);
    expect(trigger.getAttribute('aria-expanded')).toBe('false');
  });

  test('the side panel landmark has a name', async () => {
    const editor = await open();
    render(tree(editor, <SidePanel />));
    expect(screen.getByRole('complementary', { name: 'Side panel' }).tagName).toBe('ASIDE');
  });
});

describe('leaving the sorter', () => {
  test('with the rail hidden, focus goes to the slide, not to the page body', async () => {
    const editor = await open();
    editor.exec({ type: 'setPresentationView', view: 'slideSorter' });
    render(
      tree(
        editor,
        <div className="pptx-editor">
          <DefaultWorkspace
            normal={
              <>
                <SlideRail />
                <div className="pptx-surface" tabIndex={0} />
              </>
            }
          />
        </div>
      )
    );
    // A narrow window hides the rail: its items cannot take focus.
    const block = HTMLElement.prototype.focus;
    HTMLElement.prototype.focus = function (this: HTMLElement, o?: FocusOptions) {
      if (!this.closest('.pptx-rail')) block.call(this, o);
    };
    try {
      document.querySelector<HTMLElement>('.pptx-sorter__item')!.focus();
      act(() => void editor.exec({ type: 'setPresentationView', view: 'normal' }));
      expect(document.activeElement === document.querySelector('.pptx-surface')).toBe(true);
    } finally {
      HTMLElement.prototype.focus = block;
    }
  });
});
