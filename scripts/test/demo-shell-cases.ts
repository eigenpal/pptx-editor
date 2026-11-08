// The demo shell cases that the React and Vue demos share, after the Playwright
// `e2e/adapter-render.spec.ts` and `e2e/narrow-chrome.spec.ts`, which ran each case on both
// demos: the adapter loop (load, paint, a menu, a dialog, the canvas context menu, the slide
// show), a file that does not open, slide and file drags, and the overlay and focus rules of a
// narrow window, with the demos' own stylesheets so their breakpoints apply. Each adapter's
// `demo-shell` test file runs them on its demo.

import { describe, expect, test } from 'bun:test';
import { screen, within } from '@testing-library/react';
import {
  barButton,
  click,
  clickAt,
  dataTransfer,
  dragEvent,
  dragItem,
  isDisabled,
  openFile,
  press,
  slidePoint,
  SLIDE_H,
  SLIDE_W,
  step,
  type,
  until,
  type Demo,
  type DemoOptions,
} from './demo-dom.ts';

const byRole = screen.getByRole;
const rail = () => [...document.querySelectorAll<HTMLElement>('.pptx-rail__item')];
const pdf = () => new File(['%PDF-1.4'], 'tables.pdf', { type: 'application/pdf' });
const focusedClass = () => (document.activeElement as HTMLElement | null)?.className ?? '';

/** Run the shell cases on one adapter's demo. */
export function demoShellCases(
  name: 'React' | 'Vue',
  openDemo: (options?: DemoOptions) => Promise<Demo>
): void {
  describe(`${name} demo: the adapter loop`, () => {
    test('load, paint, menu, dialog, context menu, slide show', async () => {
      const demo = await openDemo();
      // Load: the rail lists the five slides, and the canvas paints the first one.
      expect(rail()).toHaveLength(5);
      expect(document.querySelectorAll('.pptx-surface__slide text').length).toBeGreaterThan(0);

      // A menu, and a dialog from it.
      await click(byRole('menuitem', { name: 'File' }));
      expect(byRole('menu', { name: 'File' })).not.toBeNull();
      await click(byRole('menuitem', { name: 'Page setup…' }));
      const dialog = byRole('dialog', { name: 'Page setup' });
      await click(within(dialog).getByRole('button', { name: 'Cancel' }));
      expect(screen.queryByRole('dialog', { name: 'Page setup' })).toBeNull();

      // The canvas context menu, on an empty part of the slide.
      await clickAt(slidePoint(SLIDE_W * 0.97, SLIDE_H * 0.97), { button: 'right' });
      const menu = () => document.querySelector('.pptx-contextmenu');
      expect(menu()).not.toBeNull();
      expect(within(menu() as HTMLElement).getAllByRole('menuitem').length).toBeGreaterThan(0);
      await press('Escape');
      expect(menu()).toBeNull();

      // Start the slide show, then end it.
      await click(byRole('button', { name: 'Present' }));
      expect(screen.queryByRole('dialog', { name: 'Slide show' })).not.toBeNull();
      await press('Escape');
      expect(screen.queryByRole('dialog', { name: 'Slide show' })).toBeNull();
      expect(demo.errors).toEqual([]);
    });

    test('a file that does not open says why and keeps the deck', async () => {
      await openDemo();
      const name = () => document.querySelector<HTMLInputElement>('.demo-file__name')!.value;
      expect(name()).toBe('Basic sample');

      // A corrupt `.pptx` from Open: the notice says why, and the deck and its name stay.
      const input = byRole('banner').querySelector<HTMLInputElement>('input[type="file"]')!;
      await openFile(new File([new Uint8Array(17).fill(7)], 'corrupt.pptx'), input);
      const failure = () =>
        screen
          .queryAllByRole('alert')
          .find((a) => a.textContent?.includes('The file did not open'))
          ?.closest<HTMLElement>('.pptx-open-failure') ?? undefined;
      expect(failure()).toBeDefined();
      expect(failure()!.textContent).toContain('Error code: zip');
      expect(failure()!.textContent).toContain('The presentation that was open stays open.');
      expect(rail()).toHaveLength(5);
      expect(name()).toBe('Basic sample');
      expect(document.title).toMatch(/^Basic sample - /);
      await click(within(failure()!).getByRole('button', { name: 'Close' }));
      expect(failure()).toBeUndefined();

      // A dropped file that is not a `.pptx`: the edit notice says why.
      const zone = document.querySelector('.demo-workspace')!;
      await step(() => dragEvent(zone, 'drop', dataTransfer([pdf()])));
      expect(document.querySelector('.pptx-edit-notice')!.textContent).toContain(
        'tables.pdf is not a .pptx file.'
      );
    });

    test('a deck that does not load says why, in place of the spinner', async () => {
      const demo = await openDemo({ failDecks: true });
      await until(() => {
        const notice = document.querySelector('.pptx-edit-notice');
        if (!notice?.textContent?.includes('The presentation did not load.'))
          throw new Error('no notice');
      });
      expect(document.querySelector('.pptx-loading')).toBeNull();
      expect(demo.editor.snapshot().status).toBe('empty');
      expect(demo.errors.join('\n')).toContain('HTTP 503 for sample.pptx');
    });

    test('a mouse drag in the rail reorders slides, with no file-drop hint', async () => {
      const demo = await openDemo();
      const undo = () => screen.getAllByRole('button', { name: /^Undo/ })[0];
      expect(isDisabled(undo())).toBe(true);
      const before = demo.editor.snapshot().slides.map((s) => s.id);

      // A drag of slide 1 below slide 3: the demo's file-drop hint stays hidden, and the drop
      // reaches the rail.
      let hint = -1;
      await dragItem(rail()[0], rail()[2], 'after', () => {
        hint = document.querySelectorAll('.demo-drop').length;
      });
      expect(hint).toBe(0);
      expect(isDisabled(undo())).toBe(false);
      expect(document.querySelectorAll('.demo-drop')).toHaveLength(0);
      expect(demo.editor.snapshot().slides.map((s) => s.id)).toEqual([
        before[1],
        before[2],
        before[0],
        before[3],
        before[4],
      ]);
    });

    test('a file drag over the top bar shows the hint, clears it, and drops', async () => {
      await openDemo();
      const bar = document.querySelector('.demo-topbar')!;
      const fire = async (type: string) => {
        let cancelled = false;
        await step(() => {
          cancelled = dragEvent(bar, type, dataTransfer([pdf()]), { relatedTarget: null });
        });
        return cancelled;
      };
      const hint = () => document.querySelectorAll('.demo-drop').length;

      // The drag shows the hint, and leaving the window clears it.
      expect(await fire('dragover')).toBe(true);
      expect(hint()).toBe(1);
      await fire('dragleave');
      expect(hint()).toBe(0);

      // A drop on the top bar is the page's own: it is cancelled and handled.
      await fire('dragover');
      expect(await fire('drop')).toBe(true);
      expect(hint()).toBe(0);
      expect(document.querySelector('.pptx-edit-notice')!.textContent).toContain(
        'tables.pdf is not a .pptx file.'
      );
    });
  });

  describe(`${name} demo: rail thumbnails`, () => {
    test('a slide move keeps every painted thumbnail, with no repaint', async () => {
      const demo = await openDemo();
      const thumbs = () =>
        rail().map((item) => item.querySelector('.pptx-rail__thumb > svg') as Element | null);
      const ids = () => demo.editor.snapshot().slides.map((s) => s.id);
      const painted = new Map(ids().map((id, i) => [id, thumbs()[i]]));
      expect([...painted.values()].every((svg) => svg !== null)).toBe(true);
      // The first slide moves to the end: each item follows its slide, picture and all.
      await step(() => demo.editor.exec({ type: 'moveSlides', indices: [0], to: 4 }));
      const after = thumbs();
      // Compare identities, not elements: a failing match would print the whole DOM.
      expect(ids().map((id, i) => after[i] === painted.get(id))).toEqual(ids().map(() => true));
    });

    test('a removed slide releases its thumbnail container', async () => {
      const demo = await openDemo();
      const ed = demo.editor;
      const own = ed.renderThumbnail.bind(ed);
      const released = new Set<Element>();
      ed.renderThumbnail = (index, container, options) => {
        const dispose = own(index, container, options);
        return () => (released.add(container), dispose());
      };
      try {
        // The first two slides swap, so both items ask again through the wrapper.
        const first = ed.snapshot().slides[0]!.id;
        await step(() => ed.exec({ type: 'moveSlides', indices: [0], to: 1 }));
        const index = ed.snapshot().slides.findIndex((s) => s.id === first);
        const box = rail()[index]!.querySelector('.pptx-rail__thumb')!;
        released.clear();
        await step(() => ed.exec({ type: 'deleteSlides', indices: [index] }));
        expect(released.has(box)).toBe(true);
      } finally {
        ed.renderThumbnail = own;
      }
    });
  });

  describe(`${name} demo: a narrow window`, () => {
    test('the narrow rail overlay closes on a pick, Escape, and a press outside', async () => {
      const demo = await openDemo({ width: 820, css: true });
      const toggle = () => byRole('button', { name: 'Show or hide slide thumbnails' });
      const nav = () => screen.queryByRole('navigation', { name: 'Slides' });
      expect(nav()).toBeNull();

      // A picked slide shows, the rail closes, and focus goes to the slide.
      await click(toggle());
      expect(nav()).not.toBeNull();
      await click(within(nav()!).getByRole('option', { name: 'Slide 3' }));
      expect(nav()).toBeNull();
      expect(toggle().getAttribute('aria-pressed')).toBe('false');
      expect(demo.editor.snapshot().currentSlide).toBe(2);
      expect(focusedClass()).toContain('pptx-surface');

      // Escape from inside the rail closes it and gives focus back to the toggle.
      await click(toggle());
      await step(() => within(nav()!).getByRole('option', { name: 'Slide 3' }).focus());
      await press('Escape');
      expect(nav()).toBeNull();
      expect(document.activeElement).toBe(toggle());

      // A press outside (the speaker notes, which start collapsed) closes it; the notes keep the
      // press and open.
      await click(toggle());
      expect(nav()).not.toBeNull();
      const notes = byRole('button', { name: 'Show speaker notes' });
      await click(notes);
      expect(nav()).toBeNull();
      expect(byRole('textbox', { name: 'Speaker notes' })).not.toBeNull();
    });

    test('the narrow side panel overlay closes on Escape', async () => {
      await openDemo({ width: 820, css: true });
      const toggle = () => byRole('button', { name: 'Toggle side panel' });
      const panel = () => screen.queryByRole('complementary', { name: 'Side panel' });
      expect(panel()).toBeNull();
      await click(toggle());
      expect(panel()).not.toBeNull();
      await step(() => within(panel()!).getByRole('tab', { name: 'Format' }).focus());
      await press('Escape');
      expect(panel()).toBeNull();
      expect(document.activeElement).toBe(toggle());
    });

    test('Enter in the sorter on a narrow window focuses the slide', async () => {
      const demo = await openDemo({ width: 820, css: true });
      await step(() => demo.editor.exec({ type: 'setPresentationView', view: 'slideSorter' }));
      const current = document.querySelector<HTMLElement>(
        '.pptx-sorter__item[aria-current="true"]'
      )!;
      await step(() => current.focus());
      await press('Enter');
      expect(document.querySelectorAll('.pptx-sorter__item')).toHaveLength(0);
      expect(focusedClass()).toContain('pptx-surface');
    });

    test('Escape in a menu bar menu focuses its menubar item', async () => {
      await openDemo({ css: true });
      await clickAt(slidePoint(0, 0));
      const format = byRole('menuitem', { name: 'Format' });
      await click(format);
      expect(screen.queryByRole('menu', { name: 'Format' })).not.toBeNull();
      await press('Escape');
      expect(screen.queryByRole('menu', { name: 'Format' })).toBeNull();
      expect(document.activeElement).toBe(format);
    });

    test('More opened from the keyboard focuses its first control', async () => {
      // The caret in a new text box: the text bar does not fit and keeps a More button.
      await openDemo({ width: 820, css: true, barWidth: 60 });
      await click(await barButton('Text box'));
      await type('alpha');
      const more = within(byRole('toolbar', { name: 'Formatting' })).getByRole('button', {
        name: 'More',
      });
      await step(() => more.focus());
      await press('Enter');
      const panel = document.querySelector('.pptx-toolbar__overflow')!;
      expect(panel).not.toBeNull();
      expect(document.activeElement).toBe(panel.querySelector('button, input, [tabindex="0"]'));
      await press('Escape');
      expect(document.activeElement).toBe(more);
    });

    test('Ctrl+F opens the find bar in the canvas, and Escape closes it', async () => {
      // Where the bar sits over the slide is CSS layout: `e2e/smoke.spec.ts` measures it.
      await openDemo({ css: true });
      await clickAt(slidePoint(0, 0));
      await press('ControlOrMeta+f');
      const bar = () => screen.queryByRole('search', { name: 'Find and replace' });
      expect(bar()).not.toBeNull();
      expect(bar()!.closest('.demo-canvas')).not.toBeNull();
      await press('Escape');
      expect(bar()).toBeNull();
    });
  });
}
