// The slide rail and the notes on the React demo in happy-dom: Shift+click and Ctrl/Cmd+click
// select slides and Delete, undo, and duplicate act on all of them; Insert > Slide numbers; and
// Tab and Escape take the focus out of the notes. Ported from the Playwright
// `e2e/slides.spec.ts`.

import { afterEach, describe, expect, test } from 'bun:test';
import {
  click,
  closeDemo,
  openDemo,
  paintedTexts,
  press,
  screen,
  type,
  until,
} from './demo-harness.tsx';

afterEach(closeDemo);

const byRole = screen.getByRole;
const items = () => [...document.querySelectorAll<HTMLElement>('.pptx-rail__item')];
const selected = () =>
  items().flatMap((el, i) => (el.getAttribute('aria-selected') === 'true' ? [i] : []));

describe('demo slide rail', () => {
  test('Shift+click and Ctrl/Cmd+click select slides; Delete and undo act on all of them', async () => {
    const demo = await openDemo();
    await click(items()[1]);
    await click(items()[3], { modifiers: { shiftKey: true } });
    expect(selected()).toEqual([1, 2, 3]);
    await click(items()[2], { modifiers: { ctrlKey: true } });
    expect(selected()).toEqual([1, 3]);
    await press('Delete');
    expect(items()).toHaveLength(3);
    await press('ControlOrMeta+z');
    expect(items()).toHaveLength(5);
    // Ctrl/Cmd+D duplicates the whole selection after its last slide.
    await click(items()[0]);
    await click(items()[1], { modifiers: { shiftKey: true } });
    await press('ControlOrMeta+d');
    expect(items()).toHaveLength(7);
    expect(selected()).toEqual([2, 3]);
    expect(demo.errors).toEqual([]);
  });

  test('Insert > Slide numbers numbers every slide and shows a check', async () => {
    await openDemo();
    expect(paintedTexts()).not.toContain('1');
    await click(byRole('menuitem', { name: 'Insert' }));
    await click(byRole('menuitemcheckbox', { name: /Slide numbers/ }));
    await until(() => expect(paintedTexts()).toContain('1'));
    await click(items()[2]);
    await until(() => expect(paintedTexts()).toContain('3'));
    await click(byRole('menuitem', { name: 'Insert' }));
    expect(byRole('menuitemcheckbox', { name: /Slide numbers/ }).getAttribute('aria-checked')).toBe(
      'true'
    );
  });

  test('Tab and Escape after typing in the notes take the focus out of the field', async () => {
    await openDemo();
    const notes = document.querySelector<HTMLElement>('.pptx-notes__text')!;
    await click(notes);
    await press('ControlOrMeta+End');
    await type(' typed');
    await press('Tab');
    expect(document.activeElement).not.toBe(notes);
    expect(notes.textContent).toContain('typed');
    await click(notes);
    await type(' more');
    await press('Escape');
    expect(document.activeElement).not.toBe(notes);
    expect(document.activeElement).toBe(
      document.querySelector('.pptx-rail__item[aria-current="true"]')
    );
    expect(notes.textContent).toContain('more');
  });
});
