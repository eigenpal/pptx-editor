// Smoke, on the React demo in happy-dom: the v0.1 test script, view mode, and the font notice
// on the sample deck. Ported from the Playwright `e2e/smoke.spec.ts`.

import { afterEach, describe, expect, test } from 'bun:test';
import { within } from '@testing-library/react';
import { presentationInfo, readPptx } from '@pptx-editor.dev/core/store';
import { notesText } from '../../../core/src/automation/commands/notes-commands.ts';
import {
  blur,
  click,
  clickAt,
  closeDemo,
  drag,
  dragItem,
  fill,
  isDisabled,
  openDemo,
  openFile,
  press,
  repoFile,
  screen,
  slidePoint,
  type,
  until,
} from './demo-harness.tsx';

afterEach(closeDemo);

const byRole = screen.getByRole;
const items = () => [...document.querySelectorAll<HTMLElement>('.pptx-rail__item')];

describe('demo smoke', () => {
  test('the v0.1 test script', async () => {
    const demo = await openDemo();

    // A click on the title's border selects it; drag it from there.
    const edge = slidePoint(640_080 + 2_000_000, 2_468_880);
    await clickAt({ x: edge.x, y: edge.y + 2 });
    expect(document.querySelectorAll('.pptx-surface__handle')).toHaveLength(9);
    await drag({ x: edge.x, y: edge.y + 2 }, { x: edge.x + 40, y: edge.y + 32 });
    expect((screen.getByLabelText('X', { exact: true }) as HTMLInputElement).value).not.toBe(
      '0.70'
    );

    // A single click on the text places the caret; then italicize it with the toolbar.
    const title = slidePoint(640_080 + 2_000_000, 2_468_880 + 500_000);
    await clickAt({ x: title.x + 40, y: title.y + 30 });
    expect(document.querySelectorAll('.pptx-surface__caret')).toHaveLength(1);
    await press('End');
    await type(' rocks');
    await press('Shift+Home');
    await click(byRole('button', { name: /^Italic \(/ }));
    expect(byRole('button', { name: /^Italic \(/ }).getAttribute('aria-pressed')).toBe('true');

    // Undo the italic, then redo it.
    await click(byRole('button', { name: /^Undo \(/ }));
    expect(byRole('button', { name: /^Italic \(/ }).getAttribute('aria-pressed')).toBe('false');
    await click(byRole('button', { name: /^Redo \(/ }));

    // Insert a shape and a text box.
    await press('Escape');
    await click(byRole('button', { name: 'Text box' }));
    await type('Inserted text');
    await click(byRole('button', { name: 'Shape' }));
    await click(byRole('menuitem', { name: 'Oval' }));

    // Slides: add, then reorder by dragging in the rail.
    await click(byRole('button', { name: 'New slide' }));
    expect(items()).toHaveLength(6);
    const order = () => demo.editor.snapshot().slides.map((s) => s.id);
    const before = order();
    await dragItem(items()[0], items()[2], 'after');
    expect(order()).toEqual([before[1], before[2], before[0], ...before.slice(3)]);

    // Speaker notes.
    await click(items()[0]);
    const notes = document.querySelector<HTMLElement>('.pptx-notes__text')!;
    await fill(notes, 'Edited in the e2e test');
    await blur(notes);

    // Export and check the file.
    await click(byRole('button', { name: 'Export' }));
    await click(byRole('menuitem', { name: /\(\.pptx\)/ }));
    await until(() => expect(demo.downloads).toHaveLength(1));
    const saved = readPptx(new Uint8Array(await demo.downloads[0].blob.arrayBuffer()));
    expect(saved.ok).toBe(true);
    if (saved.ok) {
      expect(presentationInfo(saved.pkg).slides).toHaveLength(6);
      expect(notesText(saved.pkg, 0)).toBe('Edited in the e2e test');
    }
    expect(demo.errors).toEqual([]);
  });

  test('view mode refuses edits', async () => {
    await openDemo();
    await click(byRole('button', { name: /^Mode/ }));
    await click(byRole('menuitemradio', { name: /Viewing/ }));
    expect(isDisabled(byRole('button', { name: 'Text box' }))).toBe(true);
    expect(isDisabled(byRole('button', { name: 'New slide' }))).toBe(true);
  });

  test('the font notice lists substitutes that change line breaks', async () => {
    await openDemo();
    expect(document.querySelectorAll('.pptx-font-notice')).toHaveLength(0);
    await openFile(repoFile('fixtures/font-substitution.pptx'));
    await until(() => expect(document.querySelector('.pptx-font-notice')).not.toBeNull());
    const notice = document.querySelector<HTMLElement>('.pptx-font-notice')!;
    expect(notice.textContent).toContain('Aptos');
    expect(notice.textContent).toContain('Segoe UI');
    // Calibri → Carlito keeps line breaks, so only Calibri Light (no such substitute) is listed.
    expect(notice.textContent).not.toMatch(/Calibri(?! Light)/);
    // Every substitution, metric-compatible ones included, is listed in File → Page setup.
    await click(byRole('menuitem', { name: 'File' }));
    await click(byRole('menuitem', { name: 'Page setup…' }));
    const dialog = byRole('dialog', { name: 'Page setup' });
    expect(dialog.querySelector('.pptx-dialog__list')!.textContent).toContain('Calibri → Carlito');
    await click(within(dialog).getByRole('button', { name: 'OK' }));
    expect(screen.queryByRole('dialog', { name: 'Page setup' })).toBeNull();
    await click(within(notice).getByRole('button', { name: 'Dismiss' }));
    expect(document.querySelectorAll('.pptx-font-notice')).toHaveLength(0);
  });
});
