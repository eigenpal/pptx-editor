// Tables on the React demo in happy-dom: Insert > Table from the menu grid, typing in the first
// cell, Tab to the next cell, Tab in the last cell adding a row, and Table Layout > Insert below.
// Ported from the Playwright `e2e/tables.spec.ts`.

import { afterEach, describe, expect, test } from 'bun:test';
import {
  click,
  closeDemo,
  openDemo,
  paintedText,
  press,
  screen,
  type,
  until,
} from './demo-harness.tsx';

afterEach(closeDemo);

const byRole = screen.getByRole;
const tables = () =>
  document.querySelectorAll('.pptx-surface__slide [data-kind="table-borders"]').length;
const slideText = () => document.querySelector('.pptx-surface__slide')!.textContent ?? '';

describe('demo tables', () => {
  test('insert a table from the menu, type in its cells, and move with Tab', async () => {
    const demo = await openDemo();
    const before = tables();

    // Insert > Table > 3 × 2: the table is inserted with the caret in its first cell.
    await click(byRole('menuitem', { name: 'Insert' }));
    await click(byRole('menuitem', { name: 'Table' }));
    await click(byRole('gridcell', { name: '3 × 2 table' }));
    expect(tables()).toBe(before + 1);

    // Type, then Tab to the next cell and type again.
    await type('Region');
    await press('Tab');
    await type('Sales');
    expect(slideText()).toContain('Region');
    expect(slideText()).toContain('Sales');

    // Tab through to the last cell, then once more: a row is added, and its first cell takes
    // the text below the first row.
    const height = () => {
      const all = document.querySelectorAll('.pptx-surface__slide [data-kind="table-borders"]');
      return all[all.length - 1].getBoundingClientRect().height;
    };
    const twoRows = height();
    for (let i = 0; i < 5; i += 1) await press('Tab');
    await type('North');
    const top = (text: RegExp) => paintedText(text)!.getBoundingClientRect().y;
    expect(top(/North/)).toBeGreaterThan(top(/Region/));
    expect(height()).toBeGreaterThan(twoRows);

    // Table Layout > Insert below: the table group follows the bar, so it can be in More.
    const threeRows = height();
    let below = document.querySelector<HTMLElement>('[data-slot="table.insertRowBelow"]');
    if (!below || below.closest('[hidden]')) {
      await click(byRole('button', { name: 'More' }));
      below = document.querySelector<HTMLElement>('[data-slot="table.insertRowBelow"]');
    }
    await click(below!);
    await until(() => expect(height()).toBeGreaterThan(threeRows));
    expect(demo.errors).toEqual([]);
  });

  test('the Insert > Table grid works from the keyboard', async () => {
    await openDemo();
    const before = tables();
    await click(byRole('menuitem', { name: 'Insert' }));
    await click(byRole('menuitem', { name: 'Table' }));
    // The grid takes the focus on its first cell once it shows.
    await until(() =>
      expect(document.activeElement).toBe(
        document.querySelector('[role="gridcell"][data-cell="1x1"]')
      )
    );
    await press('ArrowRight');
    await press('ArrowDown');
    expect(document.activeElement).toBe(
      document.querySelector('[role="gridcell"][data-cell="2x2"]')
    );
    await press('Enter');
    expect(tables()).toBe(before + 1);
    // The caret is in the new table's first cell: typing goes there.
    await type('Keyed');
    expect(slideText()).toContain('Keyed');
  });
});
