// Text editing on the React demo in happy-dom: arrow keys collapse a text range, and formatting
// and object moves survive undo, redo, export, and reopening. Ported from the Playwright
// `e2e/editing-workflow.spec.ts`.

import { afterEach, describe, expect, test } from 'bun:test';
import {
  click,
  closeDemo,
  drag,
  openDemo,
  openFile,
  paintedText,
  press,
  screen,
  type,
  until,
} from './demo-harness.tsx';

afterEach(closeDemo);

const byRole = screen.getByRole;
const slideText = () => document.querySelector('.pptx-surface__slide')!.textContent ?? '';
const outline = () => document.querySelector('.pptx-surface__outline')!.getBoundingClientRect();
const bold = () => byRole('button', { name: /^Bold \(/ }).getAttribute('aria-pressed');

async function shiftArrows(key: 'ArrowRight' | 'ArrowLeft', n: number): Promise<void> {
  for (let i = 0; i < n; i += 1) await press(`Shift+${key}`);
}

describe('demo editing workflow', () => {
  test('arrow keys collapse a text range before the next edit', async () => {
    await openDemo();
    await click(byRole('button', { name: 'Text box' }));
    await type('Alpha beta');
    await press('Home');
    await shiftArrows('ArrowRight', 5);
    await press('ArrowLeft');
    await type('X');
    expect(slideText()).toContain('XAlpha beta');
    await press('ControlOrMeta+z');
    await press('End');
    await shiftArrows('ArrowLeft', 4);
    await press('ArrowRight');
    await type('Y');
    expect(slideText()).toContain('Alpha betaY');
  });

  test('selected text formatting and object movement survive history and reopening', async () => {
    const demo = await openDemo();
    await click(byRole('button', { name: 'Text box' }));
    await type('Alpha beta');
    await press('Home');
    await shiftArrows('ArrowRight', 5);
    await click(byRole('button', { name: /^Bold \(/ }));
    await click(screen.getAllByRole('button', { name: 'Text color' })[1]);
    await click(byRole('menuitem', { name: 'Red' }));
    const word = () => paintedText(/^Alpha$/)?.getAttribute('fill');
    expect(word()).toBe('#ff0000');
    await press('ControlOrMeta+z');
    expect(word()).not.toBe('#ff0000');
    await press('ControlOrMeta+Shift+z');
    expect(word()).toBe('#ff0000');
    await press('Escape');
    const before = outline();
    await press('ArrowRight');
    await press('ArrowDown');
    expect(outline().x).toBeGreaterThan(before.x);
    const moved = outline();
    await press('ControlOrMeta+z');
    expect(outline().y).toBeCloseTo(before.y, 0);
    await press('ControlOrMeta+Shift+z');
    expect(outline().y).toBeCloseTo(moved.y, 0);
    const handle = document.querySelector('[data-handle="se"]')!;
    const corner = handle.getBoundingClientRect();
    const from = { x: corner.x + corner.width / 2, y: corner.y + corner.height / 2 };
    await drag(from, { x: corner.x + 50, y: corner.y + 25 }, { target: handle });
    expect(outline().width).toBeGreaterThan(moved.width);
    const finalBox = outline();
    await click(byRole('button', { name: 'Export' }));
    await click(byRole('menuitem', { name: /\(\.pptx\)/ }));
    await until(() => expect(demo.downloads).toHaveLength(1));
    const { blob, name } = demo.downloads[0];
    await openFile(new File([blob], name));
    expect(word()).toBe('#ff0000');
    expect(paintedText(/^ beta$/)?.getAttribute('fill')).not.toBe('#ff0000');
    await click(paintedText(/^Alpha$/)!);
    await press('Home');
    await shiftArrows('ArrowRight', 5);
    expect(bold()).toBe('true');
    await press('End');
    await shiftArrows('ArrowLeft', 4);
    expect(bold()).toBe('false');
    await press('Escape');
    expect(outline().x).toBeCloseTo(finalBox.x, 0);
    expect(outline().width).toBeCloseTo(finalBox.width, 0);
    expect(demo.errors).toEqual([]);
  });
});
