// Interaction on the React demo in happy-dom: selection, groups, the clipboard, the shape
// gallery, shortcuts, links, backgrounds, paragraphs and bullets, and the format painter.
// Ported from the Playwright `e2e/interaction.spec.ts` (its zoom-scrolling step needs CSS
// layout and stays in the browser spec).

import { afterEach, describe, expect, test } from 'bun:test';
import { act } from '@testing-library/react';
import {
  IN,
  barButton,
  click,
  clickAt,
  clickSlide,
  dblclick,
  closeDemo,
  dblclickSlide,
  drag,
  fill,
  openDemo,
  press,
  screen,
  slidePoint,
  type,
  until,
  wordPoint,
} from './demo-harness.tsx';

afterEach(closeDemo);

const byRole = screen.getByRole;
const rail = () => [...document.querySelectorAll<HTMLElement>('.pptx-rail__item')];
const count = (selector: string) => document.querySelectorAll(selector).length;
const chip = () => document.querySelector('.pptx-chip')?.textContent;
const shown = (text: string) => screen.queryAllByText(text).length > 0;

describe('demo interaction', () => {
  test('marquee, group, enter a group, and copy and paste', async () => {
    const demo = await openDemo();
    await click(rail()[3]);
    const paths = () => count('.pptx-surface__slide path');

    // Drag a marquee around the top row of four shapes.
    await drag(slidePoint(0.7 * IN, 1.6 * IN), slidePoint(12.6 * IN, 3.9 * IN), { steps: 6 });
    expect(count('.pptx-surface__outline')).toBe(4);
    expect(shown('4 shapes selected')).toBe(true);

    // Group them; the group is selected as one element.
    await press('ControlOrMeta+g');
    expect(count('.pptx-surface__outline')).toBe(1);
    expect(chip()).toBe('Group');

    // Double-click enters the group and selects the ellipse; drag it down inside the group.
    const ellipse = slidePoint(7.1 * IN, 2.75 * IN);
    await dblclickSlide(7.1 * IN, 2.75 * IN);
    expect(chip()).toBe('Shape');
    expect(count('.pptx-surface__group-outline')).toBe(1);
    await drag(ellipse, { x: ellipse.x, y: ellipse.y + 60 });

    // Escape leaves the group with it selected; copy and paste it through the clipboard.
    await press('Escape');
    expect(chip()).toBe('Group');
    const before = paths();
    await press('ControlOrMeta+c');
    await press('ControlOrMeta+v');
    await until(() => expect(paths()).toBeGreaterThan(before));

    // Ungroup the pasted copy into its four shapes; undo brings the group back.
    await press('ControlOrMeta+Shift+g');
    expect(shown('4 shapes selected')).toBe(true);
    await press('ControlOrMeta+z');
    expect(chip()).toBe('Group');
    expect(demo.errors).toEqual([]);
  });

  test('shape gallery, gradient fill, lines with ends, and vertical align', async () => {
    const demo = await openDemo();

    // Insert a star from the gallery and give it a gradient preset from the theme colors.
    await click(byRole('button', { name: 'Shape' }));
    await click(byRole('menuitem', { name: 'Star: 5 points' }));
    expect(chip()).toBe('Shape');
    const gradients = () => count('.pptx-surface__slide linearGradient');
    const before = gradients();
    await click(await barButton('Fill color', 1));
    await click(byRole('menuitem', { name: 'Gradient from Accent 1' }));
    await until(() => expect(gradients()).toBeGreaterThan(before));

    // Insert an arrow; its start becomes a circle.
    await click(byRole('button', { name: 'Line' }));
    await click(byRole('menuitem', { name: 'Arrow' }));
    await click(await barButton('Line start'));
    await click(byRole('menuitemradio', { name: 'Circle' }));
    await click(await barButton('Line start'));
    expect(byRole('menuitemradio', { name: 'Circle' }).getAttribute('aria-checked')).toBe('true');
    await press('Escape');

    // Vertical alignment of a text box.
    await click(byRole('button', { name: 'Text box' }));
    await type('Middle');
    await click(await barButton('Vertical align'));
    await click(byRole('menuitemradio', { name: 'Middle' }));
    await click(await barButton('Vertical align'));
    expect(byRole('menuitemradio', { name: 'Middle' }).getAttribute('aria-checked')).toBe('true');
    expect(demo.errors).toEqual([]);
  });

  test('rail and shape shortcuts, line ends, links, and backgrounds', async () => {
    const demo = await openDemo();

    // In the rail, Ctrl/Cmd+D duplicates the slide and Backspace deletes it.
    await click(rail()[1]);
    await press('ControlOrMeta+d');
    expect(rail()).toHaveLength(6);
    await press('Backspace');
    expect(rail()).toHaveLength(5);

    // On the canvas, Ctrl/Cmd+D duplicates the selected shape.
    await click(rail()[3]);
    const shapes = () => count('.pptx-surface__slide [data-kind="shape"]');
    const before = shapes();
    await clickSlide(7.95 * IN, 5.15 * IN);
    await press('ControlOrMeta+d');
    await until(() => expect(shapes()).toBe(before + 1));
    await press('Escape');

    // A line shows two end points; dragging one moves only that end.
    await click(byRole('button', { name: 'Line' }));
    await click(byRole('menuitem', { name: 'Line' }));
    expect(count('[data-handle="start"], [data-handle="end"]')).toBe(2);
    const handle = document.querySelector('[data-handle="end"]')!;
    const end = handle.getBoundingClientRect();
    await drag(
      { x: end.x + end.width / 2, y: end.y + end.height / 2 },
      { x: end.x - 60, y: end.y + 80 },
      { target: handle }
    );
    expect((screen.getByLabelText('Height', { exact: true }) as HTMLInputElement).value).not.toBe(
      '0.00'
    );

    // Links: Ctrl/Cmd+K opens the popover over selected text.
    await click(rail()[0]);
    const subtitle = slidePoint(1_000_000, 3_900_000);
    await clickAt(subtitle);
    await press('Shift+End');
    await press('ControlOrMeta+k');
    await fill(byRole('textbox', { name: 'Link' }), 'example.com');
    await click(byRole('button', { name: 'Apply' }));
    // A click back inside the linked text (it starts at the first click) opens the link panel.
    await clickAt({ x: subtitle.x + 80, y: subtitle.y });
    expect(document.querySelector('a.pptx-link__url')?.getAttribute('href')).toBe(
      'https://example.com'
    );
    // Ctrl/Cmd+K there edits the link.
    await press('ControlOrMeta+k');
    expect((byRole('textbox', { name: 'Link' }) as HTMLInputElement).value).toBe(
      'https://example.com'
    );
    await press('Escape');

    // Background: a color for this slide, from Format > Slide background (a standard color).
    await click(byRole('menuitem', { name: 'Format' }));
    await click(byRole('menuitem', { name: 'Slide background…' }));
    const dialog = byRole('dialog', { name: 'Format Background' });
    const darkBlue = [...dialog.querySelectorAll<HTMLElement>('button')].find(
      (b) => b.getAttribute('aria-label') === 'Dark Blue'
    )!;
    await click(darkBlue);
    await click(
      [...dialog.querySelectorAll<HTMLElement>('button')].find((b) => b.textContent === 'OK')!
    );
    expect(demo.editor.snapshot().slideBackground.color).toBe('#002060');
    expect(demo.errors).toEqual([]);
  });

  test('Enter starts paragraphs, bullets stay on their paragraph, and Tab sets the level', async () => {
    await openDemo();
    await click(byRole('button', { name: 'Text box' }));
    for (const [i, line] of ['one', 'two', 'three'].entries()) {
      if (i > 0) await press('Enter');
      await type(line);
    }
    await press('ArrowUp');
    await click(byRole('button', { name: 'Bulleted list' }));
    // Painted text on the slide, as [content, left, top].
    const painted = () =>
      [...document.querySelectorAll('.pptx-surface__slide text')].map((t) => {
        const r = t.getBoundingClientRect();
        return [t.textContent ?? '', Math.round(r.left), Math.round(r.top)] as const;
      });
    const at = (texts: readonly (readonly [string, number, number])[], s: string) =>
      texts.find((t) => t[0] === s)!;
    let texts = painted();
    const bullets = texts.filter((t) => t[0] === '•');
    expect(bullets).toHaveLength(1);
    expect(Math.abs(bullets[0][2] - at(texts, 'two')[2])).toBeLessThan(4);
    expect(at(texts, 'one')[1]).toBeLessThan(at(texts, 'two')[1]);

    await press('End');
    await press('Enter');
    await type('four');
    texts = painted();
    expect(texts.filter((t) => t[0] === '•')).toHaveLength(2);
    const x = at(texts, 'four')[1];
    await press('Tab');
    expect(at(painted(), 'four')[1]).toBeGreaterThan(x);
    await press('Shift+Tab');
    expect(at(painted(), 'four')[1]).toBe(x);
  });

  test('format painter: bold a word, paint another word, and keep the painter on', async () => {
    const demo = await openDemo({ width: 1920, height: 1000 });
    await click(byRole('button', { name: 'Text box' }));
    await type('alpha beta gamma');
    await press('Home');
    for (let i = 0; i < 5; i += 1) await press('Shift+ArrowRight');
    await press('ControlOrMeta+b');

    // One click on the button arms the painter; the slide shows the brush.
    const painter = () => byRole('button', { name: /Format painter/ });
    const bold = () => byRole('button', { name: /^Bold/ });
    const surface = () => document.querySelector('.pptx-surface')!;
    await click(painter());
    expect(painter().getAttribute('aria-pressed')).toBe('true');
    expect(surface().classList.contains('pptx-surface--painting')).toBe(true);

    // A click on "gamma" paints the word and turns the painter off.
    await clickAt(wordPoint('gamma'));
    expect(painter().getAttribute('aria-pressed')).toBe('false');
    expect(surface().classList.contains('pptx-surface--painting')).toBe(false);
    expect(bold().getAttribute('aria-pressed')).toBe('true');
    const isBold = (from: number, to: number) => {
      const nodeId = (demo.editor.snapshot().selection as { nodeId?: string }).nodeId!;
      act(() =>
        demo.editor.setSelection({
          kind: 'text',
          nodeId,
          anchor: { paragraph: 0, offset: from },
          focus: { paragraph: 0, offset: to },
        })
      );
      return demo.editor.snapshot().format?.bold;
    };
    expect(isBold(11, 16)).toBe(true);
    expect(isBold(6, 10)).toBe(false);
    await press('ControlOrMeta+z');
    expect(isBold(11, 16)).toBe(false);
    expect(isBold(0, 5)).toBe(true);

    // A double-click keeps it on until Escape.
    await dblclick(painter());
    expect(painter().getAttribute('aria-pressed')).toBe('true');
    await clickAt(wordPoint('beta'));
    expect(painter().getAttribute('aria-pressed')).toBe('true');
    expect(isBold(6, 10)).toBe(true);
    await press('Escape');
    expect(painter().getAttribute('aria-pressed')).toBe('false');
    expect(surface().classList.contains('pptx-surface--painting')).toBe(false);
    expect(demo.errors).toEqual([]);
  });
});
