// Menus, the slide show, and smart guides on the React demo in happy-dom: the show and its keys,
// snapping and its two ways off, the rail, shape, slide, and text context menus, the Insert menu
// and its submenus, a menu opened from the keyboard, and the font and size controls. Ported from
// the Playwright `e2e/interaction.spec.ts` (its narrow-window step needs CSS layout and stays in
// the browser spec).

import { afterEach, describe, expect, test } from 'bun:test';
import { within } from '@testing-library/react';
import {
  IN,
  click,
  clickAt,
  clickOn,
  clickSlide,
  closeDemo,
  drag,
  fill,
  hold,
  hover,
  isDisabled,
  moveTo,
  openDemo,
  press,
  screen,
  slidePoint,
  step,
  type,
  until,
  wordPoint,
} from './demo-harness.tsx';

afterEach(closeDemo);

const byRole = screen.getByRole;
const rail = () => [...document.querySelectorAll<HTMLElement>('.pptx-rail__item')];
const count = (selector: string) => document.querySelectorAll(selector).length;
const chip = () => document.querySelector('.pptx-chip')?.textContent;
const focusIn = (el: Element) => el.contains(document.activeElement);

describe('demo slide show', () => {
  test('present: open the show, step through the sample deck, and exit', async () => {
    const demo = await openDemo();
    const show = () => screen.queryByRole('dialog', { name: 'Slide show' });
    const position = () => show()?.querySelector('.pptx-show__position')?.textContent;

    // The top bar's Present button starts from the first slide, fitted on black.
    await click(byRole('button', { name: 'Present' }));
    expect(show()).not.toBeNull();
    expect(position()).toBe('Slide 1 of 5');
    expect(show()!.querySelectorAll('.pptx-show__slide svg')).toHaveLength(1);
    // The show fills the window; the slide box is fitted into it by the show's own sizing.
    const screenBox = { width: window.innerWidth, height: window.innerHeight };
    const slide = show()!.querySelector<HTMLElement>('.pptx-show__slide')!;
    const slideBox = {
      width: parseFloat(slide.style.width),
      height: parseFloat(slide.style.height),
    };
    expect(Math.abs(slideBox.width / slideBox.height - 16 / 9)).toBeLessThan(0.01);
    const fills =
      Math.abs(slideBox.width - screenBox.width) < 1 ||
      Math.abs(slideBox.height - screenBox.height) < 1;
    expect(fills).toBe(true);

    // Step through the deck with the keys, then the end screen.
    for (const [key, n] of [
      ['ArrowRight', 2],
      ['Space', 3],
      ['PageDown', 4],
      ['n', 5],
    ] as const) {
      await press(key);
      expect(position()).toBe(`Slide ${n} of 5`);
    }
    await press('ArrowRight');
    expect(within(show()!).queryByText('End of slide show. Click to exit.')).not.toBeNull();
    await press('ArrowLeft');
    expect(position()).toBe('Slide 5 of 5');
    await press('Home');
    expect(position()).toBe('Slide 1 of 5');
    await press('3');
    await press('Enter');
    expect(position()).toBe('Slide 3 of 5');

    // A click advances; the control bar appears on mouse move and ends the show.
    await clickOn(show()!, { x: screenBox.width / 2, y: 8 });
    expect(position()).toBe('Slide 4 of 5');
    await moveTo(show()!, { x: 40, y: screenBox.height - 40 });
    await click(within(show()!).getByRole('button', { name: 'End show' }));
    expect(show()).toBeNull();
    expect(demo.editor.snapshot().currentSlide).toBe(3);

    // The Slide Show menu presents from the current slide; Escape ends the show.
    await click(byRole('menuitem', { name: 'Slide Show' }));
    await click(byRole('menuitem', { name: 'Play from current slide' }));
    expect(position()).toBe('Slide 4 of 5');
    await press('Escape');
    expect(show()).toBeNull();

    // Viewing mode presents too: F5 from the start, Ctrl/Cmd+Enter from the current slide.
    await click(byRole('button', { name: /^Mode/ }));
    await click(byRole('menuitemradio', { name: /Viewing/ }));
    await press('F5');
    expect(position()).toBe('Slide 1 of 5');
    await press('Escape');
    expect(show()).toBeNull();
    await press('ControlOrMeta+Enter');
    expect(position()).toBe('Slide 1 of 5');
    await press('Escape');
    expect(show()).toBeNull();
    expect(demo.errors).toEqual([]);
  });
});

describe('demo smart guides', () => {
  test('smart guides: a drag snaps to an edge, Alt and View > Smart guides turn it off', async () => {
    const demo = await openDemo();
    const guides = () => count('.pptx-surface__guide');
    // The selected shape's left edge in the store, in EMU.
    const left = () => {
      const s = demo.editor.snapshot();
      const id = (s.selection as { nodeIds?: string[] }).nodeIds?.[0];
      return demo.editor
        .session!.layout(s.currentSlide)!
        .items.find((i) => i.element.nodeId === id)!.element.xfrm.x;
    };
    // On the title slide, drag the square icon (left 640080, 822960 wide) right until its right
    // edge is 3 pixels past the subtitle's right edge (6126480): it snaps to that edge, and
    // guides show while the pointer is down.
    const SNAPPED = 6_126_480 - 822_960;
    const y = 1_874_520;
    const from = slidePoint(1_051_560, y);
    const to = slidePoint(1_051_560 + SNAPPED - 640_080, y);
    const px = (to.x - from.x) / (SNAPPED - 640_080);
    let during = -1;
    await drag(
      from,
      { x: to.x + 3, y: to.y },
      { steps: 8, during: () => void (during = guides()) }
    );
    expect(during).toBeGreaterThan(0);
    expect(guides()).toBe(0);
    expect(Math.abs(left() - SNAPPED)).toBeLessThan(2);

    // With Alt down, a small move does not snap back.
    hold('Alt', true);
    await drag(to, { x: to.x + 4, y: to.y }, { steps: 4, during: () => void (during = guides()) });
    hold('Alt', false);
    expect(during).toBe(0);
    expect(left() - SNAPPED).toBeGreaterThan(3 / px);

    // View > Smart guides turns snapping off: a move to 1 pixel from the edge stays there.
    const item = () => byRole('menuitemcheckbox', { name: 'Smart guides' });
    await click(byRole('menuitem', { name: 'View' }));
    expect(item().getAttribute('aria-checked')).toBe('true');
    await click(item());
    await click(byRole('menuitem', { name: 'View' }));
    expect(item().getAttribute('aria-checked')).toBe('false');
    await press('Escape');
    await drag(
      { x: to.x + 4, y: to.y },
      { x: to.x + 1, y: to.y },
      { steps: 3, during: () => void (during = guides()) }
    );
    expect(during).toBe(0);
    expect(left() - SNAPPED).toBeGreaterThan(0.5 / px);
    expect(demo.errors).toEqual([]);
  });
});

describe('demo context menus', () => {
  test('context menus: duplicate a rail slide, bring a shape to front, and the empty-canvas menu', async () => {
    const demo = await openDemo();
    // Record whether each `contextmenu` reached the window with the browser's menu stopped.
    const stopped: boolean[] = [];
    const record = (e: Event) => stopped.push(e.defaultPrevented);
    window.addEventListener('contextmenu', record);

    // Rail: right-click the second slide, then Duplicate slide.
    await click(rail()[1], { button: 'right' });
    const railMenu = () => screen.queryByRole('menu', { name: 'Slide actions' });
    expect(railMenu()).not.toBeNull();
    expect(rail()[1].getAttribute('aria-selected')).toBe('true');
    await click(within(railMenu()!).getByRole('menuitem', { name: 'Duplicate slide' }));
    expect(rail()).toHaveLength(6);
    expect(railMenu()).toBeNull();

    // Canvas: right-click the first shape of the shapes slide, then Bring to front.
    await click(rail()[4]);
    await clickSlide(2 * IN, 2.75 * IN, { button: 'right' });
    const shapes = () => screen.queryByRole('menu', { name: 'Shape actions' });
    expect(shapes()).not.toBeNull();
    expect(count('.pptx-surface__outline')).toBe(1);
    const front = within(shapes()!).getByRole('menuitem', { name: 'Bring to front' });
    expect(front.getAttribute('aria-disabled')).not.toBe('true');
    const before = demo.editor.snapshot().revision;
    await click(front);
    expect(shapes()).toBeNull();
    expect(demo.editor.snapshot().revision).toBeGreaterThan(before);
    // Escape closes a menu opened again on the same shape.
    await clickSlide(2 * IN, 2.75 * IN, { button: 'right' });
    expect(shapes()).not.toBeNull();
    await press('Escape');
    expect(screen.queryAllByRole('menu')).toHaveLength(0);

    // Empty slide area: the custom menu shows, never the browser's.
    await clickSlide(6 * IN, 6.9 * IN, { button: 'right' });
    expect(railMenu()).not.toBeNull();
    expect(within(railMenu()!).getByRole('menuitem', { name: 'New slide' })).not.toBeNull();
    expect(count('.pptx-surface__outline')).toBe(0);
    expect(stopped.length).toBeGreaterThan(0);
    expect(stopped.every(Boolean)).toBe(true);
    await press('Escape');
    window.removeEventListener('contextmenu', record);
    expect(demo.errors).toEqual([]);
  });

  test('Insert menu: icons on every row, and Insert > Shape > Arrows inserts an arrow', async () => {
    await openDemo();
    await click(byRole('menuitem', { name: 'Insert' }));
    const insert = byRole('menu', { name: 'Insert' });
    const rows = within(insert).getAllByRole('menuitem');
    expect(rows.length).toBeGreaterThan(0);
    for (const row of rows) expect(row.querySelectorAll('.pptx-menu__icon > svg')).toHaveLength(1);
    await hover(within(insert).getByRole('menuitem', { name: 'Shape' }));
    const shapeMenu = () => byRole('menu', { name: 'Shape' });
    await hover(within(shapeMenu()).getByRole('menuitem', { name: 'Arrows' }));
    const arrows = () => screen.queryByRole('menu', { name: 'Arrows' });
    expect(arrows()).not.toBeNull();
    // ArrowLeft closes the arrow grid and gives focus back to its row.
    await step(() => within(arrows()!).getByRole('menuitem', { name: 'Right arrow' }).focus());
    await press('ArrowLeft');
    expect(arrows()).toBeNull();
    expect(document.activeElement).toBe(
      within(shapeMenu()).getByRole('menuitem', { name: 'Arrows' })
    );
    await press('ArrowRight');
    expect(arrows()).not.toBeNull();
    await click(within(arrows()!).getByRole('menuitem', { name: 'Right arrow' }));
    expect(screen.queryAllByRole('menu')).toHaveLength(0);
    expect(count('.pptx-surface__outline')).toBe(1);
    expect(chip()).toBe('Shape');
  });

  test('text context menu: right-click selected text keeps it, and Bold applies to it', async () => {
    const demo = await openDemo();
    await click(byRole('button', { name: 'Text box' }));
    await type('alpha beta');
    await press('Home');
    for (let i = 0; i < 5; i += 1) await press('Shift+ArrowRight');
    await clickAt(wordPoint('alpha'), { button: 'right' });
    const menu = () => screen.queryByRole('menu', { name: 'Text actions' });
    expect(menu()).not.toBeNull();
    expect(isDisabled(within(menu()!).getByRole('menuitem', { name: 'Copy' }))).toBe(false);
    // A two-state row is a checkbox row (`aria-checked`), as in DOCX Editor's menus.
    await click(within(menu()!).getByRole('menuitemcheckbox', { name: 'Bold' }));
    expect(menu()).toBeNull();
    expect(byRole('button', { name: /^Bold/ }).getAttribute('aria-pressed')).toBe('true');
    expect(demo.errors).toEqual([]);
  });

  test('a menu opened from the keyboard takes focus, and its checked row shows a check mark', async () => {
    const demo = await openDemo();
    await click(byRole('button', { name: 'Text box' }));
    await type('alpha beta');
    await press('ControlOrMeta+a');
    // Focus the Align trigger and open it with Enter: focus moves into the menu.
    await step(() => document.querySelector<HTMLElement>('[data-slot="paragraph.align"]')!.focus());
    await press('Enter');
    const menu = () => screen.getByRole('menu', { name: /Align/ });
    expect(focusIn(menu())).toBe(true);
    await press('ArrowDown');
    expect(focusIn(menu())).toBe(true);
    // The current alignment is the checked row, and only it shows the check mark.
    const checked = menu().querySelectorAll('[aria-checked="true"]');
    expect(checked).toHaveLength(1);
    expect(checked[0].querySelectorAll('.pptx-menu__check')).toHaveLength(1);
    expect(menu().querySelectorAll('.pptx-menu__check')).toHaveLength(1);
    await press('Escape');
    expect(screen.queryByRole('menu', { name: /Align/ })).toBeNull();
    expect(demo.errors).toEqual([]);
  });

  test('the font search takes the typed name, and the size stepper steps', async () => {
    const demo = await openDemo();
    await click(byRole('button', { name: 'Text box' }));
    await type('alpha beta');
    await press('ControlOrMeta+a');
    // The font search field takes the keys, so the slide text stays as it is.
    const family = () => document.querySelector<HTMLElement>('[data-slot="font.family"]')!;
    await click(family());
    await type('Georgia');
    const search = () => screen.queryByRole('searchbox', { name: /font/i }) as HTMLInputElement;
    expect(search().value).toBe('Georgia');
    await press('Enter');
    expect(family().textContent).toContain('Georgia');
    // The pick closes the list and gives the caret back to the slide text.
    expect(search()).toBeNull();
    expect(document.activeElement).toBe(document.querySelector('.pptx-surface__input'));
    // Enter in the size box applies the size and leaves the selected text as it is.
    const size = () => document.querySelector<HTMLInputElement>('.pptx-toolbar__size')!;
    await fill(size(), '40');
    await press('Enter');
    expect(size().value).toBe('40');
    expect(document.querySelector('.pptx-surface__slide')!.textContent).toContain('alpha beta');
    // Increase Font Size steps up from the current size.
    const before = Number(size().value);
    await click(byRole('button', { name: /Increase font size/ }));
    await until(() => expect(Number(size().value)).toBeGreaterThan(before));
    expect(demo.errors).toEqual([]);
  });
});
