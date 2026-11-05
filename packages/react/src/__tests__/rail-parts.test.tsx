import { afterEach, describe, expect, test } from 'bun:test';
import { act, cleanup, fireEvent, render, screen } from '@testing-library/react';
import type { ReactNode } from 'react';
import type { PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import { buildPptx } from '../../../../scripts/fixtures/builder.ts';
import { PptxEditor } from '../index.ts';

// The named parts of the slide rail's action row.

afterEach(() => cleanup());

const deck = () => buildPptx({ slides: [{ shapes: '' }, { shapes: '' }] });

async function mount(rail: ReactNode) {
  let editor: PptxEditorInstance | null = null;
  render(<PptxEditor.Root onReady={(e) => (editor = e)}>{rail}</PptxEditor.Root>);
  await act(async () => void (await editor!.load(deck())));
  return editor!;
}

const actionClasses = () =>
  [...document.querySelectorAll<HTMLElement>('.pptx-rail__actions > button')].map(
    (b) => b.className.replace('pptx-toolbar__button ', '').split(' ')[0]
  );

const columns = () =>
  document
    .querySelector<HTMLElement>('.pptx-rail__actions')!
    .style.getPropertyValue('--pptx-rail-action-columns');

describe('SlideRail parts', () => {
  test('the default rail shows every action in two rows', async () => {
    await mount(<PptxEditor.SlideRail />);
    expect(actionClasses()).toEqual([
      'pptx-rail__add',
      'pptx-rail__layouts',
      'pptx-rail__duplicate',
      'pptx-rail__move',
      'pptx-rail__move',
      'pptx-rail__delete',
      'pptx-rail__collapse',
    ]);
    expect(columns()).toBe('4');
  });

  test('hidden removes a part, className and icon change one, and the columns follow', async () => {
    await mount(
      <PptxEditor.SlideRail orientation="horizontal">
        <PptxEditor.SlideRail.MoveEarlier hidden />
        <PptxEditor.SlideRail.MoveLater hidden />
        <PptxEditor.SlideRail.AddSlide
          className="host-add"
          icon={<svg data-testid="host-icon" />}
        />
      </PptxEditor.SlideRail>
    );
    expect(document.querySelector('.pptx-rail__move')).toBeNull();
    const add = screen.getByRole('button', { name: 'New slide' });
    expect(add.classList.contains('pptx-rail__add')).toBe(true);
    expect(add.classList.contains('host-add')).toBe(true);
    expect(add.querySelector('[data-testid="host-icon"]') !== null).toBe(true);
    expect(add.querySelector('.pptx-icon')).toBeNull();
    expect(columns()).toBe('3');
  });

  test('a host action goes after Delete and before Hide thumbnails', async () => {
    let sent = 0;
    await mount(
      <PptxEditor.SlideRail>
        <PptxEditor.SlideRail.Action label="Send" onSelect={() => sent++} />
      </PptxEditor.SlideRail>
    );
    const buttons = [...document.querySelectorAll<HTMLElement>('.pptx-rail__actions > button')];
    const labels = buttons.map((b) => b.getAttribute('aria-label'));
    expect(labels.indexOf('Send')).toBe(labels.indexOf('Delete slide') + 1);
    expect(labels.at(-1)).toBe('Hide thumbnails');
    act(() => void fireEvent.click(screen.getByRole('button', { name: 'Send' })));
    expect(sent).toBe(1);
  });

  test('preset={false} renders only the children, and the parts still work', async () => {
    const editor = await mount(
      <PptxEditor.SlideRail preset={false}>
        <PptxEditor.SlideRail.Delete />
        <PptxEditor.SlideRail.AddSlide />
      </PptxEditor.SlideRail>
    );
    expect(actionClasses()).toEqual(['pptx-rail__delete', 'pptx-rail__add']);
    act(() => void fireEvent.click(screen.getByRole('button', { name: 'New slide' })));
    expect(editor.snapshot().slides.length).toBe(3);
  });

  test('asChild hands the part to the host element: label, state, reason, and click', async () => {
    const editor = await mount(
      <PptxEditor.SlideRail>
        <PptxEditor.SlideRail.AddSlide asChild>
          <button className="host-add">Add</button>
        </PptxEditor.SlideRail.AddSlide>
        <PptxEditor.SlideRail.MoveEarlier asChild>
          <span className="host-earlier">Up</span>
        </PptxEditor.SlideRail.MoveEarlier>
      </PptxEditor.SlideRail>
    );
    const add = document.querySelector<HTMLElement>('.host-add')!;
    expect(add.tagName).toBe('BUTTON');
    expect(add.getAttribute('aria-label')).toBe('New slide');
    expect(add.classList.contains('pptx-rail__add')).toBe(true);
    expect(add.querySelector('.pptx-icon')).toBeNull();
    // The first slide cannot move earlier: the host element says so, and why.
    const earlier = document.querySelector<HTMLElement>('.host-earlier')!;
    expect(earlier.getAttribute('role')).toBe('button');
    expect(earlier.getAttribute('aria-disabled')).toBe('true');
    const reason = document.getElementById(earlier.getAttribute('aria-describedby')!);
    expect(reason?.textContent?.length ?? 0).toBeGreaterThan(0);
    act(() => void fireEvent.click(add));
    expect(editor.snapshot().slides.length).toBe(3);
  });

  test('asChild without one element child renders the packaged button', async () => {
    await mount(
      <PptxEditor.SlideRail>
        <PptxEditor.SlideRail.Delete asChild>
          <span>a</span>
          <span>b</span>
        </PptxEditor.SlideRail.Delete>
      </PptxEditor.SlideRail>
    );
    const del = screen.getByRole('button', { name: 'Delete slide' });
    expect(del.tagName).toBe('BUTTON');
    expect(del.querySelector('.pptx-icon') !== null).toBe(true);
  });

  test('a part outside a rail renders nothing', () => {
    render(<PptxEditor.SlideRail.AddSlide />);
    expect(document.querySelector('button')).toBeNull();
  });

  test('a horizontal rail with toggle={false} keeps its six actions in two rows of three', async () => {
    await mount(<PptxEditor.SlideRail orientation="horizontal" toggle={false} />);
    expect(actionClasses()).not.toContain('pptx-rail__collapse');
    expect(columns()).toBe('3');
  });

  test('toggle={false} removes Collapse under preset={false} too', async () => {
    await mount(
      <PptxEditor.SlideRail preset={false} toggle={false}>
        <PptxEditor.SlideRail.AddSlide />
        <PptxEditor.SlideRail.Collapse />
      </PptxEditor.SlideRail>
    );
    expect(actionClasses()).toEqual(['pptx-rail__add']);
    expect(columns()).toBe('1');
  });

  test('SlideRail.Expand changes the button of the folded strip', async () => {
    await mount(
      <PptxEditor.SlideRail defaultOpen={false}>
        <PptxEditor.SlideRail.Expand className="host-expand" icon={<svg className="host-icon" />} />
      </PptxEditor.SlideRail>
    );
    const expand = screen.getByRole('button', { name: 'Show thumbnails' });
    expect(expand.classList.contains('host-expand')).toBe(true);
    expect(expand.querySelector('.host-icon') !== null).toBe(true);
    act(() => void fireEvent.click(expand));
    expect(document.querySelector('.pptx-rail')!.getAttribute('data-open')).toBe('true');
    // Open, the part renders nothing and does not count as an action.
    expect(document.querySelector('.host-expand')).toBeNull();
    expect(actionClasses().length).toBe(7);
  });

  test('a hidden SlideRail.Expand leaves the folded strip with no button', async () => {
    await mount(
      <PptxEditor.SlideRail defaultOpen={false}>
        <PptxEditor.SlideRail.Expand hidden />
      </PptxEditor.SlideRail>
    );
    expect(document.querySelector('.pptx-rail--collapsed button')).toBeNull();
  });

  test('the rail keeps the current slide in view when it resizes', async () => {
    const observers: { target: Element | null; run: () => void; off: boolean }[] = [];
    const saved = globalThis.ResizeObserver;
    globalThis.ResizeObserver = class {
      readonly entry: { target: Element | null; run: () => void; off: boolean };
      constructor(run: () => void) {
        this.entry = { target: null, run, off: false };
        observers.push(this.entry);
      }
      observe(target: Element) {
        this.entry.target = target;
      }
      unobserve() {}
      disconnect() {
        this.entry.off = true;
      }
    } as unknown as typeof ResizeObserver;
    try {
      const editor = await mount(<PptxEditor.SlideRail orientation="horizontal" />);
      const list = document.querySelector<HTMLElement>('.pptx-rail__list')!;
      const watcher = observers.find((o) => o.target === list && !o.off)!;
      expect(watcher).toBeDefined();
      act(() => editor.goToSlide(1));
      // The list scrolls on its own axis: 200px wide, slide 2 starts at 300px.
      const style = document.createElement('style');
      style.textContent = '.pptx-rail__list { overflow-x: auto; }';
      document.head.append(style);
      const item = document.querySelector<HTMLElement>('.pptx-rail__item[data-index="1"]')!;
      const scrolled: number[] = [];
      Object.defineProperties(list, {
        scrollWidth: { configurable: true, value: 600 },
        clientWidth: { configurable: true, value: 200 },
        getBoundingClientRect: {
          configurable: true,
          value: () => ({ top: 0, left: 0, width: 200, height: 100 }) as DOMRect,
        },
        scrollBy: {
          configurable: true,
          value: (to: ScrollToOptions) => scrolled.push(to.left ?? 0),
        },
      });
      Object.defineProperty(item, 'getBoundingClientRect', {
        configurable: true,
        value: () => ({ top: 0, left: 300, width: 100, height: 80 }) as DOMRect,
      });
      watcher.run();
      style.remove();
      expect(scrolled).toEqual([200]);
      cleanup();
      expect(watcher.off).toBe(true);
    } finally {
      globalThis.ResizeObserver = saved;
    }
  });
});

describe('the move rows and buttons follow the rail direction', () => {
  // The Vue twin is packages/vue/test/rail-parts.test.ts ('the slide menu of a row rail').
  const rows = (): string[] =>
    ['slide.moveUp', 'slide.moveDown'].map(
      (slot) => document.querySelector(`[data-slot="${slot}"]`)?.textContent ?? ''
    );
  const buttons = (): string[] =>
    [...document.querySelectorAll<HTMLElement>('.pptx-rail__actions .pptx-rail__move')].map(
      (b) => b.getAttribute('aria-label') ?? ''
    );
  const icons = (): string[] =>
    [...document.querySelectorAll<HTMLElement>('.pptx-rail__actions .pptx-rail__move svg')].map(
      (svg) => svg.innerHTML
    );
  const openMenu = (): void =>
    void fireEvent.contextMenu(document.querySelector('.pptx-rail__item[data-index="0"]')!, {
      clientX: 5,
      clientY: 5,
    });

  for (const [orientation, direction, earlier, later, earlierButton, laterButton] of [
    ['horizontal', 'ltr', /Move left.*←/, /Move right.*→/, 'Move slide left', 'Move slide right'],
    ['horizontal', 'rtl', /Move right.*→/, /Move left.*←/, 'Move slide right', 'Move slide left'],
    ['vertical', 'ltr', /Move up.*↑/, /Move down.*↓/, 'Move slide up', 'Move slide down'],
    ['vertical', 'rtl', /Move up.*↑/, /Move down.*↓/, 'Move slide up', 'Move slide down'],
  ] as const) {
    test(`a ${orientation} ${direction} rail names its moves by its direction`, async () => {
      await mount(
        <div style={{ direction }}>
          <PptxEditor.SlideRail orientation={orientation} />
        </div>
      );
      expect(buttons()).toEqual([earlierButton, laterButton]);
      openMenu();
      await act(async () => void (await new Promise((r) => setTimeout(r, 0))));
      const [first, second] = rows();
      expect(first).toMatch(earlier);
      expect(second).toMatch(later);
    });
  }

  test('a right-to-left row swaps the move icons', async () => {
    await mount(<PptxEditor.SlideRail orientation="horizontal" />);
    const ltr = icons();
    cleanup();
    await mount(
      <div style={{ direction: 'rtl' }}>
        <PptxEditor.SlideRail orientation="horizontal" />
      </div>
    );
    expect(icons()).toEqual([ltr[1]!, ltr[0]!]);
  });
});

describe('a narrow rail folds its secondary actions into one menu', () => {
  // The Vue twin is packages/vue/test/rail-parts.test.ts ('a narrow rail folds').
  let restore: (() => void) | null = null;
  /** Every `.pptx-rail` measures `width` pixels wide. */
  const railWidth = (width: number): void => {
    const saved = Object.getOwnPropertyDescriptor(HTMLElement.prototype, 'clientWidth');
    Object.defineProperty(HTMLElement.prototype, 'clientWidth', {
      configurable: true,
      get(this: HTMLElement) {
        return this.classList.contains('pptx-rail') ? width : 0;
      },
    });
    restore = () => {
      if (saved) Object.defineProperty(HTMLElement.prototype, 'clientWidth', saved);
      else delete (HTMLElement.prototype as { clientWidth?: number }).clientWidth;
    };
  };
  afterEach(() => {
    restore?.();
    restore = null;
  });
  const rowLabels = (): string[] =>
    [...document.querySelectorAll('.pptx-contextmenu [role="menuitem"]')].map(
      (row) => row.querySelector('.pptx-menu__label')?.textContent ?? ''
    );

  test('at 390px, New slide and More stay; the menu holds the rest and runs them', async () => {
    railWidth(390);
    const editor = await mount(<PptxEditor.SlideRail orientation="horizontal" />);
    expect(actionClasses()).toEqual(['pptx-rail__add', 'pptx-rail__more', 'pptx-rail__collapse']);
    const more = screen.getByRole('button', { name: 'More slide actions' });
    expect(more.getAttribute('aria-haspopup')).toBe('menu');
    act(() => void fireEvent.click(more));
    expect(rowLabels()).toEqual([
      'Add slide with layout',
      'Duplicate slide',
      'Move slide left',
      'Move slide right',
      'Delete slide',
    ]);
    expect(more.getAttribute('aria-expanded')).toBe('true');
    act(() => void fireEvent.click(screen.getByRole('menuitem', { name: 'Duplicate slide' })));
    expect(editor.snapshot().slides.length).toBe(3);
    expect(document.querySelector('.pptx-contextmenu')).toBeNull();
  });

  test('a hidden part stays hidden, and with every secondary part hidden no More shows', async () => {
    railWidth(390);
    await mount(
      <PptxEditor.SlideRail orientation="horizontal">
        <PptxEditor.SlideRail.Duplicate hidden />
      </PptxEditor.SlideRail>
    );
    act(() => void fireEvent.click(screen.getByRole('button', { name: 'More slide actions' })));
    expect(rowLabels()).not.toContain('Duplicate slide');
    expect(rowLabels().length).toBe(4);
    cleanup();
    await mount(
      <PptxEditor.SlideRail orientation="horizontal">
        <PptxEditor.SlideRail.ChooseLayout hidden />
        <PptxEditor.SlideRail.Duplicate hidden />
        <PptxEditor.SlideRail.MoveEarlier hidden />
        <PptxEditor.SlideRail.MoveLater hidden />
        <PptxEditor.SlideRail.Delete hidden />
      </PptxEditor.SlideRail>
    );
    expect(actionClasses()).toEqual(['pptx-rail__add', 'pptx-rail__collapse']);
  });

  test('a wide rail and a vertical rail keep every button', async () => {
    railWidth(800);
    await mount(<PptxEditor.SlideRail orientation="horizontal" />);
    expect(document.querySelector('.pptx-rail__more')).toBeNull();
    expect(actionClasses().length).toBe(7);
    cleanup();
    restore?.();
    railWidth(200);
    await mount(<PptxEditor.SlideRail />);
    expect(document.querySelector('.pptx-rail__more')).toBeNull();
  });
});
