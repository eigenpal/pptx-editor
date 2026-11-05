import { describe, expect, test } from 'bun:test';
import {
  RAIL_FOLD_WIDTH,
  railDragAxis,
  railFoldsActions,
  railOrientationChrome,
  railParentKey,
  railStepOfKey,
  resolveRailOrientation,
  scrollIntoNearestScroller,
} from '../rail-orientation.ts';
import {
  railKeyDown,
  sectionHeaderKeyDown,
  type PaneKeyEvent,
  type SlideListKeyContext,
} from '../pane-support.ts';
import { resetDevWarnings } from '../dev-warnings.ts';
import type { SlideSection } from '../../contracts/index.ts';
import type { PptxEditorInstance } from '../pptx-editor.ts';

// The slide rail as a row (`orientation="horizontal"`): its keys, icons, drag axis, and the
// scroll that keeps the current slide in view without moving the page around the editor.

interface FakeBox {
  top: number;
  left: number;
  width: number;
  height: number;
}

interface FakeEl {
  readonly scrolled: { left: number; top: number }[];
}

/** A stand-in element: just the fields the helpers read. */
function fakeEl(opts: {
  parent?: FakeEl | null;
  box: FakeBox;
  overflow?: string;
  /** Overflow on the x axis when it differs from `overflow`. */
  overflowX?: string;
  scrollPadding?: number;
  scroll?: { width: number; height: number; clientWidth: number; clientHeight: number };
  direction?: string;
}): FakeEl {
  const scrolled: { left: number; top: number }[] = [];
  const el = {
    parentElement: opts.parent ?? null,
    clientTop: 0,
    clientLeft: 0,
    scrollWidth: opts.scroll?.width ?? opts.box.width,
    scrollHeight: opts.scroll?.height ?? opts.box.height,
    clientWidth: opts.scroll?.clientWidth ?? opts.box.width,
    clientHeight: opts.scroll?.clientHeight ?? opts.box.height,
    overflow: opts.overflow ?? 'visible',
    overflowX: opts.overflowX ?? opts.overflow ?? 'visible',
    scrollPadding: `${opts.scrollPadding ?? 0}px`,
    direction: opts.direction ?? 'ltr',
    scrolled,
    getBoundingClientRect: () => opts.box,
    scrollBy: (to: { left: number; top: number }) => void scrolled.push(to),
    ownerDocument: {
      defaultView: {
        getComputedStyle: (e: {
          overflow: string;
          overflowX: string;
          direction: string;
          scrollPadding: string;
        }) => ({
          overflowX: e.overflowX,
          overflowY: e.overflow,
          direction: e.direction,
          scrollPaddingLeft: e.scrollPadding,
          scrollPaddingRight: e.scrollPadding,
          scrollPaddingTop: 'auto',
          scrollPaddingBottom: 'auto',
        }),
      },
    },
  };
  return el;
}

describe('rail orientation', () => {
  test('a vertical rail steps with Up and Down; a horizontal one with Left and Right', () => {
    expect(railStepOfKey('ArrowUp', undefined, null)).toBe(-1);
    expect(railStepOfKey('ArrowDown', 'vertical', null)).toBe(1);
    expect(railStepOfKey('ArrowRight', 'vertical', null)).toBe(0);
    expect(railStepOfKey('ArrowLeft', 'horizontal', null)).toBe(-1);
    expect(railStepOfKey('ArrowRight', 'horizontal', null)).toBe(1);
    expect(railStepOfKey('ArrowDown', 'horizontal', null)).toBe(0);
  });

  test('a right-to-left row reads from the right', () => {
    const list = fakeEl({ box: { top: 0, left: 0, width: 100, height: 20 }, direction: 'rtl' });
    const asList = list as unknown as HTMLElement;
    expect(railStepOfKey('ArrowLeft', 'horizontal', asList)).toBe(1);
    expect(railStepOfKey('ArrowRight', 'horizontal', asList)).toBe(-1);
    expect(railStepOfKey('ArrowUp', 'vertical', asList)).toBe(-1);
  });

  test('icons, labels, drag axis, and the key to a section header follow the orientation', () => {
    expect(railOrientationChrome(undefined).moveEarlierLabel).toBe('rail.moveUp');
    expect(railOrientationChrome('horizontal')).toMatchObject({
      moveEarlierLabel: 'rail.moveLeft',
      moveLaterLabel: 'rail.moveRight',
      moveEarlierIcon: 'keyboard_arrow_left',
      collapseIcon: 'keyboard_arrow_down',
      expandIcon: 'keyboard_arrow_up',
    });
    // A right-to-left row steps earlier with Right, so its moves say so.
    expect(railOrientationChrome('horizontal', true)).toMatchObject({
      moveEarlierLabel: 'rail.moveRight',
      moveLaterLabel: 'rail.moveLeft',
      moveEarlierIcon: 'keyboard_arrow_right',
      menuMoveEarlierLabel: 'contextMenu.moveRight',
      menuMoveEarlierShortcut: 'contextMenu.shortcuts.moveRight',
      menuMoveLaterShortcut: 'contextMenu.shortcuts.moveLeft',
      collapseIcon: 'keyboard_arrow_down',
    });
    expect(railOrientationChrome('vertical', true).moveEarlierLabel).toBe('rail.moveUp');
    // A narrow row folds its secondary actions; a column, and a row not yet laid out, never do.
    expect(railFoldsActions('horizontal', 390)).toBe(true);
    expect(railFoldsActions('horizontal', RAIL_FOLD_WIDTH)).toBe(false);
    expect(railFoldsActions('horizontal', 0)).toBe(false);
    expect(railFoldsActions('vertical', 200)).toBe(false);
    expect(railDragAxis('horizontal')).toBe('x');
    expect(railDragAxis(undefined)).toBe('y');
    expect(railParentKey('horizontal')).toBe('ArrowUp');
    expect(railParentKey('vertical')).toBe('ArrowLeft');
  });

  test('railKeyDown steps along a horizontal row and leaves Up and Down alone', () => {
    const calls: number[] = [];
    const editor = {
      goToSlide: (i: number) => void calls.push(i),
    } as unknown as PptxEditorInstance;
    const target = { classList: { contains: () => false }, closest: () => null };
    const ctx: SlideListKeyContext = {
      editor,
      readOnly: false,
      selected: [1],
      current: 1,
      count: 3,
      grouped: null,
      list: null,
      orientation: 'horizontal',
      step: () => null,
      openMenu: () => {},
      focusCurrent: () => {},
    };
    const press = (key: string): boolean => {
      let prevented = false;
      const e: PaneKeyEvent = {
        key,
        metaKey: false,
        ctrlKey: false,
        shiftKey: false,
        altKey: false,
        target: target as unknown as EventTarget,
        preventDefault: () => void (prevented = true),
        stopPropagation: () => {},
      };
      railKeyDown(e, ctx);
      return prevented;
    };
    expect(press('ArrowRight')).toBe(true);
    expect(press('ArrowLeft')).toBe(true);
    expect(press('ArrowDown')).toBe(false);
    expect(calls).toEqual([2, 0]);
  });
});

describe('scrollIntoNearestScroller', () => {
  test('scrolls the nearest scrolling ancestor by the least distance, and no outer one', () => {
    const page = fakeEl({
      box: { top: 0, left: 0, width: 1000, height: 800 },
      overflow: 'auto',
      scroll: { width: 2000, height: 800, clientWidth: 1000, clientHeight: 800 },
    });
    const strip = fakeEl({
      parent: page,
      box: { top: 600, left: 500, width: 400, height: 100 },
      overflow: 'auto',
      scroll: { width: 1200, height: 100, clientWidth: 400, clientHeight: 100 },
    });
    const wrapper = fakeEl({ parent: strip, box: { top: 600, left: 820, width: 112, height: 80 } });
    const item = fakeEl({ parent: wrapper, box: { top: 610, left: 820, width: 112, height: 70 } });
    scrollIntoNearestScroller(item as unknown as HTMLElement);
    // The item's right edge is 32px past the strip's: the strip moves 32px, the page stays.
    expect(strip.scrolled).toEqual([{ left: 32, top: 0 }]);
    expect(page.scrolled).toEqual([]);
  });

  test('an item already in view scrolls nothing', () => {
    const strip = fakeEl({
      box: { top: 0, left: 0, width: 400, height: 100 },
      overflow: 'auto',
      scroll: { width: 1200, height: 100, clientWidth: 400, clientHeight: 100 },
    });
    const item = fakeEl({ parent: strip, box: { top: 10, left: 120, width: 112, height: 70 } });
    scrollIntoNearestScroller(item as unknown as HTMLElement);
    expect(strip.scrolled).toEqual([]);
  });
});

describe('scrollIntoNearestScroller limits', () => {
  test('never scrolls past the pane that bounds it', () => {
    const page = fakeEl({
      box: { top: 0, left: 0, width: 1000, height: 800 },
      overflow: 'auto',
      scroll: { width: 2000, height: 800, clientWidth: 1000, clientHeight: 800 },
    });
    // A short list that does not overflow, inside the rail pane.
    const rail = fakeEl({ parent: page, box: { top: 600, left: 900, width: 400, height: 100 } });
    const list = fakeEl({ parent: rail, box: { top: 600, left: 900, width: 400, height: 100 } });
    const item = fakeEl({ parent: list, box: { top: 610, left: 1100, width: 112, height: 70 } });
    scrollIntoNearestScroller(item as unknown as HTMLElement, rail as unknown as Element);
    expect(page.scrolled).toEqual([]);
  });

  test('scrolls only an axis that scrolls, and keeps the scroll padding', () => {
    // A row that scrolls sideways and hides its vertical overflow, with 12px of padding.
    const strip = fakeEl({
      box: { top: 0, left: 0, width: 400, height: 80 },
      overflow: 'hidden',
      overflowX: 'auto',
      scrollPadding: 12,
      scroll: { width: 1200, height: 100, clientWidth: 400, clientHeight: 80 },
    });
    const item = fakeEl({ parent: strip, box: { top: 10, left: 4, width: 112, height: 90 } });
    scrollIntoNearestScroller(item as unknown as HTMLElement);
    expect(strip.scrolled).toEqual([{ left: -8, top: 0 }]);
  });
});

describe('horizontal rail sections', () => {
  const section: SlideSection = { id: 's1', name: 'One', firstSlide: 0, slideCount: 2 };
  function header(orientation: 'vertical' | 'horizontal') {
    const rail = { getAttribute: () => orientation };
    const node = {
      closest: (sel: string) => (sel === '.pptx-rail' ? rail : null),
      querySelector: () => null,
    };
    return node as unknown as HTMLElement;
  }
  function press(node: HTMLElement, key: string, collapsed: boolean): string[] {
    const calls: string[] = [];
    const e: PaneKeyEvent = {
      key,
      metaKey: false,
      ctrlKey: false,
      shiftKey: false,
      altKey: false,
      target: node,
      preventDefault: () => {},
      stopPropagation: () => {},
    };
    const editor = { goToSlide: (i: number) => void calls.push(`go ${i}`) };
    sectionHeaderKeyDown(e, node, {
      editor: editor as unknown as PptxEditorInstance,
      readOnly: false,
      section,
      collapsed,
      toggle: () => void calls.push('toggle'),
      rename: () => {},
      openMenu: () => {},
    });
    return calls;
  }

  test('Down expands and enters a section; Up collapses it; Left and Right do not', () => {
    const row = header('horizontal');
    expect(press(row, 'ArrowDown', true)).toEqual(['toggle']);
    expect(press(row, 'ArrowDown', false)).toEqual(['go 0']);
    expect(press(row, 'ArrowUp', false)).toEqual(['toggle']);
    expect(press(row, 'ArrowLeft', false)).toEqual([]);
    expect(press(row, 'ArrowRight', true)).toEqual([]);
  });

  test('a vertical rail keeps Right and Left on its headers', () => {
    const column = header('vertical');
    expect(press(column, 'ArrowRight', true)).toEqual(['toggle']);
    expect(press(column, 'ArrowLeft', false)).toEqual(['toggle']);
    expect(press(column, 'ArrowDown', true)).toEqual([]);
  });
});

describe('resolveRailOrientation', () => {
  test('keeps the two values and warns once for anything else', () => {
    resetDevWarnings();
    const warn = console.warn;
    const messages: string[] = [];
    console.warn = (m: string) => void messages.push(m);
    try {
      expect(resolveRailOrientation('horizontal', 'Rail')).toBe('horizontal');
      expect(resolveRailOrientation(undefined, 'Rail')).toBe('vertical');
      expect(resolveRailOrientation('row', 'Rail')).toBe('vertical');
    } finally {
      console.warn = warn;
    }
    expect(messages).toHaveLength(1);
    expect(messages[0]).toContain('Rail orientation "row"');
  });
});
