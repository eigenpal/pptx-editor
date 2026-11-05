import { describe, expect, test } from 'bun:test';
import { buildPptx, para, run, shape } from '../../../../../scripts/fixtures/builder.ts';
import { createPptxEditor } from '../pptx-editor.ts';
import {
  SlideshowController,
  slideshowShortcut,
  type SlideshowSlide,
  type SlideshowState,
} from '../slideshow.ts';
import { fitSlideToScreen, slideshowLinkAt } from '../slideshow-view.ts';

// The slide show controller: navigation, hidden slides, bounds, the end screen, blank screens,
// typed slide numbers, and the start shortcuts. Then the screen fit and the hyperlink hit test.

const deck = (...hidden: boolean[]): SlideshowSlide[] => hidden.map((h) => ({ hidden: h }));
const at = (s: SlideshowState): [number, string] => [s.index, s.screen];

describe('SlideshowController', () => {
  test('steps forward and back, then shows the end screen, then ends', () => {
    const show = new SlideshowController(deck(false, false, false));
    expect(at(show.state)).toEqual([0, 'slide']);
    show.next();
    show.next();
    expect(at(show.state)).toEqual([2, 'slide']);
    show.next();
    expect(at(show.state)).toEqual([2, 'end']);
    show.previous();
    expect(at(show.state)).toEqual([2, 'slide']);
    show.previous();
    show.previous();
    show.previous();
    expect(at(show.state)).toEqual([0, 'slide']);
    show.last();
    show.next();
    expect(show.state.done).toBe(false);
    show.next();
    expect(show.state.done).toBe(true);
    // An ended show ignores input.
    show.previous();
    expect(show.state).toMatchObject({ index: 2, done: true });
  });

  test('skips hidden slides, but shows one asked for by number or as the current slide', () => {
    const show = new SlideshowController(deck(true, false, true, false, true));
    expect(at(show.state)).toEqual([1, 'slide']);
    show.next();
    expect(at(show.state)).toEqual([3, 'slide']);
    show.next();
    expect(at(show.state)).toEqual([3, 'end']);
    show.first();
    expect(show.state.index).toBe(1);
    show.last();
    expect(show.state.index).toBe(3);
    show.goTo(2);
    expect(at(show.state)).toEqual([2, 'slide']);
    show.previous();
    expect(show.state.index).toBe(1);
    expect(new SlideshowController(deck(true, false), 0).state.index).toBe(0);
  });

  test('clamps out-of-range starts and slide numbers', () => {
    expect(new SlideshowController(deck(false, false), 9).state.index).toBe(1);
    expect(new SlideshowController(deck(false, false), -3).state.index).toBe(0);
    const show = new SlideshowController(deck(false, false, false));
    show.goTo(99);
    expect(show.state.index).toBe(2);
    show.goTo(Number.NaN);
    expect(show.state.index).toBe(2);
    show.goTo(-1);
    expect(show.state.index).toBe(0);
  });

  test('a deck with no shown slide starts on the end screen', () => {
    expect(at(new SlideshowController(deck(true, true)).state)).toEqual([0, 'end']);
    const empty = new SlideshowController([]);
    expect(empty.state.screen).toBe('end');
    empty.goTo(0);
    empty.next();
    expect(empty.state.done).toBe(true);
  });

  test('B and W toggle blank screens; an advance returns to the same slide', () => {
    const show = new SlideshowController(deck(false, false));
    expect(show.key({ key: 'b' })).toBe(true);
    expect(at(show.state)).toEqual([0, 'black']);
    show.key({ key: 'W' });
    expect(at(show.state)).toEqual([0, 'white']);
    show.key({ key: 'w' });
    expect(at(show.state)).toEqual([0, 'slide']);
    show.key({ key: '.' });
    show.key({ key: ' ' });
    expect(at(show.state)).toEqual([0, 'slide']);
  });

  test('maps the slide show keys', () => {
    const show = new SlideshowController(deck(false, false, false, false));
    for (const key of ['ArrowRight', 'ArrowDown', 'PageDown']) {
      show.first();
      show.key({ key });
      expect(show.state.index).toBe(1);
    }
    for (const key of ['Enter', ' ', 'n', 'N']) {
      show.first();
      show.key({ key });
      expect(show.state.index).toBe(1);
    }
    for (const key of ['ArrowLeft', 'ArrowUp', 'PageUp', 'Backspace', 'p', 'P']) {
      show.last();
      show.key({ key });
      expect(show.state.index).toBe(2);
    }
    show.key({ key: 'End' });
    expect(show.state.index).toBe(3);
    show.key({ key: 'Home' });
    expect(show.state.index).toBe(0);
    // Keys with Ctrl, Cmd, or Alt belong to the host; unknown keys are not used.
    expect(show.key({ key: 'ArrowRight', ctrlKey: true })).toBe(false);
    expect(show.key({ key: 'ArrowRight', metaKey: true })).toBe(false);
    expect(show.key({ key: 'x' })).toBe(false);
    expect(show.state.index).toBe(0);
    expect(show.key({ key: 'Escape' })).toBe(true);
    expect(show.state.done).toBe(true);
  });

  test('a number, then Enter, goes to that slide', () => {
    const show = new SlideshowController(deck(false, true, false, false));
    show.key({ key: '2' });
    expect(show.state.typed).toBe('2');
    show.key({ key: 'Enter' });
    expect(at(show.state)).toEqual([1, 'slide']);
    expect(show.state.typed).toBe('');
    show.key({ key: '9' });
    show.key({ key: '9' });
    show.key({ key: 'Enter' });
    expect(show.state.index).toBe(3);
    // Zero goes nowhere; another key drops the typed number.
    show.key({ key: '0' });
    show.key({ key: 'Enter' });
    expect(show.state.index).toBe(3);
    show.key({ key: '1' });
    show.key({ key: 'x' });
    expect(show.state.typed).toBe('');
    show.key({ key: 'Enter' });
    expect(show.state.screen).toBe('end');
  });

  test('notifies listeners and keeps the state object until a change', () => {
    const show = new SlideshowController(deck(false, false));
    let calls = 0;
    const off = show.subscribe(() => (calls += 1));
    const before = show.state;
    expect(show.state).toBe(before);
    show.previous();
    expect(calls).toBe(0);
    show.next();
    expect(calls).toBe(1);
    expect(show.state).not.toBe(before);
    off();
    show.previous();
    expect(calls).toBe(1);
  });

  test('follows deck changes while the show runs', () => {
    const show = new SlideshowController(deck(false, false, false), 2);
    show.update(deck(false, false));
    expect(at(show.state)).toEqual([1, 'slide']);
    show.update(deck(false, true));
    show.previous();
    expect(show.state.index).toBe(0);
    show.next();
    expect(at(show.state)).toEqual([0, 'end']);
    show.update([]);
    expect(show.state.screen).toBe('end');
  });
});

describe('slideshowShortcut', () => {
  test('F5 and Ctrl/Cmd+Shift+Enter start from the start; Shift+F5 and Ctrl/Cmd+Enter from the current slide', () => {
    expect(slideshowShortcut({ key: 'F5' })).toBe('start');
    expect(slideshowShortcut({ key: 'F5', shiftKey: true })).toBe('current');
    expect(slideshowShortcut({ key: 'Enter', ctrlKey: true })).toBe('current');
    expect(slideshowShortcut({ key: 'Enter', metaKey: true })).toBe('current');
    expect(slideshowShortcut({ key: 'Enter', metaKey: true, shiftKey: true })).toBe('start');
    expect(slideshowShortcut({ key: 'Enter' })).toBeNull();
    expect(slideshowShortcut({ key: 'F5', ctrlKey: true })).toBeNull();
    expect(slideshowShortcut({ key: 'Enter', ctrlKey: true, altKey: true })).toBeNull();
  });
});

describe('fitSlideToScreen', () => {
  test('letterboxes a 16:9 slide on a 4:3 screen and pillarboxes it on a tall one', () => {
    expect(fitSlideToScreen(1024, 768, 12_192_000, 6_858_000)).toEqual({
      left: 0,
      top: 96,
      width: 1024,
      height: 576,
    });
    const tall = fitSlideToScreen(1000, 400, 9_144_000, 6_858_000);
    expect(tall.height).toBe(400);
    expect(tall.width).toBeCloseTo(533, 0);
    expect(tall.left).toBe(233);
  });

  test('puts the edges on whole device pixels', () => {
    const fit = fitSlideToScreen(1001, 700, 12_192_000, 6_858_000, 2);
    for (const v of [fit.left, fit.top, fit.width, fit.height]) expect((v * 2) % 1).toBe(0);
    expect(fit.width).toBeLessThanOrEqual(1001);
    expect(fit.height).toBeLessThanOrEqual(700);
  });

  test('returns an empty box for an empty screen or slide', () => {
    expect(fitSlideToScreen(0, 700, 1, 1)).toEqual({ left: 0, top: 0, width: 0, height: 0 });
    expect(fitSlideToScreen(800, 600, 0, 1).width).toBe(0);
  });
});

describe('slideshowLinkAt', () => {
  const IN = 914_400;
  const HLINK = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink';
  const linked = (target: string) =>
    buildPptx({
      slides: [
        {
          layout: 'blank',
          shapes: shape({
            id: 2,
            x: IN,
            y: IN,
            cx: 4 * IN,
            cy: IN,
            text: para(run('link', '', '<a:hlinkClick r:id="rId9"/>')),
          }),
          // The builder writes targets verbatim; this one closes the attribute to add TargetMode.
          rels: [['rId9', HLINK, `${target}" TargetMode="External`]],
        },
      ],
    });
  const open = async (bytes: Uint8Array) => {
    const editor = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    expect((await editor.load(bytes)).ok).toBe(true);
    return editor.session!;
  };
  const textPoint = (session: Awaited<ReturnType<typeof open>>) => {
    const item = session.layout(0)!.items[0];
    if (item.kind !== 'shape' || !item.text) throw new Error('no text');
    const line = item.text.lines[0];
    const glyphs = line.runs.find((r) => r.kind === 'text')!;
    const mid = (glyphs.x[0] + glyphs.x[glyphs.x.length - 1]) / 2;
    return { item, mid, y: line.top + line.height / 2 };
  };

  test('finds a safe external link under the pointer and nothing beside it', async () => {
    const session = await open(linked('https://example.com/a'));
    const { item, mid, y } = textPoint(session);
    const frame = item.text!.box;
    // The shape has no rotation: the text frame starts at its box inside the shape.
    const x = IN + frame.x + mid;
    expect(slideshowLinkAt(session, 0, x, IN + frame.y + y)).toBe('https://example.com/a');
    expect(slideshowLinkAt(session, 0, 0, 0)).toBeNull();
    expect(slideshowLinkAt(session, 5, x, IN + frame.y + y)).toBeNull();
  });

  test('withholds a script link', async () => {
    const session = await open(linked('javascript:alert(1)'));
    const { item, mid, y } = textPoint(session);
    const frame = item.text!.box;
    expect(slideshowLinkAt(session, 0, IN + frame.x + mid, IN + frame.y + y)).toBeNull();
  });
});
