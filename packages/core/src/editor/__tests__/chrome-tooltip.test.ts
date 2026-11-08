// Chrome tooltips: the placement rule, the attributes a control carries, and the document
// listener's timing (hover delay, keyboard focus, dismissal), after DOCX Editor's `Tooltip`.

import { afterEach, describe, expect, test } from 'bun:test';
import {
  TOOLTIP_DELAY_MS,
  TOOLTIP_SKIP_DELAY_MS,
  listenChromeTooltips,
  placeTooltip,
  tooltipAttributes,
  type ChromeTooltipState,
} from '../index.ts';
import { GlobalWindow } from 'happy-dom';

// A DOM of this file's own: the core run has no global DOM (`bun run test`).
const win = new GlobalWindow();
const document = win.document as unknown as Document;
const MouseEvent = win.MouseEvent as unknown as typeof globalThis.MouseEvent;
const KeyboardEvent = win.KeyboardEvent as unknown as typeof globalThis.KeyboardEvent;

const wait = (ms: number) => new Promise((resolve) => setTimeout(resolve, ms));
const viewport = { width: 800, height: 600 };

describe('placeTooltip', () => {
  test('centers the bubble 8 pixels under the control', () => {
    const at = placeTooltip(
      { left: 100, right: 140, top: 10, bottom: 40 },
      { width: 60, height: 24 },
      viewport
    );
    expect(at).toEqual({ left: 90, top: 48, side: 'bottom' });
  });

  test('goes over the control when there is no room under it', () => {
    const at = placeTooltip(
      { left: 100, right: 140, top: 560, bottom: 590 },
      { width: 60, height: 24 },
      viewport
    );
    expect(at).toEqual({ left: 90, top: 528, side: 'top' });
  });

  test('stays 8 pixels inside the window at both edges', () => {
    const size = { width: 200, height: 24 };
    expect(placeTooltip({ left: 0, right: 20, top: 0, bottom: 20 }, size, viewport).left).toBe(8);
    expect(placeTooltip({ left: 780, right: 800, top: 0, bottom: 20 }, size, viewport).left).toBe(
      592
    );
  });
});

describe('tooltipAttributes', () => {
  test('names the label, and the reason only when the control is disabled', () => {
    expect(tooltipAttributes('Undo (Ctrl+Z)', true, 'Nothing to undo')).toEqual({
      'data-pptx-tooltip': 'Undo (Ctrl+Z)',
      'data-pptx-tooltip-reason': 'Nothing to undo',
    });
    expect(tooltipAttributes('Undo (Ctrl+Z)', false, 'Nothing to undo')).toEqual({
      'data-pptx-tooltip': 'Undo (Ctrl+Z)',
    });
    expect(tooltipAttributes('Undo (Ctrl+Z)', true, null)).toEqual({
      'data-pptx-tooltip': 'Undo (Ctrl+Z)',
    });
  });
});

describe('listenChromeTooltips', () => {
  const stops: (() => void)[] = [];
  afterEach(() => {
    for (const stop of stops.splice(0)) stop();
    document.body.replaceChildren();
  });

  function setup() {
    const scope = document.createElement('div');
    scope.className = 'pptx-editor';
    const button = (label: string, reason?: string): HTMLButtonElement => {
      const b = document.createElement('button');
      for (const [k, v] of Object.entries(tooltipAttributes(label, !!reason, reason)))
        b.setAttribute(k, v);
      b.append(document.createElement('span'));
      scope.append(b);
      return b;
    };
    const bold = button('Bold (Ctrl+B)');
    const undo = button('Undo (Ctrl+Z)', 'Nothing to undo');
    const plain = document.createElement('div');
    scope.append(plain);
    document.body.append(scope);
    const seen: (ChromeTooltipState | null)[] = [];
    stops.push(listenChromeTooltips(document, (s) => seen.push(s)));
    const last = () => seen[seen.length - 1] ?? null;
    return { scope, bold, undo, plain, seen, last };
  }

  const over = (el: Element) =>
    el.dispatchEvent(new MouseEvent('pointerover', { bubbles: true, composed: true }));

  test('a hover opens the tooltip after the delay, and leaving closes it', async () => {
    const { bold, plain, last, scope } = setup();
    // The pointer over a child of the control counts as the control.
    over(bold.firstElementChild!);
    await wait(TOOLTIP_DELAY_MS / 2);
    expect(last()).toBe(null);
    await wait(TOOLTIP_DELAY_MS / 2 + 30);
    expect(last()?.label).toBe('Bold (Ctrl+B)');
    expect(last()?.instant).toBe(false);
    expect(last()?.reason).toBe(null);
    expect(last()?.target).toBe(bold);
    expect(last()?.scope).toBe(scope);
    over(plain);
    expect(last()).toBe(null);
  });

  test('a disabled control shows its reason; the next control opens at once', async () => {
    const { bold, undo, plain, last } = setup();
    over(undo);
    await wait(TOOLTIP_DELAY_MS + 30);
    expect(last()?.reason).toBe('Nothing to undo');
    // Moving along the bar: no second delay.
    over(plain);
    over(bold);
    expect(last()?.label).toBe('Bold (Ctrl+B)');
    expect(last()?.instant).toBe(true);
    over(plain);
    await wait(TOOLTIP_SKIP_DELAY_MS + 30);
    over(undo);
    expect(last()).toBe(null);
  });

  test('a press closes it, and it opens again only after the pointer leaves', async () => {
    const { bold, plain, last } = setup();
    over(bold);
    await wait(TOOLTIP_DELAY_MS + 30);
    expect(last()?.target).toBe(bold);
    bold.dispatchEvent(new MouseEvent('pointerdown', { bubbles: true }));
    expect(last()).toBe(null);
    over(bold);
    await wait(TOOLTIP_DELAY_MS + 30);
    expect(last()).toBe(null);
    over(plain);
    over(bold);
    await wait(TOOLTIP_DELAY_MS + 30);
    expect(last()?.target).toBe(bold);
  });

  test('keyboard focus opens it; Escape and blur close it; pointer focus does not', async () => {
    const { bold, undo, last } = setup();
    document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Tab', bubbles: true }));
    bold.focus();
    expect(last()?.instant).toBe(true);
    expect(last()?.target).toBe(bold);
    document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Escape', bubbles: true }));
    expect(last()).toBe(null);
    undo.focus();
    await wait(TOOLTIP_DELAY_MS + 30);
    expect(last()?.target).toBe(undo);
    undo.blur();
    expect(last()).toBe(null);
    // A pointer press moves focus: no tooltip for the focus alone.
    bold.dispatchEvent(new MouseEvent('pointerdown', { bubbles: true }));
    bold.focus();
    await wait(TOOLTIP_DELAY_MS + 30);
    expect(last()).toBe(null);
  });

  test('a control with its menu open shows no tooltip', async () => {
    const { bold, last } = setup();
    bold.setAttribute('aria-expanded', 'true');
    over(bold);
    await wait(TOOLTIP_DELAY_MS + 30);
    expect(last()).toBe(null);
  });

  test('keyboard focus takes over a pending hover without waiting or later reopening', async () => {
    const { bold, last } = setup();
    over(bold);
    expect(last()).toBe(null);
    document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Tab', bubbles: true }));
    bold.focus();
    expect(last()?.target).toBe(bold);
    expect(last()?.instant).toBe(true);
    bold.blur();
    expect(last()).toBe(null);
    await wait(TOOLTIP_DELAY_MS + 30);
    expect(last()).toBe(null);
  });

  test('keyboard focus takes over a visible hover and closes it on blur', async () => {
    const { bold, last } = setup();
    over(bold);
    await wait(TOOLTIP_DELAY_MS + 30);
    expect(last()?.instant).toBe(false);
    document.dispatchEvent(new KeyboardEvent('keydown', { key: 'Tab', bubbles: true }));
    bold.focus();
    expect(last()?.target).toBe(bold);
    expect(last()?.instant).toBe(true);
    bold.blur();
    expect(last()).toBe(null);
  });

  test('two roots on a page: only the first one still mounted renders the bubble', async () => {
    const { bold, seen } = setup();
    const second: (ChromeTooltipState | null)[] = [];
    const stopSecond = listenChromeTooltips(document, (s) => second.push(s));
    stops.push(stopSecond);
    over(bold);
    await wait(TOOLTIP_DELAY_MS + 30);
    expect(seen.length).toBe(1);
    expect(second.length).toBe(0);
    // The first root unmounts: the second takes over the tooltip on screen.
    stops.shift()!();
    expect(second[second.length - 1]?.target).toBe(bold);
  });
});
