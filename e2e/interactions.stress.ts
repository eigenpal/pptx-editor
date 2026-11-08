import { expect, test } from '@playwright/test';
import type { PptxEditorInstance } from '../packages/core/src/editor/index.ts';
import { textDeck } from '../scripts/bench/text-deck.ts';

declare global {
  interface Window {
    __pptxEditor: PptxEditorInstance;
  }
}

test('rapid navigation, zoom, resize, find, and history keep both adapters responsive', async ({
  page,
}) => {
  const errors: string[] = [];
  page.on('pageerror', (error) => errors.push(error.message));
  await page.goto('/?deck=basic');
  await page.waitForSelector('.pptx-surface__slide');
  const bytes = Array.from(textDeck(80));
  await page.evaluate(async (bytes) => {
    const result = await window.__pptxEditor.load(new Uint8Array(bytes));
    if (!result.ok) throw new Error(JSON.stringify(result));
  }, bytes);
  await expect(page.locator('.pptx-rail__item')).toHaveCount(80);
  for (let round = 0; round < 8; round++) {
    await page.evaluate((round) => {
      const ed = window.__pptxEditor;
      for (let i = 0; i < 20; i++) {
        ed.goToSlide((round * 23 + i * 13) % 80);
        ed.setZoom(i % 3 === 0 ? 'fit' : i % 2 === 0 ? 4 : 0.1);
        ed.find.next(i % 2 === 0 ? 'text' : '');
      }
      ed.setZoom('fit');
      const before = ed.snapshot().slides.length;
      for (let i = 0; i < 4; i++) {
        if (!ed.exec({ type: 'duplicateSlide', index: 0 }).ok) throw new Error('duplicate refused');
        if (!ed.exec({ type: 'undo' }).ok) throw new Error('undo refused');
        if (!ed.exec({ type: 'redo' }).ok) throw new Error('redo refused');
        if (!ed.exec({ type: 'undo' }).ok) throw new Error('undo refused');
      }
      if (ed.snapshot().slides.length !== before) throw new Error('slide count changed');
    }, round);
    await page.setViewportSize({ width: round % 2 ? 1440 : 720, height: round % 2 ? 900 : 600 });
    // A frame proves that browser paint can run after each burst.
    await page.evaluate(
      () => new Promise<void>((resolve) => requestAnimationFrame(() => resolve()))
    );
  }
  await expect(page.locator('.pptx-surface__slide')).toBeVisible();
  expect(errors).toEqual([]);
  expect(await page.evaluate(() => window.__pptxEditor.snapshot().status)).toBe('ready');
});

test('repeated malformed opens recover and concurrent opens settle', async ({ page }) => {
  const errors: string[] = [];
  page.on('pageerror', (error) => errors.push(error.message));
  await page.goto('/?deck=basic');
  await page.waitForSelector('.pptx-surface__slide');
  const bytes = Array.from(textDeck(10));
  const result = await page.evaluate(async (bytes) => {
    const ed = window.__pptxEditor;
    for (let i = 0; i < 8; i++) {
      const bad = await ed.load(new Uint8Array([80, 75, 3, 4, i]));
      if (bad.ok) throw new Error('invalid archive accepted');
      const good = await ed.load(new Uint8Array(bytes));
      if (!good.ok) throw new Error('recovery failed');
    }
    const pending = Array.from({ length: 12 }, () => ed.load(new Uint8Array(bytes)));
    await Promise.all(pending);
    return {
      status: ed.snapshot().status,
      slides: ed.snapshot().slides.length,
      saved: !!ed.save(),
    };
  }, bytes);
  expect(result).toEqual({ status: 'ready', slides: 10, saved: true });
  await expect(page.locator('.pptx-rail__item')).toHaveCount(10);
  await expect(page.locator('.pptx-surface__slide')).toBeVisible();
  expect(errors).toEqual([]);
});

test('keyboard, composition, malformed HTML paste, and canceled drags keep input usable', async ({
  page,
}) => {
  const errors: string[] = [];
  page.on('pageerror', (error) => errors.push(error.message));
  await page.goto('/?deck=basic');
  await page.waitForSelector('.pptx-surface__slide');
  await page.evaluate(
    async (bytes) => {
      const ed = window.__pptxEditor;
      await ed.load(new Uint8Array(bytes));
      const item = ed.session!.layout(0)!.items.find((item) => item.element.name === 'Body')!;
      const at = { paragraph: 0, offset: 0 };
      ed.setSelection({ kind: 'text', nodeId: item.element.nodeId, anchor: at, focus: at });
      ed.focus();
    },
    Array.from(textDeck(3))
  );
  await page.locator('.pptx-surface__input').focus();
  await page.keyboard.type('Adversarial keyboard sequence '.repeat(10));
  expect(
    await page.evaluate(
      () => window.__pptxEditor.find.matches('Adversarial keyboard sequence').length
    )
  ).toBe(10);
  for (let i = 0; i < 12; i++) {
    await page.keyboard.press('ArrowLeft');
    await page.keyboard.press('Shift+ArrowLeft');
    await page.keyboard.press('Backspace');
    await page.keyboard.press('ControlOrMeta+z');
  }
  for (const html of [
    '<p>Clipboard marker</p>' + '<'.repeat(50_000),
    '<p>Clipboard marker</p>' + '<span title="'.repeat(5_000),
    '<div>'.repeat(5_000) + '<p>Clipboard marker</p>' + '</div>'.repeat(5_000),
    '<p>Clipboard marker</p><!--' + '>'.repeat(50_000),
  ]) {
    await page.evaluate((html) => {
      const target = document.querySelector('.pptx-surface__input')!;
      const data = new DataTransfer();
      data.setData('text/html', html);
      data.setData('text/plain', 'Clipboard marker');
      const event = new ClipboardEvent('paste', {
        clipboardData: data,
        bubbles: true,
        cancelable: true,
      });
      // Firefox clears data passed to an untrusted ClipboardEvent constructor.
      // Supply the same clipboard payload to the event handler in all engines.
      Object.defineProperty(event, 'clipboardData', { value: data });
      target.dispatchEvent(event);
    }, html);
    await page.evaluate(
      () => new Promise<void>((resolve) => requestAnimationFrame(() => resolve()))
    );
    expect(
      await page.evaluate(() => window.__pptxEditor.find.matches('Clipboard marker').length)
    ).toBeGreaterThan(0);
    if (html.includes('<'.repeat(50_000))) {
      expect(
        await page.evaluate(() => window.__pptxEditor.find.matches('<'.repeat(1000)).length)
      ).toBe(50);
    }
    await page.keyboard.press('ControlOrMeta+z');
  }
  await page.evaluate(() => {
    const input = document.querySelector('.pptx-surface__input')!;
    input.dispatchEvent(new CompositionEvent('compositionstart', { bubbles: true }));
    input.dispatchEvent(
      new CompositionEvent('compositionupdate', { bubbles: true, data: '日本語' })
    );
    input.dispatchEvent(new CompositionEvent('compositionend', { bubbles: true, data: '日本語' }));
  });
  expect(await page.evaluate(() => window.__pptxEditor.find.matches('日本語').length)).toBe(1);
  await page.keyboard.press('Escape');
  const stage = await page.locator('.pptx-surface__stage').boundingBox();
  expect(stage).not.toBeNull();
  for (let i = 0; i < 10; i++) {
    await page.mouse.move(stage!.x + 10, stage!.y + 10);
    await page.mouse.down();
    await page.mouse.move(stage!.x + 100, stage!.y + 100, { steps: 4 });
    await page.keyboard.press('Escape');
    await page.mouse.up();
  }
  await page.keyboard.press('ControlOrMeta+f');
  await expect(page.getByRole('search', { name: 'Find and replace' })).toBeVisible();
  await page.keyboard.press('Escape');
  expect(await page.evaluate(() => !!window.__pptxEditor.save())).toBe(true);
  expect(errors).toEqual([]);
});
