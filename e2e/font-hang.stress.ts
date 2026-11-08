import { fileURLToPath } from 'node:url';
import { expect, test, type Route } from '@playwright/test';
import type { Editor } from '../packages/core/src/contracts/editor.ts';
import type { FontConfigurationFragment } from '../packages/core/src/layout/font-composition.ts';

type Probe = Window & { __pptxEditor: Editor };

test('a stalled browser font load releases the loading screen and permits another open', async ({
  page,
}) => {
  await page.goto('/?deck=basic');
  await page.waitForSelector('.pptx-surface__slide');
  const result = await page.evaluate(async () => {
    const editor = (window as Probe).__pptxEditor;
    const bytes = new Uint8Array(await (await fetch('/sample.pptx')).arrayBuffer());
    const original = FontFace.prototype.load;
    let calls = 0;
    FontFace.prototype.load = function () {
      calls += 1;
      return new Promise<FontFace>(() => {});
    };
    const start = performance.now();
    try {
      const opened = await editor.load(bytes);
      await editor.setFonts({ sources: [] });
      return { opened, calls, ms: performance.now() - start, status: editor.snapshot().status };
    } finally {
      FontFace.prototype.load = original;
    }
  });
  expect(result.calls).toBeGreaterThan(0);
  expect(result.opened.ok).toBe(true);
  expect(result.status).toBe('ready');
  expect(result.ms).toBeLessThan(6000);
  await expect(page.locator('.pptx-surface__slide')).toBeVisible();
  expect(
    await page.evaluate(async () => {
      const bytes = new Uint8Array(await (await fetch('/sample.pptx')).arrayBuffer());
      return (await (window as Probe).__pptxEditor.load(bytes)).ok;
    })
  ).toBe(true);
});

test('cyclic deferred font answers leave browser timers and the editor responsive', async ({
  page,
}) => {
  const errors: string[] = [];
  page.on('pageerror', (error) => errors.push(error.message));
  await page.goto('/?deck=basic');
  await page.waitForSelector('.pptx-surface__slide');
  const result = await page.evaluate(async () => {
    const editor = (window as Probe).__pptxEditor;
    const answer: { sources: []; deferred?: Promise<FontConfigurationFragment> } = { sources: [] };
    answer.deferred = Promise.resolve(answer);
    let ticks = 0;
    const heartbeat = setInterval(() => (ticks += 1), 10);
    try {
      await editor.setFonts(answer);
      await new Promise((resolve) => setTimeout(resolve, 100));
      await editor.setFonts(undefined);
      return {
        ticks,
        status: editor.snapshot().status,
        resolving: editor.fontMeasurement().resolving,
      };
    } finally {
      clearInterval(heartbeat);
    }
  });
  expect(errors).toEqual([]);
  expect(result.ticks).toBeGreaterThan(0);
  expect(result.status).toBe('ready');
  expect(result.resolving).toBe(false);
  await expect(page.locator('.pptx-surface__slide')).toBeVisible();
});

test('a stalled host font resolver cannot hold open and replacement cancels it', async ({
  page,
}) => {
  const errors: string[] = [];
  page.on('pageerror', (error) => errors.push(error.message));
  await page.goto('/?deck=basic');
  await page.waitForSelector('.pptx-surface__slide');
  const moduleUrl = `/@fs${fileURLToPath(new URL('../packages/core/src/editor/index.ts', import.meta.url))}`;
  const result = await page.evaluate(async (url) => {
    const { defineFontResolver } = (await import(
      url
    )) as typeof import('../packages/core/src/editor/index.ts');
    const editor = (window as Probe).__pptxEditor;
    const bytes = new Uint8Array(await (await fetch('/sample.pptx')).arrayBuffer());
    let calls = 0;
    let aborted = 0;
    const pending = editor.setFonts(
      defineFontResolver(({ signal }) => {
        calls += 1;
        signal?.addEventListener('abort', () => (aborted += 1), { once: true });
        return new Promise<FontConfigurationFragment>(() => {});
      })
    );
    const start = performance.now();
    const opened = await editor.load(bytes);
    const ms = performance.now() - start;
    const status = editor.snapshot().status;
    await pending;
    await editor.setFonts(undefined);
    const recovery = await editor.load(bytes);
    return { opened, status, ms, calls, aborted, recovery };
  }, moduleUrl);
  expect(errors).toEqual([]);
  expect(result.opened.ok).toBe(true);
  expect(result.status).toBe('ready');
  expect(result.ms).toBeLessThan(6000);
  expect(result.calls).toBeGreaterThanOrEqual(2);
  expect(result.aborted).toBe(result.calls);
  expect(result.recovery.ok).toBe(true);
  await expect(page.locator('.pptx-surface__slide')).toBeVisible();
});

for (const resource of ['wasm', 'packaged fonts'] as const) {
  test(`startup stays usable while ${resource} requests never answer`, async ({ page }) => {
    const held: Route[] = [];
    const pattern = resource === 'wasm' ? /\.wasm(?:\?|$)/ : /\.(?:ttf|otf|woff2?)(?:\?|$)/;
    await page.route(pattern, async (route) => {
      // Vite's asset URL modules must load; only the actual binary fetch stalls.
      if (route.request().resourceType() === 'fetch') held.push(route);
      else await route.continue();
    });
    try {
      await page.goto('/?deck=basic', { waitUntil: 'domcontentloaded' });
      await page.waitForFunction(
        () => (window as Probe).__pptxEditor?.snapshot().status === 'ready',
        undefined,
        { timeout: 10_000 }
      );
      expect(held.length).toBeGreaterThan(0);
      await expect(page.locator('.pptx-loading')).toHaveCount(0);
      await expect(page.locator('.pptx-surface__slide')).toBeVisible();
      await page.getByRole('menuitem', { name: 'File', exact: true }).click();
      await expect(page.getByRole('menuitem', { name: 'Open…', exact: true })).toBeVisible();
      await page.keyboard.press('Escape');
      const responsive = await page.evaluate(async () => {
        await new Promise((resolve) => setTimeout(resolve, 20));
        return (window as Probe).__pptxEditor.snapshot().status;
      });
      expect(responsive).toBe('ready');
    } finally {
      await page.unroute(pattern);
      await Promise.allSettled(held.map((route) => route.abort()));
    }
  });
}
