import { expect, test } from '@playwright/test';
import type { Editor } from '../packages/core/src/contracts/editor.ts';
import { HANG_DECKS } from '../scripts/stress/hang-decks.ts';

type Probe = Window & { __pptxEditor: Editor };
const only = process.env.HANG_ONLY?.split(',');
for (const deck of HANG_DECKS.filter((d) => !d.stalledFonts && (!only || only.includes(d.id)))) {
  test(deck.id, async ({ page }, info) => {
    const errors: string[] = [];
    page.on('pageerror', (error) => errors.push(error.message));
    const bytes = Buffer.from(deck.bytes());
    await page.route('**/adversarial.pptx', (route) => route.fulfill({ body: bytes }));
    await page.goto('/?deck=basic');
    await page.waitForSelector('.pptx-surface__slide');
    const result = await page.evaluate(async () => {
      const editor = (window as Probe).__pptxEditor;
      const bytes = new Uint8Array(await (await fetch('/adversarial.pptx')).arrayBuffer());
      let maxGap = 0;
      let last = performance.now();
      const timer = setInterval(() => {
        const now = performance.now();
        maxGap = Math.max(maxGap, now - last);
        last = now;
      }, 25);
      const start = performance.now();
      try {
        const opened = await editor.load(bytes);
        while (editor.snapshot().pendingSlides) await new Promise((r) => setTimeout(r, 25));
        await new Promise((r) => setTimeout(r, 100));
        return { opened, ms: performance.now() - start, maxGap, status: editor.snapshot().status };
      } finally {
        clearInterval(timer);
      }
    });
    console.log(JSON.stringify({ id: deck.id, ...result, errors }));
    await info.attach('open-result', {
      body: JSON.stringify({ ...result, errors }),
      contentType: 'application/json',
    });
    expect(errors).toEqual([]);
    expect(result.ms).toBeLessThan(30_000);
    expect(result.maxGap).toBeLessThan(5000);
    if (result.opened.ok) expect(result.status).toBe('ready');
    // A refused input must still leave the editor usable for the next file.
    const recovery = await page.evaluate(async () => {
      const editor = (window as Probe).__pptxEditor;
      const bytes = new Uint8Array(await (await fetch('/sample.pptx')).arrayBuffer());
      return editor.load(bytes);
    });
    expect(recovery.ok).toBe(true);
    await expect(page.locator('.pptx-surface__slide')).toBeVisible();
  });
}
