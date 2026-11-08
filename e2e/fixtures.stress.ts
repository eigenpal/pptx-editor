import { readdirSync, readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { expect, test } from '@playwright/test';
import type { PptxEditorInstance } from '../packages/core/src/editor/index.ts';

const fixtures = fileURLToPath(new URL('../fixtures/', import.meta.url));
for (const file of readdirSync(fixtures)
  .filter((file) => file.endsWith('.pptx'))
  .sort()) {
  test(`${file}: every slide paints and the saved presentation reopens`, async ({ page }, info) => {
    const errors: string[] = [];
    page.on('pageerror', (error) => errors.push(error.message));
    await page.goto('/?deck=basic');
    await page.waitForSelector('.pptx-surface__slide');
    const bytes = Array.from(readFileSync(`${fixtures}/${file}`));
    const start = Date.now();
    const result = await page.evaluate(async (bytes) => {
      const ed = (window as unknown as { __pptxEditor: PptxEditorInstance }).__pptxEditor;
      const opened = await ed.load(new Uint8Array(bytes));
      if (!opened.ok) throw new Error(JSON.stringify(opened));
      const count = ed.snapshot().slides.length;
      for (let i = 0; i < count; i++) {
        ed.goToSlide(i);
        await new Promise<void>((resolve) => requestAnimationFrame(() => resolve()));
        if (!document.querySelector('.pptx-surface__slide')) throw new Error(`Missing slide ${i}`);
      }
      const saved = ed.save();
      if (!saved) throw new Error('Save refused');
      const reopened = await ed.load(saved);
      if (!reopened.ok) throw new Error(JSON.stringify(reopened));
      return { count, reopenedCount: ed.snapshot().slides.length, status: ed.snapshot().status };
    }, bytes);
    expect(result.count).toBeGreaterThan(0);
    expect(result.reopenedCount).toBe(result.count);
    expect(result.status).toBe('ready');
    expect(errors).toEqual([]);
    info.annotations.push({
      type: 'coverage',
      description: `${result.count} slides, ${Date.now() - start} ms`,
    });
  });
}
