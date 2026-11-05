// End-to-end verification of the demo in Playwright Chromium.
//
//   bun scripts/verify-demo.ts            # sample and showcase decks + fixtures through the real UI
//   bun scripts/verify-demo.ts --corpus   # also the private decks (local only)
//   bun scripts/verify-demo.ts --parity   # family-parity screenshots vs the live DOCX demo
//
// For every deck: open it with the demo's own Open button, visit every slide in the rail,
// export with the Export menu, reopen the downloaded bytes, and fail on any console error.

import { chromium, type Page } from '@playwright/test';
import { existsSync, mkdirSync, readdirSync, readFileSync } from 'node:fs';
import { join, resolve } from 'node:path';
import { createServer } from 'vite';
import { readPptx, presentationInfo } from '../packages/core/src/store/index.ts';

const args = process.argv.slice(2);

async function openDeck(page: Page, file: string): Promise<number> {
  const expected = (() => {
    const r = readPptx(new Uint8Array(readFileSync(file)));
    return r.ok ? presentationInfo(r.pkg).slides.length : -1;
  })();
  await page.locator('input[type=file][accept*=".pptx"]').setInputFiles(file);
  await page.waitForFunction(
    (n) => document.querySelectorAll('.pptx-rail__item').length === n,
    expected,
    { timeout: 60_000 }
  );
  return expected;
}

async function visitSlides(page: Page, count: number): Promise<void> {
  for (let i = 0; i < count; i += 1) {
    await page.locator('.pptx-rail__item').nth(i).click();
    await page.waitForFunction(
      (n) =>
        document.querySelector('.pptx-rail__item[aria-current="true"] .pptx-rail__number')
          ?.textContent === String(n),
      i + 1
    );
    await page.waitForSelector('.pptx-surface__slide');
  }
}

async function exportDeck(page: Page): Promise<Uint8Array> {
  const [download] = await Promise.all([
    page.waitForEvent('download'),
    (async () => {
      await page.getByRole('button', { name: 'Export' }).click();
      await page.getByRole('menuitem', { name: /\.pptx/ }).click();
    })(),
  ]);
  const path = await download.path();
  return new Uint8Array(readFileSync(path!));
}

async function parity(page: Page, url: string): Promise<void> {
  const out = 'screenshots/family-parity';
  mkdirSync(out, { recursive: true });
  for (const theme of ['light', 'dark'] as const) {
    await page.emulateMedia({ colorScheme: theme });
    await page.goto(`${url}?theme=${theme}`);
    await page.waitForSelector('.pptx-rail__item');
    await page.waitForTimeout(1500);
    await page.screenshot({ path: join(out, `pptx-editor-${theme}.png`) });
    try {
      await page.goto('https://docx-editor.dev/editor');
      await page.waitForTimeout(6000);
      if (theme === 'dark') {
        // The DOCX demo keeps its theme in the Settings tab of its side panel.
        const settings = page.getByRole('tab', { name: /Settings/i });
        if (await settings.count()) {
          await settings.first().click();
          const dark = page.getByRole('radio', { name: /Dark/i });
          if (await dark.count()) await dark.first().click();
          await page.waitForTimeout(800);
        }
      }
      await page.screenshot({ path: join(out, `docx-editor-${theme}.png`) });
    } catch (e) {
      console.log(`DOCX demo screenshot (${theme}) failed: ${(e as Error).message.split('\n')[0]}`);
    }
  }
  console.log(`Parity screenshots in ${out}/`);
}

async function main(): Promise<void> {
  const server = await createServer({
    configFile: resolve('examples/vite/vite.config.ts'),
    root: resolve('examples/vite'),
    server: { port: 5298, strictPort: false },
    logLevel: 'error',
  });
  await server.listen();
  const url = server.resolvedUrls?.local[0] ?? 'http://localhost:5298/';
  const browser = await chromium.launch();
  const page = await browser.newPage({
    viewport: { width: 1440, height: 900 },
    acceptDownloads: true,
  });
  const errors: string[] = [];
  page.on('console', (m) => {
    if (m.type() === 'error') errors.push(m.text());
  });
  page.on('pageerror', (e) => errors.push(e.message));
  let failed = false;
  try {
    if (args.includes('--parity')) {
      await parity(page, url);
      return;
    }
    await page.goto(url);
    await page.waitForSelector('.pptx-rail__item', { timeout: 60_000 });
    const decks: { name: string; file: string }[] = [
      { name: 'sample', file: 'examples/vite/public/sample.pptx' },
      { name: 'showcase', file: 'examples/vite/public/showcase.pptx' },
      ...readdirSync('fixtures')
        .filter((f) => f.endsWith('.pptx'))
        .sort()
        .map((f) => ({ name: f.replace('.pptx', ''), file: join('fixtures', f) })),
    ];
    if (args.includes('--corpus') && existsSync('demo-presentations'))
      readdirSync('demo-presentations')
        .filter((f) => f.endsWith('.pptx'))
        .sort()
        .forEach((f, i) =>
          decks.push({ name: `deck${i + 1}`, file: join('demo-presentations', f) })
        );
    for (const deck of decks) {
      const before = errors.length;
      const t0 = performance.now();
      const count = await openDeck(page, deck.file);
      const opened = performance.now() - t0;
      await visitSlides(page, count);
      const bytes = await exportDeck(page);
      const reopened = readPptx(bytes);
      const ok = reopened.ok && presentationInfo(reopened.pkg).slides.length === count;
      const newErrors = errors.slice(before);
      if (!ok || newErrors.length) failed = true;
      console.log(
        `${deck.name}: ${count} slides, opened in ${Math.round(opened)} ms, export ${ok ? 'reopens' : 'FAILED'}, console errors ${newErrors.length}`
      );
      for (const e of newErrors.slice(0, 3)) console.log(`  ${e}`);
    }
  } finally {
    await browser.close();
    await server.close();
  }
  process.exit(failed ? 1 : 0);
}

await main();
