// Browser open bench: how soon a long deck shows its first slide in the demo (React chrome, real
// rail), how soon the whole deck is ready, and how long a keystroke takes while the rest of the
// deck opens in the background. Uses the synthetic decks of `scripts/bench/text-deck.ts`.
//
// - First slide: wall time from the file input to the first frame that shows slide 1's text on
//   the canvas (the page polls once a frame, so a busy main thread delays it).
// - Ready: wall time until every rail item is in place and the editor reports no slide still
//   opening (`snapshot().pendingSlides` is absent or 0), in an open with no input.
// - Keys: 20 keys typed into slide 1's body at 50 per second, once as soon as slide 1 shows
//   (while the rest of the deck opens) and once after the deck is ready. For each key, the time
//   from the key event to the end of the task after the next frame: the input delay, the
//   engine, React, and the paint. A blocked main thread shows as a long first key.
//
// Usage: bun run bench:open:browser            (sizes from BENCH_SIZES, default 800,5000)

import { test, type Page } from '@playwright/test';
import { textDeck } from '../scripts/bench/text-deck.ts';

const SIZES = (process.env.BENCH_SIZES ?? '800,5000').split(',').map(Number);
const KEYS = 20;

interface Keys {
  medianMs: number;
  maxMs: number;
}

interface Row {
  slides: number;
  firstSlideMs: number;
  readyMs: number;
  /** Keys typed as soon as the first slide shows, while the rest of the deck opens. */
  openingKeys: Keys;
  /** Keys typed after the deck is ready: the same work with nothing in the background. */
  readyKeys: Keys;
}

type Probe = { __keys: number[]; __pptxEditor: unknown };
type Item = { element: { name?: string; nodeId: string }; text?: { paragraphLengths: number[] } };

/** Open the demo and arm the key probe: from each key's time stamp to the task after the next frame. */
async function start(page: Page): Promise<void> {
  await page.goto('/');
  await page.waitForSelector('.pptx-surface__slide', { timeout: 30_000 });
  await page.waitForFunction(() => (window as unknown as Probe).__pptxEditor !== undefined);
  await page.evaluate(() => {
    const w = window as unknown as Probe;
    w.__keys = [];
    window.addEventListener(
      'keydown',
      (e) => {
        const at = e.timeStamp;
        requestAnimationFrame(() => setTimeout(() => w.__keys.push(performance.now() - at), 0));
      },
      { capture: true }
    );
  });
}

/** Choose a deck of `n` slides; resolves when slide 1 shows, with the wall time. */
async function openDeck(page: Page, n: number): Promise<{ t0: number; firstSlideMs: number }> {
  const buffer = Buffer.from(textDeck(n));
  const t0 = Date.now();
  await page.locator('input[type=file][accept*=".pptx"]').setInputFiles({
    name: `bench-${n}.pptx`,
    mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    buffer,
  });
  await page.waitForFunction(
    () => document.querySelector('.pptx-surface__slide')?.textContent?.includes('Slide 1 title'),
    undefined,
    { timeout: 300_000, polling: 'raf' }
  );
  return { t0, firstSlideMs: Date.now() - t0 };
}

/** Wait until every rail item is in place and the editor reports no slide still opening. */
async function ready(page: Page, n: number): Promise<void> {
  await page.waitForFunction(
    (count) => {
      const ed = (window as unknown as Probe).__pptxEditor as {
        snapshot(): { pendingSlides?: number };
      };
      return (
        document.querySelectorAll('.pptx-rail__item').length === count &&
        !ed.snapshot().pendingSlides
      );
    },
    n,
    { timeout: 300_000, polling: 'raf' }
  );
}

/** Type `KEYS` characters at the end of slide 1's body; the latency of each key. */
async function typeKeys(page: Page): Promise<Keys> {
  await page.evaluate(() => {
    const w = window as unknown as Probe;
    w.__keys = [];
    const ed = w.__pptxEditor as {
      session: { layout(i: number): { items: Item[] } };
      setSelection(s: unknown): void;
      goToSlide(i: number): void;
      focus(): void;
    };
    ed.goToSlide(0);
    const body = ed.session.layout(0).items.find((x) => x.element.name === 'Body')!;
    const at = { paragraph: 0, offset: body.text?.paragraphLengths[0] ?? 0 };
    ed.setSelection({ kind: 'text', nodeId: body.element.nodeId, anchor: at, focus: at });
    ed.focus();
  });
  await page.keyboard.type('x'.repeat(KEYS), { delay: 20 });
  await page.waitForFunction((keys) => (window as unknown as Probe).__keys.length >= keys, KEYS, {
    timeout: 120_000,
  });
  const keys = await page.evaluate(() => (window as unknown as Probe).__keys.slice());
  const sorted = [...keys].sort((a, b) => a - b);
  return { medianMs: sorted[Math.floor(sorted.length / 2)] ?? 0, maxMs: sorted.at(-1) ?? 0 };
}

async function measure(page: Page, n: number): Promise<Row> {
  // Open with no input: the time to the first slide and to a ready deck.
  await start(page);
  const first = await openDeck(page, n);
  await ready(page, n);
  const readyMs = Date.now() - first.t0;
  await page.waitForTimeout(1_000);
  const readyKeys = await typeKeys(page);
  // Open again and type at once, while the rest of the deck opens.
  await start(page);
  await openDeck(page, n);
  const openingKeys = await typeKeys(page);
  await ready(page, n);
  return { slides: n, firstSlideMs: first.firstSlideMs, readyMs, openingKeys, readyKeys };
}

test('open bench', async ({ page }) => {
  const rows: Row[] = [];
  for (const n of SIZES) rows.push(await measure(page, n));
  const k = (x: Keys) => `${x.medianMs.toFixed(1)} / ${x.maxMs.toFixed(1)}`;
  console.log(
    [
      '| Slides | First slide ms | Ready ms | Keys while opening ms (median / max) | Keys when ready ms (median / max) |',
      '| --: | --: | --: | --: | --: |',
      ...rows.map(
        (r) =>
          `| ${r.slides} | ${r.firstSlideMs} | ${r.readyMs} | ${k(r.openingKeys)} | ${k(r.readyKeys)} |`
      ),
    ].join('\n')
  );
});
