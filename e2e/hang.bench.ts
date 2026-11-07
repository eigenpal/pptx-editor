// Browser hang bench: opens each adversarial deck of `scripts/stress/hang-decks.ts` (with the
// decks of `scripts/stress/ui-scale-decks.ts`) in the demo (React chrome, real rail, real canvas) and
// measures what a person feels: how long the open takes, the longest main-thread task, the
// total blocking time, and the heap and DOM size after the open. The happy-dom check
// (`bun run stress:hang`) finds algorithmic hangs; this bench finds the browser's own costs:
// style, layout, paint, canvas, fonts, and memory.
//
// - HANG: the deck is not ready, and not refused, within the budget, or the page stops
//   answering.
// - FREEZE: one main-thread task is longer than the freeze mark: the page does not answer
//   input for that long.
// - SLOW: the open is longer than the slow mark.
//
// Each deck opens in a fresh browser context, so one deck's heap never counts in the next.
//
// Usage:
//   bun run bench:hang:browser
//   HANG_ONLY=id,id HANG_BUDGET_MS=20000 HANG_FREEZE_MS=1000 bun run bench:hang:browser
//   HANG_STRICT=1 bun run bench:hang:browser     (fails on any HANG or FREEZE)

import { mkdirSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { test, type Browser, type Page } from '@playwright/test';
import type { HangDeck } from '../scripts/stress/deck-parts.ts';
import { HANG_DECKS } from '../scripts/stress/hang-decks.ts';

const BUDGET_MS = Number(process.env.HANG_BUDGET_MS ?? 20_000);
const FREEZE_MS = Number(process.env.HANG_FREEZE_MS ?? 1_000);
const SLOW_MS = Number(process.env.HANG_SLOW_MS ?? 3_000);
const ONLY = process.env.HANG_ONLY?.split(',').filter(Boolean);

interface Row {
  id: string;
  verdict: 'ok' | 'SLOW' | 'FREEZE' | 'HANG';
  openMs: number;
  longestTaskMs: number;
  blockingMs: number;
  heapMB: number;
  domNodes: number;
  status: string;
  refused?: string;
}

type Probe = {
  __pptxEditor?: { snapshot(): { status: string; pendingSlides?: number; error?: unknown } };
  __tasks: number[];
};

/** Record every main-thread task over 50 ms (`longtask`), from the page's first script. */
const OBSERVE_TASKS = () => {
  const w = window as unknown as Probe;
  w.__tasks = [];
  try {
    new PerformanceObserver((list) => {
      for (const e of list.getEntries()) w.__tasks.push(e.duration);
    }).observe({ type: 'longtask', buffered: true });
  } catch {
    // No long task timing in this browser: the bench reports 0.
  }
};

/** Resolves with `work`'s value, or with `fallback` when the page does not answer in `ms`. */
async function within<T>(work: Promise<T>, ms: number, fallback: T): Promise<T> {
  let timer: ReturnType<typeof setTimeout> | undefined;
  const late = new Promise<T>((resolve) => (timer = setTimeout(() => resolve(fallback), ms)));
  try {
    return await Promise.race([work, late]);
  } finally {
    clearTimeout(timer);
  }
}

async function measure(browser: Browser, deck: HangDeck): Promise<Row> {
  const context = await browser.newContext({ viewport: { width: 1440, height: 900 } });
  const page: Page = await context.newPage();
  await page.addInitScript(OBSERVE_TASKS);
  const cdp = await context.newCDPSession(page);
  await cdp.send('Performance.enable');
  const row: Row = {
    id: deck.id,
    verdict: 'ok',
    openMs: 0,
    longestTaskMs: 0,
    blockingMs: 0,
    heapMB: 0,
    domNodes: 0,
    status: '',
  };
  try {
    await page.goto('/?deck=basic');
    await page.waitForFunction(
      () => (window as unknown as Probe).__pptxEditor?.snapshot().status === 'ready',
      undefined,
      { timeout: 30_000 }
    );
    await page.evaluate(() => ((window as unknown as Probe).__tasks = []));
    const t0 = Date.now();
    await page.locator('input[type=file][accept*=".pptx"]').setInputFiles({
      name: `${deck.id}.pptx`,
      mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
      buffer: Buffer.from(deck.bytes()),
    });
    // Ready: the deck opened with no slide still opening, or the editor refused it.
    const settled = await within(
      page
        .waitForFunction(
          () => {
            const s = (window as unknown as Probe).__pptxEditor?.snapshot();
            return !!s && (s.status === 'error' || (s.status === 'ready' && !s.pendingSlides));
          },
          undefined,
          { timeout: BUDGET_MS, polling: 100 }
        )
        .then(() => true)
        .catch(() => false),
      BUDGET_MS + 5_000,
      false
    );
    row.openMs = Date.now() - t0;
    if (!settled) row.verdict = 'HANG';
    // Let the idle work (thumbnails, font loads) run, then read the page.
    const answered = await within(
      page.waitForTimeout(1_500).then(() =>
        page.evaluate(() => {
          const w = window as unknown as Probe;
          const s = w.__pptxEditor?.snapshot();
          return {
            tasks: w.__tasks,
            status: s?.status ?? '',
            refused: s?.error ? JSON.stringify(s.error) : undefined,
          };
        })
      ),
      10_000,
      null
    );
    if (!answered) {
      row.verdict = 'HANG';
      return row;
    }
    row.status = answered.status;
    if (answered.refused) row.refused = answered.refused;
    row.longestTaskMs = Math.round(Math.max(0, ...answered.tasks));
    row.blockingMs = Math.round(answered.tasks.reduce((sum, d) => sum + Math.max(0, d - 50), 0));
    const { metrics } = (await within(cdp.send('Performance.getMetrics'), 5_000, {
      metrics: [],
    })) as { metrics: { name: string; value: number }[] };
    const metric = (name: string) => metrics.find((m) => m.name === name)?.value ?? 0;
    row.heapMB = Math.round(metric('JSHeapUsedSize') / 1e6);
    row.domNodes = metric('Nodes');
    if (row.verdict === 'ok' && row.longestTaskMs > FREEZE_MS) row.verdict = 'FREEZE';
    const slow = Math.max(SLOW_MS, deck.slowMs ?? 0);
    if (row.verdict === 'ok' && row.openMs > slow) row.verdict = 'SLOW';
    return row;
  } finally {
    await within(context.close(), 10_000, undefined);
  }
}

test('hang bench', async ({ browser, browserName }) => {
  test.skip(browserName !== 'chromium', 'The bench reads Chrome DevTools protocol metrics.');
  const decks = HANG_DECKS.filter((d) => !d.stalledFonts && (!ONLY || ONLY.includes(d.id)));
  test.setTimeout(decks.length * (BUDGET_MS + 60_000));
  const rows: Row[] = [];
  for (const deck of decks) {
    const row = await measure(browser, deck);
    rows.push(row);
    console.log(
      `${row.verdict.padEnd(6)} ${String(row.openMs).padStart(6)} ms open  ` +
        `${String(row.longestTaskMs).padStart(6)} ms longest  ` +
        `${String(row.blockingMs).padStart(6)} ms blocking  ${String(row.heapMB).padStart(5)} MB  ` +
        `${String(row.domNodes).padStart(7)} nodes  ${row.id}${row.refused ? ` (refused ${row.refused})` : ''}`
    );
  }
  const dir = join(import.meta.dirname, '..', 'test-results');
  mkdirSync(dir, { recursive: true });
  writeFileSync(join(dir, 'hang-bench.json'), `${JSON.stringify(rows, null, 2)}\n`);
  const bad = rows.filter((r) => r.verdict === 'HANG' || r.verdict === 'FREEZE');
  console.log(
    `\n${rows.length - bad.length}/${rows.length} without a hang or freeze; ` +
      `budget ${BUDGET_MS} ms, freeze mark ${FREEZE_MS} ms, slow mark ${SLOW_MS} ms`
  );
  if (process.env.HANG_STRICT) {
    if (bad.length) throw new Error(`Hangs or freezes: ${bad.map((r) => r.id).join(', ')}`);
  }
});
