// Browser hang check: opens each adversarial deck of `scripts/stress/hang-decks.ts` in the demo
// (React chrome, real rail, real canvas), in Chromium, Firefox, or WebKit (`HANG_BROWSER`), and
// measures what a person feels. The happy-dom check (`bun run stress:hang`) finds algorithmic
// hangs; this check finds the browser's own costs: style, layout, paint, picture decoding,
// fonts, and memory.
//
// Each deck is one test in a fresh browser context, so one deck's heap, crash, or stuck page
// never counts in the next. The deck goes in through the demo's file input, as a person opens
// it, from a file on disk: a buffer over 50 MB cannot pass to the browser, and request
// interception is not reliable with large bodies in every engine.
//
// Verdicts, worst first. Each one but `ok` and `SLOW` fails the test:
// - CRASH: the page or the browser closed.
// - ERROR: the page threw an uncaught error, or the editor did not open another deck after it.
// - HANG: the deck is not ready, and not refused, within the budget, or the page stops
//   answering.
// - BLOCK: the main thread did not run a timer for longer than the block mark.
// - FETCH: the page requested a URL outside the demo's own origin. A deck never makes the
//   editor fetch: external relationships and references inside pictures stay unfetched.
// - FREEZE: one main-thread task, timer gap, or gap between frames is longer than the freeze
//   mark. It fails the test only with `HANG_STRICT=1`.
// - SLOW: the open is longer than the slow mark (or the deck's own `slowMs`). It never fails.
//
// In Chromium the check also reads the DevTools protocol: the JS heap and the DOM node count
// after the open. Rows go to `test-results/hang-<browser>/rows.jsonl`.
//
// Usage:
//   bun run test:hang:browser
//   HANG_BROWSER=firefox HANG_ONLY=id,id bun run test:hang:browser
//   HANG_BUDGET_MS=20000 HANG_BLOCK_MS=5000 HANG_FREEZE_MS=1000 HANG_IDLE_MS=1500 HANG_STRICT=1 \
//     bun run bench:hang:browser

import { appendFileSync, mkdirSync, rmSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { expect, test, type Page } from '@playwright/test';
import type { HangDeck } from '../scripts/stress/deck-parts.ts';
import { HANG_DECKS } from '../scripts/stress/hang-decks.ts';

const BUDGET_MS = Number(process.env.HANG_BUDGET_MS ?? 20_000);
const BLOCK_MS = Number(process.env.HANG_BLOCK_MS ?? 5_000);
const FREEZE_MS = Number(process.env.HANG_FREEZE_MS ?? 1_000);
const SLOW_MS = Number(process.env.HANG_SLOW_MS ?? 3_000);
/** How long the idle work after the open (thumbnails, pictures, fonts) runs before the read. */
const IDLE_MS = Number(process.env.HANG_IDLE_MS ?? 1_500);
const STRICT = !!process.env.HANG_STRICT;
const ONLY = process.env.HANG_ONLY?.split(',').filter(Boolean);
const DECK_DIR = join(import.meta.dirname, '..', 'test-results', 'hang-decks');

type Verdict = 'ok' | 'SLOW' | 'FREEZE' | 'FETCH' | 'BLOCK' | 'HANG' | 'ERROR' | 'CRASH';

interface Row {
  id: string;
  browser: string;
  verdict: Verdict;
  openMs: number;
  /** The longest time the page's timer did not run: the main thread was busy. */
  longestGapMs: number;
  /** The longest main-thread task (`longtask`, Chromium only). */
  longestTaskMs: number;
  /** The longest time between two animation frames. */
  longestFrameMs: number;
  /** The sum over long tasks of their time over 50 ms (Chromium only). */
  blockingMs: number;
  /** The JS heap after the open, in MB (Chromium only). */
  heapMB?: number;
  /** The DOM nodes after the open (Chromium only). */
  domNodes?: number;
  status: string;
  refused?: string;
  /** Requests outside the demo's origin, after the deck opened. */
  fetched?: string[];
  errors?: string[];
  /** Why the deck has its verdict, when it is not `ok`. */
  why?: string;
}

type Probe = {
  __pptxEditor?: {
    snapshot(): { status: string; pendingSlides?: number; error?: unknown };
    subscribe(listener: () => void): () => void;
    load(bytes: Uint8Array): Promise<{ ok: boolean }>;
  };
  __tasks: number[];
  /** Gaps between animation frames over 50 ms. */
  __frames: number[];
  /** Gaps over 50 ms between runs of a 25 ms timer. */
  __gaps: number[];
  /** Whether the editor left the open deck: the new open started. */
  __left?: boolean;
};

/**
 * From the page's first script, record every main-thread task over 50 ms (`longtask`, where
 * the browser has it), every gap between animation frames over 50 ms, and every gap over 50 ms
 * between runs of a 25 ms timer. The timer gap works in every browser: it is how long the page
 * did not answer input.
 */
const OBSERVE = () => {
  const w = window as unknown as Probe;
  w.__tasks = [];
  w.__frames = [];
  w.__gaps = [];
  let frame = performance.now();
  const onFrame = (t: number) => {
    if (t - frame > 50) w.__frames.push(t - frame);
    frame = t;
    requestAnimationFrame(onFrame);
  };
  requestAnimationFrame(onFrame);
  let tick = performance.now();
  setInterval(() => {
    const now = performance.now();
    if (now - tick > 50) w.__gaps.push(now - tick);
    tick = now;
  }, 25);
  try {
    new PerformanceObserver((list) => {
      for (const e of list.getEntries()) w.__tasks.push(e.duration);
    }).observe({ type: 'longtask', buffered: true });
  } catch {
    // No long task timing in this browser: the timer gap covers it.
  }
};

/** Resolves with `work`'s value, or with `fallback` when the page does not answer in `ms`. */
async function within<T>(work: Promise<T>, ms: number, fallback: T): Promise<T> {
  let timer: ReturnType<typeof setTimeout> | undefined;
  const late = new Promise<T>((resolve) => (timer = setTimeout(() => resolve(fallback), ms)));
  try {
    return await Promise.race([work.catch(() => fallback), late]);
  } finally {
    clearTimeout(timer);
  }
}

const max = (values: readonly number[]) => Math.round(Math.max(0, ...values));

/** Open `deck` in `page` and measure it. Never throws: a closed page is a `CRASH` row. */
async function measure(
  page: Page,
  deck: HangDeck,
  row: Row,
  closed: () => boolean,
  watch: (on: boolean) => void
) {
  const fail = (verdict: Verdict, why: string) => {
    row.verdict = closed() ? 'CRASH' : verdict;
    row.why = closed() ? 'the page or the browser closed' : why;
  };
  try {
    await page.goto('/?deck=basic');
    await page.waitForFunction(
      () => (window as unknown as Probe).__pptxEditor?.snapshot().status === 'ready',
      undefined,
      { timeout: 30_000 }
    );
  } catch {
    return fail('ERROR', 'the demo did not open its own deck');
  }
  // The deck is built before the clock starts, so the build never counts as open time.
  const file = join(DECK_DIR, `${deck.id}.pptx`);
  writeFileSync(file, deck.bytes());
  // The old deck is ready too: a settled open is one that left it first.
  await page.evaluate(() => {
    const w = window as unknown as Probe;
    w.__tasks = [];
    w.__frames = [];
    w.__gaps = [];
    w.__left = false;
    w.__pptxEditor?.subscribe(() => {
      if (w.__pptxEditor?.snapshot().status !== 'ready') w.__left = true;
    });
  });
  const t0 = Date.now();
  watch(true);
  try {
    await page.locator('input[type=file][accept*=".pptx"]').setInputFiles(file);
  } catch {
    return fail('ERROR', 'the file input did not take the deck');
  }
  // Ready: the deck opened with no slide still opening, or the editor refused it.
  const settled = await within(
    page
      .waitForFunction(
        () => {
          const w = window as unknown as Probe;
          const s = w.__pptxEditor?.snapshot();
          if (!s || !w.__left) return false;
          return s.status === 'error' || (s.status === 'ready' && !s.pendingSlides);
        },
        undefined,
        { timeout: BUDGET_MS, polling: 100 }
      )
      .then(() => true),
    BUDGET_MS + 5_000,
    false
  );
  row.openMs = Date.now() - t0;
  if (!settled) fail('HANG', `not ready within ${BUDGET_MS} ms`);
  // Let the idle work (thumbnails, picture decoding, font loads) run, then read the page.
  const answered = await within(
    page.waitForTimeout(IDLE_MS).then(() =>
      page.evaluate(() => {
        const w = window as unknown as Probe;
        const s = w.__pptxEditor?.snapshot();
        return {
          tasks: w.__tasks,
          frames: w.__frames,
          gaps: w.__gaps,
          status: s?.status ?? '',
          refused: s?.error ? JSON.stringify(s.error) : undefined,
        };
      })
    ),
    IDLE_MS + 10_000,
    null
  );
  watch(false);
  if (!answered) return fail('HANG', 'the page did not answer after the open');
  row.status = answered.status;
  if (answered.refused) row.refused = answered.refused;
  row.longestGapMs = max(answered.gaps);
  row.longestTaskMs = max(answered.tasks);
  row.longestFrameMs = max(answered.frames);
  row.blockingMs = Math.round(answered.tasks.reduce((sum, d) => sum + Math.max(0, d - 50), 0));
  if (row.verdict !== 'ok') return;
  const blocked = Math.max(row.longestGapMs, row.longestTaskMs);
  if (blocked > BLOCK_MS) return fail('BLOCK', `the main thread was busy for ${blocked} ms`);
  // A refused deck must still leave the editor usable for the next file.
  const recovered = await within(
    page.evaluate(async () => {
      const bytes = new Uint8Array(await (await fetch('/sample.pptx')).arrayBuffer());
      return (await (window as unknown as Probe).__pptxEditor!.load(bytes)).ok;
    }),
    BUDGET_MS,
    false
  );
  const visible = recovered && (await page.locator('.pptx-surface__slide').first().isVisible());
  if (!visible) return fail('ERROR', 'the editor did not open the sample deck after this one');
}

const decks = HANG_DECKS.filter((d) => !d.stalledFonts && (!ONLY || ONLY.includes(d.id)));
for (const deck of decks) {
  test(deck.id, async ({ browser, browserName }, info) => {
    test.setTimeout(BUDGET_MS + IDLE_MS + 90_000);
    mkdirSync(DECK_DIR, { recursive: true });
    const context = await browser.newContext({ viewport: { width: 1440, height: 900 } });
    const page = await context.newPage();
    await page.addInitScript(OBSERVE);
    let closed = false;
    page.on('crash', () => (closed = true));
    page.on('close', () => (closed = true));
    browser.on('disconnected', () => (closed = true));
    const errors: string[] = [];
    page.on('pageerror', (error) => errors.push(error.message));
    const row: Row = {
      id: deck.id,
      browser: browserName,
      verdict: 'ok',
      fetched: [],
      openMs: 0,
      longestGapMs: 0,
      longestTaskMs: 0,
      longestFrameMs: 0,
      blockingMs: 0,
      status: '',
    };
    let watching = false;
    page.on('request', (request) => {
      const url = request.url();
      const own = /^(blob|data):/.test(url) || new URL(url).origin === new URL(page.url()).origin;
      if (watching && !own) row.fetched?.push(url);
    });
    const cdp =
      browserName === 'chromium'
        ? await context.newCDPSession(page).catch(() => undefined)
        : undefined;
    await cdp?.send('Performance.enable').catch(() => undefined);
    try {
      await measure(
        page,
        deck,
        row,
        () => closed,
        (on) => (watching = on)
      );
      if (cdp && !closed) {
        const { metrics } = await within(cdp.send('Performance.getMetrics'), 5_000, {
          metrics: [] as { name: string; value: number }[],
        });
        const metric = (name: string) => metrics.find((m) => m.name === name)?.value;
        const heap = metric('JSHeapUsedSize');
        if (heap !== undefined) row.heapMB = Math.round(heap / 1e6);
        row.domNodes = metric('Nodes');
      }
    } finally {
      await within(context.close(), 10_000, undefined);
      rmSync(join(DECK_DIR, `${deck.id}.pptx`), { force: true });
    }
    if (!row.fetched?.length) delete row.fetched;
    if (errors.length) row.errors = errors;
    if (row.verdict === 'ok' && errors.length) {
      row.verdict = 'ERROR';
      row.why = `uncaught error: ${errors[0]}`;
    }
    if (row.verdict === 'ok' && row.fetched) {
      row.verdict = 'FETCH';
      row.why = `fetched ${row.fetched.slice(0, 3).join(', ')}`;
    }
    const longest = Math.max(row.longestGapMs, row.longestTaskMs, row.longestFrameMs);
    if (row.verdict === 'ok' && longest > FREEZE_MS) {
      row.verdict = 'FREEZE';
      row.why = `${longest} ms without an answer or a frame`;
    }
    const slow = Math.max(SLOW_MS, deck.slowMs ?? 0);
    if (row.verdict === 'ok' && row.openMs > slow) {
      row.verdict = 'SLOW';
      row.why = `open took ${row.openMs} ms`;
    }
    console.log(
      `${row.verdict.padEnd(6)} ${String(row.openMs).padStart(6)} ms open  ` +
        `${String(row.longestGapMs).padStart(6)} ms gap  ` +
        `${String(row.longestTaskMs).padStart(6)} ms task  ` +
        `${String(row.longestFrameMs).padStart(6)} ms frame  ` +
        (row.heapMB === undefined ? '' : `${String(row.heapMB).padStart(5)} MB  `) +
        `${row.id}${row.refused ? ` (refused ${row.refused})` : ''}` +
        (row.why ? ` (${row.why})` : '')
    );
    mkdirSync(info.project.outputDir, { recursive: true });
    appendFileSync(join(info.project.outputDir, 'rows.jsonl'), `${JSON.stringify(row)}\n`);
    await info.attach('row', {
      body: JSON.stringify(row, null, 2),
      contentType: 'application/json',
    });
    const failing: Verdict[] = ['CRASH', 'ERROR', 'HANG', 'BLOCK', 'FETCH'];
    if (STRICT) failing.push('FREEZE');
    expect(failing.includes(row.verdict), `${row.verdict}: ${row.why}`).toBe(false);
  });
}
