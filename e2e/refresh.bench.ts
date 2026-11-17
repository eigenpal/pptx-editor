// Browser document refresh bench (`bun run bench:refresh:browser`): one accepted result in the
// React demo, in both modes, on the synthetic decks of `scripts/bench/refresh-deck.ts`.
//
// Each run opens the deck, waits until every thumbnail has painted, captures, and applies a
// returned file with one changed paragraph on the current slide and one inserted slide. The row
// reports, as median and maximum over the runs:
//
// - `apply ms`: wall time from the `applyUpdate()` call until its promise resolves.
// - `longest task ms`: the longest long task (PerformanceObserver `longtask`) from the call
//   until one second after it resolves, so the rail and canvas repaints that follow count too.
// - `dropped frames`: frames that animation frame callbacks missed in the same span, from the
//   gaps between callbacks at 60 Hz.
//
// The spec asks for no task over 50 ms (openspec/changes/document-refresh). The bench prints the
// numbers; `e2e/document-refresh.spec.ts` asserts the budget on a 20-slide deck.
//
// Usage: bun run bench:refresh:browser   (sizes from BENCH_SIZES, default 20,100; BENCH_RUNS)

import { test, type Page } from '@playwright/test';
import { resolve } from 'node:path';
import { refreshDeck } from '../scripts/bench/refresh-deck.ts';

const SIZES = (process.env.BENCH_SIZES ?? '20,100').split(',').map(Number);
const RUNS = Number(process.env.BENCH_RUNS ?? 5);
const ROOT = resolve(import.meta.dirname, '..');
const FRAME_MS = 1000 / 60;

interface Sample {
  applyMs: number;
  longestTask: number;
  tasks: number[];
  droppedFrames: number;
}

async function open(page: Page, n: number): Promise<void> {
  await page.goto('/?deck=basic');
  await page.waitForSelector('.pptx-surface__slide', { timeout: 30_000 });
  await page.evaluate(
    async (data) => {
      const ed = (window as unknown as Record<string, unknown>).__pptxEditor as {
        load(b: Uint8Array): Promise<{ ok: boolean }>;
      };
      if (!(await ed.load(new Uint8Array(data))).ok) throw new Error('the deck did not open');
    },
    Array.from(refreshDeck(n))
  );
  await page.waitForFunction(
    (count) => document.querySelectorAll('.pptx-rail [data-index]').length === count,
    n,
    { timeout: 120_000 }
  );
  await page.evaluate(async (root) => {
    const w = window as unknown as Record<string, unknown>;
    const url = `/@fs${root}/packages/core/src/editor/index.ts`;
    const core = (await import(/* @vite-ignore */ url)) as {
      createDocumentRefresh(e: unknown): unknown;
    };
    w.__refresh = core.createDocumentRefresh(w.__pptxEditor);
    (w.__pptxEditor as { goToSlide(i: number): void }).goToSlide(1);
  }, ROOT);
  // The thumbnails in view: past `MAX_PAINTED` (thumbnails.ts), the rest are not all painted.
  await page.waitForFunction(
    () => {
      const box = document.querySelector('.pptx-rail__list')!.getBoundingClientRect();
      return [...document.querySelectorAll('.pptx-rail .pptx-rail__thumb')].every((t) => {
        const r = t.getBoundingClientRect();
        return t.firstElementChild || r.bottom <= box.top || r.top >= box.bottom;
      });
    },
    undefined,
    { timeout: 120_000 }
  );
  await page.waitForTimeout(1_000);
}

async function applyOnce(page: Page, bytes: Uint8Array, mode: 'reload' | 'edit') {
  return page.evaluate(
    async ({ data, mode, frameMs }): Promise<Sample> => {
      type Refresh = {
        capture(): Promise<unknown>;
        applyUpdate(u: unknown): Promise<{ ok: boolean; code?: string }>;
        finish(s: unknown): void;
      };
      const refresh = (window as unknown as { __refresh: Refresh }).__refresh;
      const submission = await refresh.capture();
      const tasks: number[] = [];
      const perf = new PerformanceObserver((list) => {
        for (const e of list.getEntries()) tasks.push(e.duration);
      });
      perf.observe({ type: 'longtask' });
      let last = performance.now();
      let dropped = 0;
      let sampling = true;
      const frame = (now: number) => {
        dropped += Math.max(0, Math.round((now - last) / frameMs) - 1);
        last = now;
        if (sampling) requestAnimationFrame(frame);
      };
      // Start from a frame boundary.
      await new Promise((r) => requestAnimationFrame(r));
      last = performance.now();
      requestAnimationFrame(frame);
      const t0 = performance.now();
      const result = await refresh.applyUpdate({
        submission,
        sequence: 1,
        bytes: new Uint8Array(data),
        mode,
      });
      const applyMs = performance.now() - t0;
      if (!result.ok) throw new Error(`refresh failed: ${result.code}`);
      await new Promise((r) => setTimeout(r, 1_000));
      sampling = false;
      await new Promise((r) => requestAnimationFrame(() => setTimeout(r, 50)));
      for (const e of perf.takeRecords()) tasks.push(e.duration);
      perf.disconnect();
      refresh.finish(submission);
      return {
        applyMs,
        longestTask: Math.max(0, ...tasks),
        tasks: tasks.map(Math.round),
        droppedFrames: dropped,
      };
    },
    { data: Array.from(bytes), mode, frameMs: FRAME_MS }
  );
}

const median = (xs: number[]): number => {
  const s = [...xs].sort((a, b) => a - b);
  return s[Math.floor(s.length / 2)] ?? 0;
};
const cell = (xs: number[]): string => `${median(xs).toFixed(0)} / ${Math.max(...xs).toFixed(0)}`;

test('refresh bench', async ({ page }) => {
  const rows: string[] = [];
  for (const n of SIZES)
    for (const mode of ['reload', 'edit'] as const) {
      const samples: Sample[] = [];
      const next = refreshDeck(n, { text: [1], insertAfter: Math.floor(n / 2) });
      for (let run = 0; run < RUNS; run++) {
        await open(page, n);
        samples.push(await applyOnce(page, next, mode));
      }
      console.log(`${mode} ${n}: tasks ${JSON.stringify(samples.map((s) => s.tasks))}`);
      rows.push(
        `| ${mode} | ${n} | ${cell(samples.map((s) => s.applyMs))} | ` +
          `${cell(samples.map((s) => s.longestTask))} | ` +
          `${cell(samples.map((s) => s.droppedFrames))} |`
      );
    }
  console.log(
    [
      '',
      '| Mode | Slides | apply ms | longest task ms | dropped frames |',
      '| --- | --: | --: | --: | --: |',
      ...rows,
      '',
      `Each cell: median / max over ${RUNS} runs.`,
    ].join('\n')
  );
});
