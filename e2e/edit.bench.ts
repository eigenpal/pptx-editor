// Browser editing bench: main-thread time per keystroke, per drag step, and per slide switch in
// the demo (React chrome, real rail), on the synthetic decks of `scripts/bench/edit-bench.ts`.
//
// Main-thread time is Chrome's `TaskDuration` metric (DevTools protocol), less the same span of
// idle time. It includes event dispatch, the engine, React, style, layout, and paint, and the
// thumbnail repaints the editor defers until typing stops. `ScriptDuration` is the JavaScript
// part; `RecalcStyleDuration` plus `LayoutDuration` is the style and layout part.
//
// Usage: bun run bench:edit:browser            (sizes from BENCH_SIZES, default 10,60,200)

import { test, type CDPSession, type Page } from '@playwright/test';
import { textDeck } from '../scripts/bench/text-deck.ts';

const SIZES = (process.env.BENCH_SIZES ?? '10,60,200').split(',').map(Number);
const KEYS = 30;
const STEPS = 30;
const SWITCHES = 20;

interface Times {
  task: number;
  script: number;
  styleLayout: number;
}

interface Row extends Times {
  scenario: string;
  slides: number;
  thumbnailCalls: number;
  thumbnailPaints: number;
}

async function times(cdp: CDPSession): Promise<Times> {
  const { metrics } = (await cdp.send('Performance.getMetrics')) as {
    metrics: { name: string; value: number }[];
  };
  const ms = (name: string): number => (metrics.find((m) => m.name === name)?.value ?? 0) * 1000;
  return {
    task: ms('TaskDuration'),
    script: ms('ScriptDuration'),
    styleLayout: ms('RecalcStyleDuration') + ms('LayoutDuration'),
  };
}

const minus = (a: Times, b: Times): Times => ({
  task: a.task - b.task,
  script: a.script - b.script,
  styleLayout: a.styleLayout - b.styleLayout,
});

/** Main-thread time of `work` per operation, less the idle time of a span as long. */
async function mainThread(
  page: Page,
  cdp: CDPSession,
  operations: number,
  work: () => Promise<void>
): Promise<Times> {
  const t0 = Date.now();
  const before = await times(cdp);
  await work();
  await page.waitForTimeout(1_000);
  const busy = minus(await times(cdp), before);
  const span = Date.now() - t0;
  const idle0 = await times(cdp);
  await page.waitForTimeout(span);
  const idle = minus(await times(cdp), idle0);
  const per = (k: keyof Times): number => Math.max(0, busy[k] - idle[k]) / operations;
  return { task: per('task'), script: per('script'), styleLayout: per('styleLayout') };
}

/** Count `renderThumbnail` calls and rail repaints. */
async function instrument(page: Page): Promise<void> {
  await page.evaluate(() => {
    const w = window as unknown as {
      __pptxEditor: { renderThumbnail: (i: number, c: HTMLElement) => void };
      __thumbCalls: number;
      __thumbPaints: number;
    };
    const ed = w.__pptxEditor;
    w.__thumbCalls = 0;
    w.__thumbPaints = 0;
    const orig = ed.renderThumbnail.bind(ed);
    ed.renderThumbnail = (i, c) => {
      w.__thumbCalls += 1;
      orig(i, c);
    };
    new MutationObserver((records) => {
      for (const r of records)
        if ((r.target as Element).classList?.contains('pptx-rail__thumb') && r.addedNodes.length)
          w.__thumbPaints += 1;
    }).observe(document.querySelector('.pptx-rail')!, { childList: true, subtree: true });
  });
}

const counters = (page: Page) =>
  page.evaluate(() => {
    const w = window as unknown as { __thumbCalls: number; __thumbPaints: number };
    const out = { calls: w.__thumbCalls, paints: w.__thumbPaints };
    w.__thumbCalls = 0;
    w.__thumbPaints = 0;
    return out;
  });

async function open(page: Page, n: number): Promise<void> {
  await page.goto('/');
  await page.waitForSelector('.pptx-surface__slide', { timeout: 30_000 });
  await page.locator('input[type=file][accept*=".pptx"]').setInputFiles({
    name: `bench-${n}.pptx`,
    mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    buffer: Buffer.from(textDeck(n)),
  });
  await page.waitForFunction(
    (count) => document.querySelectorAll('.pptx-rail__item').length === count,
    n,
    { timeout: 120_000 }
  );
  await page.waitForTimeout(1_500);
}

/** The middle slide's body box: its node id and its top border in page coordinates. */
async function bodyOf(page: Page, index: number) {
  return page.evaluate((i) => {
    type Item = {
      kind: string;
      matrix: number[];
      element: { name?: string; nodeId: string; xfrm: { cx: number } };
      text?: { paragraphLengths: number[] };
    };
    const ed = (window as unknown as Record<string, unknown>).__pptxEditor as {
      session: { layout(i: number): { items: Item[] } };
    };
    const item = ed.session.layout(i).items.find((x) => x.element.name === 'Body')!;
    return {
      id: item.element.nodeId,
      end: item.text?.paragraphLengths[0] ?? 0,
      x: item.matrix[4] + item.element.xfrm.cx / 2,
      y: item.matrix[5] + 2_000,
    };
  }, index);
}

test('edit bench', async ({ page }) => {
  const rows: Row[] = [];
  const cdp = await page.context().newCDPSession(page);
  await cdp.send('Performance.enable');
  for (const n of SIZES) {
    await open(page, n);
    await instrument(page);
    const middle = Math.floor(n / 2);
    await page.evaluate((i) => {
      const ed = (window as unknown as Record<string, unknown>).__pptxEditor as {
        goToSlide(i: number): void;
      };
      ed.goToSlide(i);
    }, middle);
    const body = await bodyOf(page, middle);

    // Keystrokes at the end of the body text, typed at 25 characters per second.
    await page.evaluate((b) => {
      const ed = (window as unknown as Record<string, unknown>).__pptxEditor as {
        setSelection(s: unknown): void;
      };
      const at = { paragraph: 0, offset: b.end };
      ed.setSelection({ kind: 'text', nodeId: b.id, anchor: at, focus: at });
    }, body);
    await page.locator('.pptx-surface__input').focus();
    await page.waitForTimeout(500);
    await counters(page);
    const typing = await mainThread(page, cdp, KEYS, () =>
      page.keyboard.type('x'.repeat(KEYS), { delay: 40 })
    );
    let c = await counters(page);
    rows.push({
      scenario: 'keystroke',
      slides: n,
      ...typing,
      thumbnailCalls: c.calls / KEYS,
      thumbnailPaints: c.paints / KEYS,
    });

    // Drag steps of the body box by its top border.
    await page.keyboard.press('Escape');
    const stage = (await page.locator('.pptx-surface__stage').boundingBox())!;
    const size = await page.evaluate(() => {
      const ed = (window as unknown as Record<string, unknown>).__pptxEditor as {
        snapshot(): { slideWidth: number };
      };
      return ed.snapshot().slideWidth;
    });
    const px = stage.x + (body.x / size) * stage.width;
    const py = stage.y + (body.y / size) * stage.width;
    await page.mouse.move(px, py);
    await page.mouse.down();
    await page.mouse.move(px + 4, py + 2);
    await page.waitForTimeout(300);
    await counters(page);
    const drag = await mainThread(page, cdp, STEPS, async () => {
      for (let k = 1; k <= STEPS; k += 1) {
        await page.mouse.move(px + 4 + k * 3, py + 2 + k * 2);
        await page.waitForTimeout(16);
      }
    });
    await page.mouse.up();
    c = await counters(page);
    rows.push({
      scenario: 'drag-step',
      slides: n,
      ...drag,
      thumbnailCalls: c.calls / STEPS,
      thumbnailPaints: c.paints / STEPS,
    });

    // Slide switches between two neighbors.
    await page.waitForTimeout(500);
    await counters(page);
    const switching = await mainThread(page, cdp, SWITCHES, async () => {
      for (let k = 0; k < SWITCHES; k += 1) {
        await page.evaluate(
          (i) =>
            (
              (window as unknown as Record<string, unknown>).__pptxEditor as {
                goToSlide(i: number): void;
              }
            ).goToSlide(i),
          k % 2 ? middle : Math.min(n - 1, middle + 1)
        );
        await page.waitForTimeout(50);
      }
    });
    c = await counters(page);
    rows.push({
      scenario: 'slide-switch',
      slides: n,
      ...switching,
      thumbnailCalls: c.calls / SWITCHES,
      thumbnailPaints: c.paints / SWITCHES,
    });
  }
  const lines = [
    '| Scenario | Slides | Main-thread ms | Script ms | Style and layout ms | Thumbnail calls | Thumbnail paints |',
    '| --- | --: | --: | --: | --: | --: | --: |',
    ...rows.map(
      (r) =>
        `| ${r.scenario} | ${r.slides} | ${r.task.toFixed(2)} | ${r.script.toFixed(2)} | ${r.styleLayout.toFixed(2)} | ${r.thumbnailCalls.toFixed(2)} | ${r.thumbnailPaints.toFixed(2)} |`
    ),
  ];
  console.log(lines.join('\n'));
});
