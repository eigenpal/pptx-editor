// Document refresh in a real browser, in the React and the Vue demo: what happy-dom cannot
// show (packages/core/src/editor/__tests__/document-refresh.test.ts). After DOCX Editor's
// e2e/document-refresh.interaction.spec.ts and document-refresh-highlights.interaction.spec.ts.
//
// - An accepted result never shows the loading screen, never empties the canvas or a painted
//   thumbnail, keeps the DOM nodes of what did not change, keeps the slide and the zoom, and
//   runs no long task over 50 ms. A MutationObserver and a sample once a frame watch the page
//   from the `applyUpdate()` call until the follow-up work settles.
// - The change flash: the phases and the times of a thumbnail, the canvas band over changed
//   text, a thumbnail that comes into view or mounts during the hold, reduced motion,
//   `clearHighlights()`, and saved files without the flash.
//
// Screenshots during the flash go to `test-results/` (`REFRESH_SHOTS` sets another folder).

import { mkdirSync } from 'node:fs';
import { join, resolve } from 'node:path';
import { unzipSync, strFromU8 } from 'fflate';
import { expect, test, type Page } from '@playwright/test';
import type { PptxEditorInstance } from '../packages/core/src/editor/index.ts';
import type { DocumentRefresh } from '../packages/core/src/editor/document-refresh-types.ts';
import {
  BODY_BOX,
  SLIDE_SIZE,
  refreshDeck,
  type RefreshEdits,
} from '../scripts/bench/refresh-deck.ts';

declare global {
  interface Window {
    __pptxEditor: PptxEditorInstance;
    __refresh: DocumentRefresh;
    __flash: { t0: number; els: Element[]; log: { el: number; value: string | null; t: number }[] };
  }
}

const ROOT = resolve(import.meta.dirname, '..');
const SHOTS = process.env.REFRESH_SHOTS ?? resolve(ROOT, 'test-results/document-refresh');
/** The most a task may take during a refresh (openspec/changes/document-refresh). */
const LONG_TASK_MS = 50;

/** What the page saw from the `applyUpdate()` call until the work after it settled. */
interface Recording {
  ok: boolean;
  mode: string;
  loadingSeen: number;
  emptyCanvas: number;
  emptiedThumbs: number;
  samples: number;
  /** Unchanged slides whose thumbnail picture is another DOM node after the refresh. */
  replacedUnchanged: string[];
  /** Unchanged slides whose thumbnail picture changed in place. */
  mutatedUnchanged: string[];
  /** Slides whose thumbnail picture changed: in place, or as a new node. */
  repainted: string[];
  /** The slides whose thumbnails show `marker` text after the refresh. */
  showing: string[];
  canvasSame: boolean;
  slideBefore: number;
  slideAfter: number;
  zoomBefore: number;
  zoomAfter: number;
  longTasks: number[];
  applyMs: number;
  canUndo: boolean;
}

/** Open `bytes` in the demo at `base`, and wait until every thumbnail has painted. */
async function openDeck(page: Page, base: string, bytes: Uint8Array): Promise<void> {
  await page.goto(`${base}/?deck=basic`);
  await page.waitForSelector('.pptx-surface__slide', { timeout: 30_000 });
  await page.evaluate(async (data) => {
    const r = await window.__pptxEditor.load(new Uint8Array(data));
    if (!r.ok) throw new Error(JSON.stringify(r));
  }, Array.from(bytes));
  const n = await page.evaluate(() => window.__pptxEditor.snapshot().slides.length);
  await expect(page.locator('.pptx-rail [data-index]')).toHaveCount(n, { timeout: 30_000 });
  // The engine module of the demo itself: the same URL, so the same instance.
  await page.evaluate(async (root) => {
    const url = `/@fs${root}/packages/core/src/editor/index.ts`;
    const core = (await import(/* @vite-ignore */ url)) as {
      createDocumentRefresh(e: PptxEditorInstance): DocumentRefresh;
    };
    window.__refresh = core.createDocumentRefresh(window.__pptxEditor);
  }, ROOT);
  await scrollRail(page, 'bottom');
  await expect
    .poll(
      () =>
        page.evaluate(
          () => document.querySelectorAll('.pptx-rail .pptx-rail__thumb:not(:has(> *))').length
        ),
      { timeout: 30_000 }
    )
    .toBe(0);
  await scrollRail(page, 'top');
  // Idle work (deferred thumbnails, fonts) settles before a measurement.
  await page.waitForTimeout(500);
}

async function scrollRail(page: Page, to: 'top' | 'bottom'): Promise<void> {
  await page.evaluate((to) => {
    const list = document.querySelector<HTMLElement>('.pptx-rail__list');
    if (list) list.scrollTop = to === 'top' ? 0 : list.scrollHeight;
  }, to);
}

/**
 * Capture, apply `bytes` in `mode`, and record the page until `settleMs` after the result.
 * `changed` holds the title keys of the slides the result changes.
 */
function recordApply(
  page: Page,
  bytes: Uint8Array,
  mode: 'reload' | 'edit',
  changed: string[],
  settleMs = 1_000
): Promise<Recording> {
  return page.evaluate(
    async ({ data, mode, changed, settleMs }) => {
      const ed = window.__pptxEditor;
      const refresh = window.__refresh;
      const rail = document.querySelector<HTMLElement>('.pptx-rail__list')!;
      const keyOf = (thumb: Element): string =>
        /Slide \d+ title|Inserted slide/.exec(thumb.textContent ?? '')?.[0] ?? '?';
      const thumbs = () => [...rail.querySelectorAll<HTMLElement>('.pptx-rail__thumb')];
      const before = new Map<string, Element>();
      for (const t of thumbs()) if (t.firstElementChild) before.set(keyOf(t), t.firstElementChild);
      const canvasBefore = document.querySelector('.pptx-surface__slide');
      const slideBefore = ed.snapshot().currentSlide;
      const zoomBefore = ed.snapshot().zoom;
      const submission = await refresh.capture();
      let loadingSeen = 0;
      let emptyCanvas = 0;
      let emptiedThumbs = 0;
      let samples = 0;
      const check = () => {
        samples += 1;
        if (document.querySelector('.pptx-loading')) loadingSeen += 1;
        const slide = document.querySelector('.pptx-surface__slide');
        if (!slide || slide.childElementCount === 0) emptyCanvas += 1;
        const box = rail.getBoundingClientRect();
        for (const t of thumbs()) {
          if (t.firstElementChild) continue;
          const r = t.getBoundingClientRect();
          if (r.bottom > box.top && r.top < box.bottom) emptiedThumbs += 1;
        }
      };
      const longTasks: number[] = [];
      const perf = new PerformanceObserver((list) => {
        for (const e of list.getEntries()) longTasks.push(Math.round(e.duration));
      });
      perf.observe({ type: 'longtask' });
      const touched = new Set<string>();
      const mo = new MutationObserver((records) => {
        for (const r of records) {
          if (r.attributeName?.startsWith('data-refresh') || r.attributeName === 'aria-description')
            continue;
          const node = r.target.nodeType === 1 ? (r.target as Element) : r.target.parentElement;
          const thumb = node?.closest('.pptx-rail__thumb');
          // The thumbnail container itself: its item moved, or a flash mark.
          if (thumb && thumb !== node) touched.add(keyOf(thumb));
          else if (thumb && r.type === 'childList') touched.add(keyOf(thumb));
        }
        check();
      });
      mo.observe(document.body, {
        childList: true,
        subtree: true,
        attributes: true,
        characterData: true,
      });
      let sampling = true;
      const frame = () => {
        check();
        if (sampling) requestAnimationFrame(frame);
      };
      requestAnimationFrame(frame);
      const t0 = performance.now();
      const result = await refresh.applyUpdate({
        submission,
        sequence: 1,
        bytes: new Uint8Array(data),
        mode,
      });
      const applyMs = performance.now() - t0;
      await new Promise((r) => setTimeout(r, settleMs));
      sampling = false;
      mo.disconnect();
      // Long task entries arrive after their task: one more frame and a task.
      await new Promise((r) => requestAnimationFrame(() => setTimeout(r, 50)));
      for (const e of perf.takeRecords()) longTasks.push(Math.round(e.duration));
      perf.disconnect();
      refresh.finish(submission);
      const after = new Map<string, Element | null>();
      for (const t of thumbs()) after.set(keyOf(t), t.firstElementChild);
      const replacedUnchanged: string[] = [];
      const repainted = new Set(touched);
      for (const [key, node] of before) {
        if (!after.has(key)) continue;
        const same = after.get(key) === node;
        if (!same) repainted.add(key);
        if (!changed.includes(key) && !same) replacedUnchanged.push(key);
      }
      const showing = thumbs()
        .filter((t) => t.textContent?.includes('Updated review date'))
        .map(keyOf);
      return {
        ok: result.ok,
        mode: result.ok ? result.mode : result.code,
        loadingSeen,
        emptyCanvas,
        emptiedThumbs,
        samples,
        replacedUnchanged,
        mutatedUnchanged: [...touched].filter((k) => !changed.includes(k)),
        repainted: [...repainted].sort(),
        showing: showing.sort(),
        canvasSame: document.querySelector('.pptx-surface__slide') === canvasBefore,
        slideBefore,
        slideAfter: ed.snapshot().currentSlide,
        zoomBefore,
        zoomAfter: ed.snapshot().zoom,
        longTasks,
        applyMs: Math.round(applyMs),
        canUndo: ed.snapshot().canUndo,
      };
    },
    { data: Array.from(bytes), mode, changed, settleMs }
  );
}

/** Start the flash, and log each `data-refresh-highlight` value of thumbnails with its time. */
async function startFlash(page: Page, options: Record<string, unknown> = {}): Promise<number> {
  return page.evaluate((options) => {
    const els: Element[] = [];
    const idOf = (el: Element) => {
      const id = els.indexOf(el);
      return id >= 0 ? id : els.push(el) - 1;
    };
    const t0 = performance.now();
    const log: Window['__flash']['log'] = [];
    window.__flash = { t0, els, log };
    new MutationObserver((records) => {
      for (const r of records) {
        const el = r.target as Element;
        if (!el.classList?.contains('pptx-rail__thumb')) continue;
        log.push({
          el: idOf(el),
          value: el.getAttribute('data-refresh-highlight'),
          t: performance.now() - t0,
        });
      }
    }).observe(document.body, {
      subtree: true,
      attributes: true,
      attributeFilter: ['data-refresh-highlight'],
    });
    const count = window.__refresh.highlightChanges(options);
    for (const el of document.querySelectorAll('.pptx-rail__thumb[data-refresh-highlight]'))
      log.push({ el: idOf(el), value: el.getAttribute('data-refresh-highlight'), t: 0 });
    return count;
  }, options);
}

/**
 * The phases of each thumbnail in the flash log, with their first times. Thumbnails that left
 * the page (the rail in the slide sorter) are left out.
 */
async function phases(page: Page): Promise<Map<number, { value: string | null; t: number }[]>> {
  const { log, connected } = await page.evaluate(() => ({
    log: window.__flash.log,
    connected: window.__flash.els.map((el) => el.isConnected),
  }));
  const out = new Map<number, { value: string | null; t: number }[]>();
  for (const e of log) {
    if (!connected[e.el]) continue;
    const list = out.get(e.el) ?? [];
    if (list.at(-1)?.value !== e.value) list.push({ value: e.value, t: Math.round(e.t) });
    out.set(e.el, list);
  }
  return out;
}

function shotPath(name: string): string {
  mkdirSync(SHOTS, { recursive: true });
  return join(SHOTS, `${name}.png`);
}

/** The thumbnail of the rail item at `index`. */
const railThumb = (page: Page, index: number) =>
  page.locator(`.pptx-rail [data-index="${index}"] .pptx-rail__thumb`);

for (const adapter of [{ name: 'react' }, { name: 'vue' }] as const) {
  test.describe(`${adapter.name}: document refresh`, () => {
    /** The demo's origin: the configured base URL for React, the Vue server for Vue. */
    const base = (): string =>
      adapter.name === 'vue' ? String(test.info().config.metadata.vueUrl) : '';

    for (const mode of ['reload', 'edit'] as const) {
      test(`${mode}: no loading screen, no empty canvas or thumbnail, kept nodes, no long task`, async ({
        page,
      }) => {
        const errors: string[] = [];
        page.on('pageerror', (e) => errors.push(e.message));
        await openDeck(page, base(), refreshDeck(20));
        await page.evaluate(() => {
          window.__pptxEditor.goToSlide(3);
          window.__pptxEditor.setZoom(1.25);
        });
        await page.waitForTimeout(300);
        const edits: RefreshEdits = { text: [3, 7], insertAfter: 10, layout: [5] };
        const rec = await recordApply(page, refreshDeck(20, edits), mode, [
          'Inserted slide',
          'Slide 4 title',
          'Slide 6 title',
          'Slide 8 title',
        ]);
        console.log(`[${adapter.name} ${mode}]`, JSON.stringify(rec));
        expect(rec.ok).toBe(true);
        expect(rec.mode).toBe(mode);
        expect(rec.samples).toBeGreaterThan(3);
        expect(rec.loadingSeen).toBe(0);
        expect(rec.emptyCanvas).toBe(0);
        expect(rec.emptiedThumbs).toBe(0);
        expect(rec.replacedUnchanged).toEqual([]);
        // The thumbnails of the slides after the inserted one show other slides now: the rail
        // keeps each item by slide id, but its thumbnail paints again at the new index, in
        // place. Before the inserted slide, nothing unchanged paints again.
        const shifted = (key: string) => Number(/\d+/.exec(key)?.[0]) > 11;
        expect(rec.mutatedUnchanged.filter((k) => !shifted(k))).toEqual([]);
        // The slides with new text painted again and show it. Slide 6 changed only its layout,
        // which paints the same, so its thumbnail may stay as it is.
        expect(rec.repainted).toEqual(
          expect.arrayContaining(['Inserted slide', 'Slide 4 title', 'Slide 8 title'])
        );
        expect(rec.showing).toEqual(['Slide 4 title', 'Slide 8 title']);
        expect(rec.slideAfter).toBe(rec.slideBefore);
        expect(rec.zoomAfter).toBe(rec.zoomBefore);
        expect(rec.canUndo).toBe(mode === 'edit');
        expect(await page.locator('.pptx-rail [data-index]').count()).toBe(21);
        expect(rec.longTasks.filter((d) => d > LONG_TASK_MS)).toEqual([]);
        expect(errors).toEqual([]);
      });
    }

    test('100 slides: no long task in either mode', async ({ page }) => {
      // Before the steps ran in tasks of their own, a reload of this deck took one 66 ms task.
      const edits: RefreshEdits = { text: [1], insertAfter: 50 };
      for (const mode of ['reload', 'edit'] as const) {
        await openDeck(page, base(), refreshDeck(100));
        await page.evaluate(() => window.__pptxEditor.goToSlide(1));
        await page.waitForTimeout(300);
        const rec = await recordApply(page, refreshDeck(100, edits), mode, [
          'Inserted slide',
          'Slide 2 title',
        ]);
        console.log(`[${adapter.name} ${mode} 100]`, JSON.stringify(rec.longTasks), rec.applyMs);
        expect(rec.ok).toBe(true);
        expect(rec.loadingSeen + rec.emptyCanvas + rec.emptiedThumbs).toBe(0);
        expect(rec.longTasks.filter((d) => d > LONG_TASK_MS)).toEqual([]);
      }
    });

    test('an unchanged current slide keeps its canvas node', async ({ page }) => {
      await openDeck(page, base(), refreshDeck(12));
      await page.evaluate(() => window.__pptxEditor.goToSlide(1));
      await page.waitForTimeout(200);
      const rec = await recordApply(page, refreshDeck(12, { text: [6] }), 'reload', [
        'Slide 7 title',
      ]);
      expect(rec.ok).toBe(true);
      expect(rec.canvasSame).toBe(true);
      expect(rec.replacedUnchanged).toEqual([]);
      expect(rec.mutatedUnchanged).toEqual([]);
      expect(rec.repainted).toEqual(['Slide 7 title']);
    });

    test('flash: thumbnail phases and times, canvas band over the changed text', async ({
      page,
    }) => {
      await openDeck(page, base(), refreshDeck(12));
      await page.evaluate(() => window.__pptxEditor.goToSlide(2));
      await page.waitForTimeout(200);
      const rec = await recordApply(page, refreshDeck(12, { text: [2, 5] }), 'reload', [
        'Slide 3 title',
        'Slide 6 title',
      ]);
      expect(rec.ok).toBe(true);
      const boxesBefore = await page.evaluate(() =>
        [...document.querySelectorAll('.pptx-rail [data-index], .pptx-surface__slide')].map((e) =>
          JSON.stringify(e.getBoundingClientRect())
        )
      );
      expect(await startFlash(page)).toBe(2);
      await expect(railThumb(page, 2)).toHaveAttribute('data-refresh-highlight', /entering|shown/);
      await page.waitForTimeout(60);
      await page.screenshot({ path: shotPath(`${adapter.name}-flash-entering`) });
      await page.waitForTimeout(500);
      await page.screenshot({ path: shotPath(`${adapter.name}-flash-shown`) });
      // The flash never moves the rail or the slide.
      const boxesDuring = await page.evaluate(() =>
        [...document.querySelectorAll('.pptx-rail [data-index], .pptx-surface__slide')].map((e) =>
          JSON.stringify(e.getBoundingClientRect())
        )
      );
      expect(boxesDuring).toEqual(boxesBefore);
      // The band over the edited paragraph: inside the body box, on its first line.
      const canvas = await page.evaluate(
        ({ body, size }) => {
          const stage = document.querySelector('.pptx-surface__stage')!;
          const layer = stage.querySelector<SVGSVGElement>(':scope > .pptx-surface__refresh');
          const slide = stage.querySelector('.pptx-surface__slide')!.getBoundingClientRect();
          const bands = [...(layer?.querySelectorAll('.pptx-refresh-band') ?? [])].map((b) =>
            b.getBoundingClientRect()
          );
          const sx = slide.width / size.width;
          const sy = slide.height / size.height;
          return {
            phase: layer?.getAttribute('data-refresh-highlight'),
            opacity: layer ? Number(getComputedStyle(layer).opacity) : -1,
            belowOverlay:
              !!layer && layer.nextElementSibling?.classList.contains('pptx-surface__overlay'),
            bands: bands.map((r) => ({
              left: (r.left - slide.left) / sx,
              top: (r.top - slide.top) / sy,
              right: (r.right - slide.left) / sx,
              bottom: (r.bottom - slide.top) / sy,
            })),
            body,
            fill: layer?.querySelector('.pptx-refresh-band')
              ? getComputedStyle(layer.querySelector('.pptx-refresh-band')!).fill
              : '',
          };
        },
        { body: BODY_BOX, size: SLIDE_SIZE }
      );
      console.log(`[${adapter.name} canvas]`, JSON.stringify(canvas));
      expect(canvas.phase).toBe('shown');
      expect(canvas.opacity).toBe(1);
      expect(canvas.belowOverlay).toBe(true);
      expect(canvas.bands.length).toBeGreaterThan(0);
      const emu = 9525 * 6; // Six CSS pixels of tolerance at 100% zoom.
      for (const b of canvas.bands) {
        expect(b.left).toBeGreaterThanOrEqual(BODY_BOX.x - emu);
        expect(b.right).toBeLessThanOrEqual(BODY_BOX.x + BODY_BOX.cx + emu);
        expect(b.top).toBeGreaterThanOrEqual(BODY_BOX.y - emu);
        // One line of 18 pt text: under one inch below the box top.
        expect(b.bottom).toBeLessThanOrEqual(BODY_BOX.y + 914_400);
        expect(b.right - b.left).toBeGreaterThan(914_400);
      }
      expect(canvas.fill).toMatch(/rgb\(227, 116, 0\)/);
      await page.waitForTimeout(3_400 - 560);
      const byEl = await phases(page);
      const lists = [...byEl.values()];
      console.log(`[${adapter.name} phases]`, JSON.stringify(lists));
      // Two rail thumbnails, each entering, shown, exiting, then removed.
      expect(lists.length).toBe(2);
      for (const list of lists) {
        expect(list.map((p) => p.value)).toEqual(['entering', 'shown', 'exiting', null]);
        const [, shown, exiting, removed] = list;
        expect(shown.t).toBeLessThan(100);
        expect(exiting.t).toBeGreaterThanOrEqual(2_990);
        expect(exiting.t).toBeLessThan(3_150);
        expect(removed.t - exiting.t).toBeGreaterThanOrEqual(170);
        expect(removed.t - exiting.t).toBeLessThan(330);
      }
      // One start time: both thumbnails change phase together.
      expect(Math.abs(lists[0][2].t - lists[1][2].t)).toBeLessThan(20);
      await expect(page.locator('[data-refresh-highlight]')).toHaveCount(0);
      await expect(page.locator('.pptx-surface__refresh')).toHaveCount(0);
      expect(await page.evaluate(() => window.__refresh.snapshot().highlightsVisible)).toBe(false);
    });

    test('flash: a thumbnail scrolled into view or mounted during the hold shows the rest only', async ({
      page,
    }) => {
      await openDeck(page, base(), refreshDeck(40));
      const rec = await recordApply(page, refreshDeck(40, { text: [0, 36] }), 'reload', [
        'Slide 1 title',
        'Slide 37 title',
      ]);
      expect(rec.ok).toBe(true);
      await scrollRail(page, 'top');
      const offscreen = railThumb(page, 36);
      await expect(offscreen).not.toBeInViewport();
      expect(await startFlash(page)).toBe(2);
      await page.waitForTimeout(800);
      await offscreen.scrollIntoViewIfNeeded();
      const seen = await offscreen.evaluate((el) => ({
        phase: el.getAttribute('data-refresh-highlight'),
        running: el.getAnimations({ subtree: true }).filter((a) => a.playState === 'running')
          .length,
        wash: Number(getComputedStyle(el, '::after').opacity),
        painted: el.firstElementChild?.textContent?.includes('Updated review date') ?? false,
      }));
      expect(seen).toEqual({ phase: 'shown', running: 0, wash: 1, painted: true });
      // The compositor rasterizes the newly exposed part of the rail in the next frames.
      await page.waitForTimeout(100);
      await page.screenshot({ path: shotPath(`${adapter.name}-flash-scrolled`) });
      // A view that mounts thumbnails during the hold: the slide sorter.
      await page.evaluate(() =>
        window.__pptxEditor.exec({ type: 'setPresentationView', view: 'slideSorter' })
      );
      await page.waitForTimeout(300);
      const sorter = await page.evaluate(() => {
        const marked = [...document.querySelectorAll('.pptx-rail__thumb[data-refresh-highlight]')];
        return marked.map((el) => ({
          phase: el.getAttribute('data-refresh-highlight'),
          running: el.getAnimations({ subtree: true }).filter((a) => a.playState === 'running')
            .length,
        }));
      });
      console.log(`[${adapter.name} sorter]`, JSON.stringify(sorter));
      expect(sorter.length).toBeGreaterThan(0);
      for (const s of sorter) expect(s).toEqual({ phase: 'shown', running: 0 });
      await page.screenshot({ path: shotPath(`${adapter.name}-flash-sorter`) });
      await page.waitForTimeout(3_400 - 1_100);
      const lists = [...(await phases(page)).values()];
      console.log(`[${adapter.name} late phases]`, JSON.stringify(lists));
      for (const list of lists) {
        // Thumbnails that mounted late never enter; every one exits at the shared time.
        if (list[0].t > 0) expect(list[0].value).toBe('shown');
        const exiting = list.find((p) => p.value === 'exiting');
        expect(exiting?.t).toBeGreaterThanOrEqual(2_990);
        expect(exiting?.t).toBeLessThan(3_150);
      }
      await expect(page.locator('[data-refresh-highlight]')).toHaveCount(0);
    });

    test('flash: reduced motion caps fades at 125 ms; clearHighlights fades and removes', async ({
      page,
    }) => {
      await openDeck(page, base(), refreshDeck(8));
      await page.evaluate(() => window.__pptxEditor.goToSlide(1));
      const rec = await recordApply(page, refreshDeck(8, { text: [1] }), 'edit', ['Slide 2 title']);
      expect(rec.ok).toBe(true);
      await page.emulateMedia({ reducedMotion: 'reduce' });
      await startFlash(page, { animation: { durationMs: 400, exitDurationMs: 600 } });
      const reduced = await page.evaluate(() => {
        const thumb = document.querySelector<HTMLElement>(
          '.pptx-rail [data-index="1"] .pptx-rail__thumb'
        )!;
        const layer = document.querySelector<SVGSVGElement>('.pptx-surface__refresh')!;
        return {
          enter: thumb.style.getPropertyValue('--pptx-refresh-enter'),
          exit: thumb.style.getPropertyValue('--pptx-refresh-exit'),
          layerTransition: getComputedStyle(layer).transitionDuration,
          afterTransition: getComputedStyle(thumb, '::after').transitionDuration,
        };
      });
      expect(reduced).toEqual({
        enter: '125ms',
        exit: '125ms',
        layerTransition: '0.125s',
        afterTransition: '0.125s',
      });
      // clearHighlights() fades out over the capped exit, then removes every mark.
      await page.waitForTimeout(300);
      const t = await page.evaluate(async () => {
        const t0 = performance.now();
        window.__refresh.clearHighlights();
        const exiting = document.querySelectorAll('[data-refresh-highlight="exiting"]').length;
        while (document.querySelector('[data-refresh-highlight]'))
          await new Promise((r) => requestAnimationFrame(r));
        return { exiting, ms: performance.now() - t0 };
      });
      expect(t.exiting).toBeGreaterThan(0);
      expect(t.ms).toBeGreaterThanOrEqual(120);
      expect(t.ms).toBeLessThan(300);
      await expect(page.locator('.pptx-surface__refresh')).toHaveCount(0);
      // Without animation, the marks go at once.
      await page.emulateMedia({ reducedMotion: 'no-preference' });
      await startFlash(page, { timeoutMs: null });
      const gone = await page.evaluate(() => {
        window.__refresh.clearHighlights({ animation: false });
        return document.querySelectorAll('[data-refresh-highlight], .pptx-surface__refresh').length;
      });
      expect(gone).toBe(0);
      expect(await page.evaluate(() => window.__refresh.snapshot().highlightsVisible)).toBe(false);
    });

    test('flash: a saved file never holds the flash', async ({ page }) => {
      await openDeck(page, base(), refreshDeck(6));
      await page.evaluate(() => window.__pptxEditor.goToSlide(2));
      const rec = await recordApply(page, refreshDeck(6, { text: [2] }), 'reload', [
        'Slide 3 title',
      ]);
      expect(rec.ok).toBe(true);
      const clean = await page.evaluate(() => Array.from(window.__pptxEditor.save()!));
      await startFlash(page, { timeoutMs: null });
      await page.waitForTimeout(300);
      const during = await page.evaluate(() => Array.from(window.__pptxEditor.save()!));
      await page.evaluate(() => window.__refresh.clearHighlights({ animation: false }));
      const a = unzipSync(new Uint8Array(clean));
      const b = unzipSync(new Uint8Array(during));
      expect(Object.keys(b).sort()).toEqual(Object.keys(a).sort());
      for (const [name, bytes] of Object.entries(b)) {
        const text = strFromU8(bytes);
        expect(text).not.toMatch(/refresh/i);
        expect(text, name).toBe(strFromU8(a[name]!));
      }
    });
  });
}
