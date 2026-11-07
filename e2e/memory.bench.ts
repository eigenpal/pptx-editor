// Browser memory bench: heap, DOM nodes, and listeners across repeated work in the demo (React
// chrome, real rail), to find leaks and unbounded caches. Each scenario repeats one piece of
// work and takes a reading after each iteration (or each chunk of iterations).
//
// A reading forces garbage collection (`HeapProfiler.collectGarbage`), then reads Chrome's
// `JSHeapUsedSize`, `Nodes` (live DOM nodes, detached ones too), `JSEventListeners`, and
// `Documents` (DevTools protocol), the array buffer backing stores (`Runtime.getHeapUsage`), and
// the page probe: live blob URLs, `document.fonts` faces, connected observers, workers, timers,
// listeners on `window` and `document`, and the SVG `image` elements in the document.
//
// The verdict compares the growth per iteration over the second half of the readings with the
// scenario's budget: a bounded cache levels off, a leak keeps its slope. With `MEMORY_SNAPSHOTS=1`
// (or on a leak verdict) it also takes a heap snapshot at the baseline and at the end, and
// prints the constructors that grew most, with the detached DOM nodes.
//
// Usage: bun run bench:memory:browser                  (every scenario)
//        bun run bench:memory:browser -g "open-same"    (one scenario)
//        MEMORY_SCALE=0.25 bun run bench:memory:browser (fewer iterations)

import { execFileSync } from 'node:child_process';
import { mkdirSync, readFileSync, writeFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { expect, test, type CDPSession, type Page } from '@playwright/test';
import { textDeck } from '../scripts/bench/text-deck.ts';
import { pictureDeck } from '../scripts/bench/picture-deck.ts';

const ROOT = fileURLToPath(new URL('..', import.meta.url)).replace(/\/$/, '');
const VUE_URL = `http://localhost:${Number(process.env.MEMORY_PORT ?? 5520) + 1}/`;
const SCALE = Number(process.env.MEMORY_SCALE ?? 1);
const SNAPSHOTS = process.env.MEMORY_SNAPSHOTS === '1';
const OUT = `${ROOT}/test-results/memory`;
const n = (count: number): number => Math.max(2, Math.round(count * SCALE));
const publicDeck = (name: string): Uint8Array =>
  new Uint8Array(readFileSync(`${ROOT}/examples/vite/public/${name}`));

// ---------------------------------------------------------------------------------------------
// The page probe, installed before the demo's own scripts.

interface Probe {
  blobs: number;
  fonts: number;
  observers: number;
  workers: number;
  bitmaps: number;
  intervals: number;
  windowListeners: number;
  documentListeners: number;
  images: number;
}

function installProbe(): void {
  const w = window as unknown as { __memProbe: () => Probe };
  const blobs = new Set<string>();
  const create = URL.createObjectURL.bind(URL);
  const revoke = URL.revokeObjectURL.bind(URL);
  URL.createObjectURL = (o: Blob | MediaSource) => {
    const url = create(o);
    blobs.add(url);
    return url;
  };
  URL.revokeObjectURL = (url: string) => {
    blobs.delete(url);
    revoke(url);
  };

  // Observers that observe something and are neither disconnected nor collected.
  let observers = 0;
  const gone = new FinalizationRegistry<null>(() => (observers -= 1));
  for (const name of ['ResizeObserver', 'MutationObserver', 'IntersectionObserver'] as const) {
    const proto = (window[name] as { prototype: { observe: unknown; disconnect: unknown } })
      ?.prototype;
    if (!proto) continue;
    const live = new WeakSet<object>();
    const observe = proto.observe as (...a: unknown[]) => void;
    const disconnect = proto.disconnect as () => void;
    proto.observe = function (this: object, ...a: unknown[]) {
      if (!live.has(this)) {
        live.add(this);
        observers += 1;
        gone.register(this, null, this);
      }
      return observe.apply(this, a);
    };
    proto.disconnect = function (this: object) {
      if (live.delete(this)) {
        observers -= 1;
        gone.unregister(this);
      }
      return disconnect.call(this);
    };
  }

  let workers = 0;
  const WorkerCtor = window.Worker;
  if (WorkerCtor) {
    window.Worker = class extends WorkerCtor {
      #alive = true;
      constructor(...a: ConstructorParameters<typeof Worker>) {
        super(...a);
        workers += 1;
      }
      override terminate(): void {
        if (this.#alive) workers -= 1;
        this.#alive = false;
        super.terminate();
      }
    };
  }

  let bitmaps = 0;
  const bitmapGone = new FinalizationRegistry<null>(() => (bitmaps -= 1));
  const makeBitmap = window.createImageBitmap.bind(window) as (
    ...a: unknown[]
  ) => Promise<ImageBitmap>;
  (window as unknown as { createImageBitmap: unknown }).createImageBitmap = async (
    ...a: unknown[]
  ) => {
    const bitmap = await makeBitmap(...a);
    bitmaps += 1;
    bitmapGone.register(bitmap, null, bitmap);
    return bitmap;
  };
  const close = ImageBitmap.prototype.close;
  ImageBitmap.prototype.close = function (this: ImageBitmap) {
    if (bitmapGone.unregister(this)) bitmaps -= 1;
    close.call(this);
  };

  const intervals = new Set<number>();
  const setI = window.setInterval.bind(window);
  const clearI = window.clearInterval.bind(window);
  (window as unknown as { setInterval: unknown }).setInterval = (...a: unknown[]) => {
    const id = (setI as (...b: unknown[]) => number)(...a);
    intervals.add(id);
    return id;
  };
  (window as unknown as { clearInterval: unknown }).clearInterval = (id?: number) => {
    if (id !== undefined) intervals.delete(id);
    clearI(id);
  };

  // Listeners on `window` and `document`, which outlive any editor that forgets them.
  const listeners = new Map<EventTarget, Map<string, Set<unknown>>>();
  const key = (type: string, options: unknown): string =>
    `${type}:${typeof options === 'boolean' ? options : !!(options as { capture?: boolean })?.capture}`;
  const add = EventTarget.prototype.addEventListener;
  const remove = EventTarget.prototype.removeEventListener;
  EventTarget.prototype.addEventListener = function (this: EventTarget, type, fn, options) {
    if ((this === window || this === document) && fn) {
      const byType = listeners.get(this) ?? new Map<string, Set<unknown>>();
      listeners.set(this, byType);
      const k = key(type, options);
      const set = byType.get(k) ?? new Set();
      byType.set(k, set);
      set.add(fn);
      const signal = (options as AddEventListenerOptions | undefined)?.signal;
      signal?.addEventListener('abort', () => set.delete(fn), { once: true });
    }
    return add.call(this, type, fn, options);
  };
  EventTarget.prototype.removeEventListener = function (this: EventTarget, type, fn, options) {
    listeners.get(this)?.get(key(type, options))?.delete(fn);
    return remove.call(this, type, fn, options);
  };
  const count = (target: EventTarget): number =>
    [...(listeners.get(target)?.values() ?? [])].reduce((sum, set) => sum + set.size, 0);

  w.__memProbe = () => ({
    blobs: blobs.size,
    fonts: document.fonts?.size ?? 0,
    observers,
    workers,
    bitmaps,
    intervals: intervals.size,
    windowListeners: count(window),
    documentListeners: count(document),
    images: document.getElementsByTagNameNS('http://www.w3.org/2000/svg', 'image').length,
  });
}

// ---------------------------------------------------------------------------------------------
// Readings.

interface Reading extends Probe {
  /** Resident memory of the browser this worker launched, all its processes: decoded images too. */
  browserMB: number;
  heapMB: number;
  buffersMB: number;
  nodes: number;
  listeners: number;
  documents: number;
}

/**
 * The resident memory of every process under this worker: the browser it launched, with its
 * renderer and GPU processes, where decoded pictures and raster tiles live.
 */
function browserMB(): number {
  const rows = execFileSync('ps', ['-axo', 'pid=,ppid=,rss='], { encoding: 'utf8' })
    .trim()
    .split('\n')
    .map((line) => line.trim().split(/\s+/).map(Number) as [number, number, number]);
  const mine = new Set([process.pid]);
  let grew = true;
  while (grew) {
    grew = false;
    for (const [pid, ppid] of rows)
      if (mine.has(ppid) && !mine.has(pid)) (mine.add(pid), (grew = true));
  }
  mine.delete(process.pid);
  return rows.reduce((kb, [pid, , rss]) => kb + (mine.has(pid) ? rss : 0), 0) / 1024;
}

async function reading(page: Page, cdp: CDPSession): Promise<Reading> {
  // Let pending timers and idle work run, then collect twice: finalizers free more on the second.
  await page.waitForTimeout(300);
  for (let i = 0; i < 3; i += 1) {
    await cdp.send('HeapProfiler.collectGarbage');
    await page.waitForTimeout(100);
  }
  const { metrics } = (await cdp.send('Performance.getMetrics')) as {
    metrics: { name: string; value: number }[];
  };
  const metric = (name: string): number => metrics.find((m) => m.name === name)?.value ?? 0;
  const usage = (await cdp.send('Runtime.getHeapUsage')) as { backingStorageSize?: number };
  const probe = await page.evaluate(() =>
    (window as unknown as { __memProbe: () => Probe }).__memProbe()
  );
  return {
    browserMB: browserMB(),
    heapMB: metric('JSHeapUsedSize') / 2 ** 20,
    buffersMB: (usage.backingStorageSize ?? 0) / 2 ** 20,
    nodes: metric('Nodes'),
    listeners: metric('JSEventListeners'),
    documents: metric('Documents'),
    ...probe,
  };
}

type Summary = Map<string, { count: number; size: number }>;

/**
 * A heap snapshot, summed by constructor name (detached DOM nodes keep their own name). The
 * snapshot of a large heap is longer than the longest string, so it stays bytes: the node list
 * is read number by number, and only the meta data and the string table become JSON.
 */
async function heapSummary(cdp: CDPSession): Promise<Summary> {
  const chunks: Buffer[] = [];
  const onChunk = (p: { chunk: string }): void => void chunks.push(Buffer.from(p.chunk));
  cdp.on('HeapProfiler.addHeapSnapshotChunk', onChunk);
  await cdp.send('HeapProfiler.collectGarbage');
  await cdp.send('HeapProfiler.takeHeapSnapshot', { reportProgress: false });
  cdp.off('HeapProfiler.addHeapSnapshotChunk', onChunk);
  const bytes = Buffer.concat(chunks);
  chunks.length = 0;
  const at = (marker: string): number => {
    const i = bytes.indexOf(marker);
    if (i < 0) throw new Error(`no ${marker} in the heap snapshot`);
    return i + marker.length;
  };
  const nodesAt = at('"nodes":[');
  const meta = JSON.parse(
    bytes.subarray(bytes.indexOf('{'), bytes.lastIndexOf(',', nodesAt)).toString() + '}'
  ) as { snapshot: { meta: { node_fields: string[]; node_types: [string[]] } } };
  const stringsAt = at('"strings":[');
  const strings = JSON.parse(
    `[${bytes.subarray(stringsAt, bytes.lastIndexOf(']')).toString()}]`
  ) as string[];
  const fields = meta.snapshot.meta.node_fields;
  const types = meta.snapshot.meta.node_types[0];
  const [iType, iName, iSize] = ['type', 'name', 'self_size'].map((f) => fields.indexOf(f));
  const out: Summary = new Map();
  const node: number[] = [];
  let value = -1;
  for (let i = nodesAt; i < bytes.length; i += 1) {
    const c = bytes[i]!;
    if (c >= 48 && c <= 57) {
      value = (value < 0 ? 0 : value * 10) + (c - 48);
      continue;
    }
    if (value >= 0) node.push(value);
    value = -1;
    if (node.length === fields.length) {
      const type = types[node[iType!]!]!;
      if (type !== 'hidden' && type !== 'synthetic') {
        const name =
          type === 'object' || type === 'native' || type === 'closure'
            ? strings[node[iName!]!]!
            : `(${type})`;
        const entry = out.get(name) ?? { count: 0, size: 0 };
        entry.count += 1;
        entry.size += node[iSize!]!;
        out.set(name, entry);
      }
      node.length = 0;
    }
    if (c === 93) break; // The `]` that ends the node list.
  }
  return out;
}

function summaryDiff(before: Summary, after: Summary, top = 15): string[] {
  const rows = [...after].map(([name, a]) => {
    const b = before.get(name) ?? { count: 0, size: 0 };
    return { name, count: a.count - b.count, size: a.size - b.size };
  });
  const detached = rows.filter((r) => r.name.startsWith('Detached ') && r.count > 0);
  // The browser's own layout objects (`blink::`, `cppgc::`) count apart from the page's objects.
  const native = (r: { name: string }): boolean => /^(blink|cppgc|v8)::/.test(r.name);
  const blink = rows.filter(native).reduce((kb, r) => kb + r.size / 1024, 0);
  const grown = rows
    .filter((r) => !native(r))
    .sort((x, y) => y.size - x.size)
    .slice(0, top);
  return [
    `  ${blink.toFixed(0).padStart(8)} KB          (browser objects: layout, style, text shaping)`,
    ...grown.map(
      (r) =>
        `  ${(r.size / 1024).toFixed(0).padStart(8)} KB ${String(r.count).padStart(7)}  ${r.name.slice(0, 80)}`
    ),
    ...(detached.length
      ? [`  detached: ${detached.map((r) => `${r.name} +${r.count}`).join(', ')}`]
      : []),
  ];
}

// ---------------------------------------------------------------------------------------------
// Verdicts.

interface Budget {
  /** The heap growth per iteration that counts as a leak, in KB. */
  heapKB: number;
  /** DOM nodes per iteration. */
  nodes?: number;
  /** Listeners per iteration. */
  listeners?: number;
}

/** Least-squares slope of `ys` per step. */
function slope(ys: number[]): number {
  const k = ys.length;
  if (k < 2) return 0;
  const mx = (k - 1) / 2;
  const my = ys.reduce((a, b) => a + b, 0) / k;
  let num = 0;
  let den = 0;
  ys.forEach((y, x) => {
    num += (x - mx) * (y - my);
    den += (x - mx) ** 2;
  });
  return num / den;
}

interface Result {
  scenario: string;
  unit: string;
  iterations: number;
  base: Reading;
  end: Reading;
  heapKBPerIteration: number;
  nodesPerIteration: number;
  listenersPerIteration: number;
  leaks: string[];
  extra?: Record<string, unknown>;
}

const HEADER =
  'scenario         | unit    | iter | heap MB base -> end (KB/iter) | nodes base -> end (/iter) | ' +
  'listeners base -> end | blobs | fonts | observers | verdict';

function row(r: Result): string {
  const f = (x: number, d = 1): string => x.toFixed(d);
  return (
    `${r.scenario.padEnd(16)} | ${r.unit.padEnd(7)} | ${String(r.iterations).padStart(4)} | ` +
    `${f(r.base.heapMB)} -> ${f(r.end.heapMB)} (${f(r.heapKBPerIteration, 0)}) | ` +
    `${r.base.nodes} -> ${r.end.nodes} (${f(r.nodesPerIteration)}) | ` +
    `${r.base.listeners} -> ${r.end.listeners} | ${r.base.blobs} -> ${r.end.blobs} | ` +
    `${r.base.fonts} -> ${r.end.fonts} | ${r.base.observers} -> ${r.end.observers} | ` +
    `${r.leaks.length ? `LEAK: ${r.leaks.join('; ')}` : 'ok'}`
  );
}

/** Growth checks over the second half of the readings (`perStep` iterations per reading). */
function verdict(readings: Reading[], perStep: number, budget: Budget): string[] {
  const late = readings.slice(Math.floor(readings.length / 2));
  const per = (f: (r: Reading) => number): number => slope(late.map(f)) / perStep;
  const leaks: string[] = [];
  const heap = per((r) => r.heapMB) * 1024;
  if (heap > budget.heapKB) leaks.push(`heap +${heap.toFixed(0)} KB/iteration`);
  const nodes = per((r) => r.nodes);
  if (nodes > (budget.nodes ?? 5)) leaks.push(`DOM nodes +${nodes.toFixed(1)}/iteration`);
  const listeners = per((r) => r.listeners);
  if (listeners > (budget.listeners ?? 1))
    leaks.push(`listeners +${listeners.toFixed(1)}/iteration`);
  // Resources the editor must release: any steady growth is a leak.
  const base = readings[0]!;
  const end = readings.at(-1)!;
  const iterations = (readings.length - 1) * perStep;
  for (const k of [
    'blobs',
    'fonts',
    'observers',
    'workers',
    'bitmaps',
    'intervals',
    'windowListeners',
    'documentListeners',
    'documents',
  ] as const) {
    const grew = end[k] - base[k];
    if (grew > 0 && grew >= Math.max(2, iterations / 4) && per((r) => r[k]) > 0)
      leaks.push(`${k} ${base[k]} -> ${end[k]}`);
  }
  return leaks;
}

/**
 * Run `iterations` iterations of `step` after `warmup` ones, a reading every `perStep`, and
 * record the result. The baseline is the reading after the warm-up.
 */
async function scenario(
  page: Page,
  cdp: CDPSession,
  o: {
    name: string;
    unit: string;
    warmup: number;
    iterations: number;
    perStep?: number;
    budget: Budget;
    step: (i: number) => Promise<void>;
    extra?: () => Promise<Record<string, unknown>>;
  }
): Promise<Result> {
  const perStep = o.perStep ?? 1;
  for (let i = 0; i < o.warmup; i += 1) await o.step(-1 - i);
  const readings = [await reading(page, cdp)];
  const before = SNAPSHOTS ? await heapSummary(cdp) : null;
  for (let i = 0; i < o.iterations; i += 1) {
    await o.step(i);
    if ((i + 1) % perStep === 0) {
      readings.push(await reading(page, cdp));
      const r = readings.at(-1)!;
      console.log(
        `  ${o.name} ${i + 1}/${o.iterations}: browser ${r.browserMB.toFixed(0)} MB, heap ` +
          `${r.heapMB.toFixed(1)} MB, buffers ` +
          `${r.buffersMB.toFixed(1)} MB, nodes ${r.nodes}, listeners ${r.listeners}, blobs ` +
          `${r.blobs}, fonts ${r.fonts}, observers ${r.observers}, images ${r.images}`
      );
    }
  }
  const leaks = verdict(readings, perStep, o.budget);
  const base = readings[0]!;
  const end = readings.at(-1)!;
  const late = readings.slice(Math.floor(readings.length / 2));
  const result: Result = {
    scenario: o.name,
    unit: o.unit,
    iterations: o.iterations,
    base,
    end,
    heapKBPerIteration: (slope(late.map((r) => r.heapMB)) * 1024) / perStep,
    nodesPerIteration: slope(late.map((r) => r.nodes)) / perStep,
    listenersPerIteration: slope(late.map((r) => r.listeners)) / perStep,
    leaks,
    ...(o.extra ? { extra: await o.extra() } : {}),
  };
  if (before || leaks.length) {
    const after = await heapSummary(cdp);
    const lines = before
      ? summaryDiff(before, after)
      : ['  (no baseline snapshot: run with MEMORY_SNAPSHOTS=1)'];
    console.log(`  ${o.name}: heap growth by constructor, baseline to end\n${lines.join('\n')}`);
  }
  for (const [text, k] of CONSOLE.get(page) ?? []) console.log(`  console x${k} ${text}`);
  console.log(`${HEADER}\n${row(result)}`);
  mkdirSync(OUT, { recursive: true });
  writeFileSync(`${OUT}/${o.name}.json`, JSON.stringify({ result, readings }, null, 2));
  return result;
}

// ---------------------------------------------------------------------------------------------
// Driving the demo.

type Ed = {
  session: unknown;
  snapshot(): {
    status: string;
    slides: unknown[];
    pendingSlides?: number;
    isOpening?: boolean;
    currentSlide: number;
  };
  goToSlide(i: number): void;
  setSelection(s: unknown): void;
  focus(): void;
};
type Win = { __pptxEditor: Ed; __prevSession?: WeakRef<object> };

/** Console messages by text, so a warning that repeats per iteration shows in the log. */
const CONSOLE = new WeakMap<Page, Map<string, number>>();

function countConsole(page: Page): void {
  const seen = new Map<string, number>();
  CONSOLE.set(page, seen);
  page.on('console', (m) => {
    if (m.type() !== 'error' && m.type() !== 'warning') return;
    const text = `${m.type()}: ${m.text().slice(0, 160)}`;
    seen.set(text, (seen.get(text) ?? 0) + 1);
  });
}

async function start(page: Page): Promise<CDPSession> {
  countConsole(page);
  await page.addInitScript(installProbe);
  await page.goto('/');
  await page.waitForSelector('.pptx-surface__slide', { timeout: 60_000 });
  await page.waitForFunction(
    () => (window as unknown as Win).__pptxEditor?.snapshot().status === 'ready'
  );
  const cdp = await page.context().newCDPSession(page);
  await cdp.send('Performance.enable');
  await cdp.send('HeapProfiler.enable');
  return cdp;
}

/** Wait until the editor shows a deck that is fully open, and the rail has every slide. */
async function settled(page: Page): Promise<void> {
  await page.waitForFunction(
    () => {
      const ed = (window as unknown as Win).__pptxEditor;
      const s = ed.snapshot();
      return (
        s.status === 'ready' &&
        !s.pendingSlides &&
        !s.isOpening &&
        document.querySelectorAll('.pptx-rail__item').length === s.slides.length
      );
    },
    undefined,
    { timeout: 300_000, polling: 100 }
  );
}

/** Open `bytes` through the demo's file input, as a person does. */
async function openDeck(page: Page, name: string, bytes: Uint8Array): Promise<void> {
  // A weak reference: the bench itself must not keep the old session alive.
  await page.evaluate(() => {
    const w = window as unknown as Win;
    w.__prevSession = new WeakRef(w.__pptxEditor.session as object);
  });
  await page.locator('input[type=file][accept*=".pptx"]').setInputFiles({
    name: `${name}.pptx`,
    mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    buffer: Buffer.from(bytes),
  });
  await page.waitForFunction(
    () => {
      const w = window as unknown as Win;
      return w.__pptxEditor.session !== w.__prevSession?.deref();
    },
    undefined,
    { timeout: 300_000 }
  );
  await settled(page);
  await page.waitForTimeout(500);
}

/** The node id of the first shape with text on the current slide, from the layout. */
async function firstTextShape(page: Page): Promise<{ nodeId: string; length: number }> {
  return page.evaluate(() => {
    type Item = { element: { nodeId: string }; text?: { paragraphLengths: number[] } };
    const ed = (window as unknown as Win).__pptxEditor as unknown as {
      session: { layout(i: number): { items: Item[] } };
      snapshot(): { currentSlide: number };
    };
    const items = ed.session.layout(ed.snapshot().currentSlide).items;
    const item = items.find((x) => (x.text?.paragraphLengths.length ?? 0) > 0)!;
    return { nodeId: item.element.nodeId, length: item.text!.paragraphLengths[0]! };
  });
}

// ---------------------------------------------------------------------------------------------
// Scenarios.

const expectNoLeak = (r: Result): void =>
  expect.soft(r.leaks, `${r.scenario} grows per ${r.unit}`).toEqual([]);

test('open-same: open the sample deck 20 times', async ({ page }) => {
  const cdp = await start(page);
  const deck = publicDeck('sample.pptx');
  const r = await scenario(page, cdp, {
    name: 'open-same',
    unit: 'open',
    warmup: 2,
    iterations: n(20),
    budget: { heapKB: 100 },
    step: () => openDeck(page, 'sample', deck),
  });
  expectNoLeak(r);
});

test('open-mixed: open four different decks in turn, 20 opens', async ({ page }) => {
  const cdp = await start(page);
  const decks: [string, Uint8Array][] = [
    ['text-300', textDeck(300)],
    ['pictures-30', pictureDeck(30)],
    ['showcase', publicDeck('showcase.pptx')],
    ['sample', publicDeck('sample.pptx')],
  ];
  const round = async (): Promise<void> => {
    for (const [name, bytes] of decks) await openDeck(page, name, bytes);
  };
  const r = await scenario(page, cdp, {
    name: 'open-mixed',
    unit: 'round',
    warmup: 1,
    iterations: n(5),
    budget: { heapKB: 400 },
    step: round,
  });
  expectNoLeak(r);
});

/** Mount and unmount `<PptxEditor>` in a box over the demo, through the demo's own modules. */
async function mountCycles(page: Page, cdp: CDPSession, adapter: 'react' | 'vue'): Promise<Result> {
  const deck = [...publicDeck('sample.pptx')];
  await page.evaluate(
    async ({ adapter, root }) => {
      // The demo's own module instances: the same React or Vue, and the same engine.
      const dep = (name: string): string => {
        const url = performance
          .getEntriesByType('resource')
          .map((e) => e.name)
          .find((u) => new URL(u).pathname.endsWith(`/.vite/deps/${name}`));
        if (!url) throw new Error(`no ${name} in the page`);
        return url;
      };
      const w = window as unknown as Record<string, unknown>;
      const pkg = `/@fs${root}/packages/${adapter}/src/index.ts`;
      const lib = (await import(/* @vite-ignore */ pkg)) as { PptxEditor: unknown };
      if (adapter === 'react') {
        const React = (await import(/* @vite-ignore */ dep('react.js'))) as {
          createElement: (...a: unknown[]) => unknown;
          default?: { createElement: (...a: unknown[]) => unknown };
        };
        const create = React.createElement ?? React.default!.createElement;
        const client = (await import(/* @vite-ignore */ dep('react-dom_client.js'))) as {
          createRoot?: (e: Element) => { render(n: unknown): void; unmount(): void };
          default?: { createRoot: (e: Element) => { render(n: unknown): void; unmount(): void } };
        };
        const createRoot = client.createRoot ?? client.default!.createRoot;
        w.__mount = (bytes: number[], ready: (ed: unknown) => void) => {
          const box = document.createElement('div');
          box.style.cssText = 'position:fixed;inset:40px;z-index:9999;background:#fff';
          document.body.append(box);
          const root = createRoot(box);
          root.render(create(lib.PptxEditor, { document: new Uint8Array(bytes), onReady: ready }));
          return () => (root.unmount(), box.remove());
        };
      } else {
        const Vue = (await import(/* @vite-ignore */ dep('vue.js'))) as {
          createApp: (...a: unknown[]) => { mount(e: Element): void; unmount(): void };
          h: (...a: unknown[]) => unknown;
        };
        w.__mount = (bytes: number[], ready: (ed: unknown) => void) => {
          const box = document.createElement('div');
          box.style.cssText = 'position:fixed;inset:40px;z-index:9999;background:#fff';
          document.body.append(box);
          const app = Vue.createApp({
            render: () =>
              Vue.h(lib.PptxEditor, { document: new Uint8Array(bytes), onReady: ready }),
          });
          app.mount(box);
          return () => (app.unmount(), box.remove());
        };
      }
    },
    { adapter, root: ROOT }
  );
  const cycle = async (): Promise<void> => {
    await page.evaluate(async (bytes) => {
      const w = window as unknown as Record<string, unknown>;
      const mount = w.__mount as (b: number[], r: (ed: unknown) => void) => () => void;
      let editor: Ed | null = null;
      const unmount = mount(bytes, (ed) => (editor = ed as Ed));
      const t0 = performance.now();
      await new Promise<void>((resolve, reject) => {
        const poll = (): void => {
          const ed = editor as Ed | null;
          if (ed?.snapshot().status === 'ready' && !ed.snapshot().pendingSlides) resolve();
          else if (performance.now() - t0 > 60_000) reject(new Error('the editor never opened'));
          else setTimeout(poll, 50);
        };
        poll();
      });
      await new Promise((r) => setTimeout(r, 500));
      unmount();
    }, deck);
  };
  return scenario(page, cdp, {
    name: `mount-${adapter}`,
    unit: 'mount',
    warmup: 2,
    iterations: n(20),
    budget: { heapKB: 100 },
    step: cycle,
  });
}

test('mount-react: mount and unmount the React editor 20 times', async ({ page }) => {
  const cdp = await start(page);
  expectNoLeak(await mountCycles(page, cdp, 'react'));
});

test('mount-vue: mount and unmount the Vue editor 20 times', async ({ page }) => {
  await page.addInitScript(installProbe);
  await page.goto(VUE_URL);
  await page.waitForSelector('.pptx-surface__slide', { timeout: 60_000 });
  const cdp = await page.context().newCDPSession(page);
  await cdp.send('Performance.enable');
  await cdp.send('HeapProfiler.enable');
  expectNoLeak(await mountCycles(page, cdp, 'vue'));
});

test('typing: 5000 keystrokes in one text box', async ({ page }) => {
  const cdp = await start(page);
  await openDeck(page, 'text-20', textDeck(20));
  const { nodeId, length } = await firstTextShape(page);
  await page.evaluate(
    ({ nodeId, length }) => {
      const ed = (window as unknown as Win).__pptxEditor;
      const at = { paragraph: 0, offset: length };
      ed.setSelection({ kind: 'text', nodeId, anchor: at, focus: at });
      ed.focus();
    },
    { nodeId, length }
  );
  const words = 'steady growth across every region ';
  const r = await scenario(page, cdp, {
    name: 'typing',
    unit: 'key',
    warmup: 200,
    iterations: n(5000),
    perStep: 500,
    // The text grows: a few bytes a key, and a painted line of a few nodes each 50 or so keys.
    // History keeps at most 200 entries, and the measured advances at most a fixed size.
    budget: { heapKB: 2, nodes: 0.2, listeners: 0.01 },
    step: async (i) => {
      const ch = words[(i + 200) % words.length]!;
      await page.keyboard.press(ch === ' ' ? 'Space' : ch);
      // A paragraph break now and then keeps lines short, as real typing does.
      if (i % 1000 === 999) await page.keyboard.press('Enter');
    },
  });
  expectNoLeak(r);
});

test('moves: 1000 shape nudges', async ({ page }) => {
  const cdp = await start(page);
  await openDeck(page, 'text-20', textDeck(20));
  const { nodeId } = await firstTextShape(page);
  await page.evaluate((id) => {
    const ed = (window as unknown as Win).__pptxEditor;
    ed.setSelection({ kind: 'shapes', nodeIds: [id] });
    ed.focus();
  }, nodeId);
  const r = await scenario(page, cdp, {
    name: 'moves',
    unit: 'move',
    warmup: 50,
    iterations: n(1000),
    perStep: 100,
    budget: { heapKB: 2, nodes: 0.05, listeners: 0.01 },
    step: (i) => page.keyboard.press(Math.floor(i / 10) % 2 ? 'ArrowLeft' : 'ArrowRight'),
  });
  expectNoLeak(r);
});

test('slides-3000: switch through a 3000-slide deck and scroll the rail', async ({ page }) => {
  const cdp = await start(page);
  const slides = n(3000);
  await openDeck(page, 'text-3000', textDeck(slides));
  const opened = await reading(page, cdp);
  const STRIDE = 10;
  const passes = Math.floor(slides / STRIDE);
  // A long deck parses and lays out a slide the first time it shows, and keeps it: the first
  // pass grows with the deck. The second pass shows the same slides again: it must not grow.
  const r = await scenario(page, cdp, {
    name: 'slides-3000',
    unit: 'switch',
    warmup: passes,
    iterations: passes,
    perStep: Math.max(1, Math.floor(passes / 10)),
    budget: { heapKB: 20, nodes: 2, listeners: 0.1 },
    step: async (i) => {
      const index = (i < 0 ? -1 - i : i) * STRIDE;
      await page.evaluate((index) => {
        const ed = (window as unknown as Win).__pptxEditor;
        ed.goToSlide(index);
        // Show the current slide's rail item, as a scroll to it does.
        document
          .querySelector(`.pptx-rail__item[aria-current='true']`)
          ?.scrollIntoView({ block: 'center' });
      }, index);
      await page.evaluate(() => new Promise((r) => requestAnimationFrame(() => setTimeout(r, 0))));
    },
    extra: async () => ({
      openedHeapMB: opened.heapMB,
      openedNodes: opened.nodes,
      openedBrowserMB: opened.browserMB,
      paintedThumbnails: await page.evaluate(
        () => document.querySelectorAll('.pptx-rail__thumb > svg').length
      ),
    }),
  });
  console.log(
    `  slides-3000: after open heap ${opened.heapMB.toFixed(1)} MB, nodes ${opened.nodes}, ` +
      `browser ${opened.browserMB.toFixed(0)} MB; painted thumbnails ${r.extra?.paintedThumbnails}`
  );
  // The painter keeps at most 200 thumbnails painted (`MAX_PAINTED`), and those in view.
  expect.soft(r.extra?.paintedThumbnails as number).toBeLessThanOrEqual(240);
  expectNoLeak(r);
});

test('pictures-200: visit every slide of a deck with 200 large pictures', async ({ page }) => {
  const cdp = await start(page);
  const slides = n(200);
  await openDeck(page, 'pictures-200', pictureDeck(slides));
  const r = await scenario(page, cdp, {
    name: 'pictures-200',
    unit: 'slide',
    warmup: 5,
    iterations: slides,
    perStep: Math.max(1, Math.floor(slides / 10)),
    // Each picture's blob URL (its file bytes) lives as long as the deck; decoded bitmaps are
    // the browser's, out of the JS heap.
    budget: { heapKB: 20, nodes: 2, listeners: 0.1 },
    step: async (i) => {
      const index = i < 0 ? -1 - i : i;
      await page.evaluate((index) => {
        (window as unknown as Win).__pptxEditor.goToSlide(index);
        document
          .querySelector(`.pptx-rail__item[aria-current='true']`)
          ?.scrollIntoView({ block: 'center' });
      }, index);
      await page.evaluate(() => new Promise((r) => requestAnimationFrame(() => setTimeout(r, 30))));
    },
  });
  // The blobs are one per picture: bounded by the deck, not by the visits.
  expect.soft(r.end.blobs).toBeLessThanOrEqual(slides + 5);
  expectNoLeak(r);
});

test('exports: PNG and ZIP exports, 10 times each', async ({ page }) => {
  const cdp = await start(page);
  const exportOnce = async (item: string): Promise<void> => {
    await page.getByRole('menuitem', { name: 'File', exact: true }).click();
    const [download] = await Promise.all([
      page.waitForEvent('download', { timeout: 120_000 }),
      page.getByRole('menuitem', { name: item }).click(),
    ]);
    await download.path();
    await download.delete();
  };
  const r = await scenario(page, cdp, {
    name: 'exports',
    unit: 'pair',
    warmup: 1,
    iterations: n(10),
    budget: { heapKB: 100 },
    step: async () => {
      await exportOnce('Current slide as PNG');
      await exportOnce('All slides as PNG (.zip)');
      // The demo revokes its download URL after a second.
      await page.waitForTimeout(1_200);
    },
  });
  expectNoLeak(r);
});

test('find: open, search, and close the find bar 30 times', async ({ page }) => {
  const cdp = await start(page);
  const stage = page.locator('.pptx-surface__stage');
  const bar = page.getByRole('search', { name: 'Find and replace' });
  const r = await scenario(page, cdp, {
    name: 'find',
    unit: 'session',
    warmup: 2,
    iterations: n(30),
    budget: { heapKB: 50 },
    step: async () => {
      await stage.click({ position: { x: 5, y: 5 } });
      await page.keyboard.press('ControlOrMeta+f');
      await bar.waitFor();
      await page.keyboard.type('the');
      await page.keyboard.press('Enter');
      await page.keyboard.press('Escape');
      await bar.waitFor({ state: 'hidden' });
    },
  });
  expectNoLeak(r);
});

test('slideshow: start and end the slide show 30 times', async ({ page }) => {
  const cdp = await start(page);
  const stage = page.locator('.pptx-surface__stage');
  const r = await scenario(page, cdp, {
    name: 'slideshow',
    unit: 'show',
    warmup: 2,
    iterations: n(30),
    budget: { heapKB: 50 },
    step: async () => {
      await stage.click({ position: { x: 5, y: 5 } });
      await page.keyboard.press('F5');
      await page.locator('.pptx-show').waitFor();
      await page.keyboard.press('ArrowRight');
      await page.keyboard.press('ArrowRight');
      await page.keyboard.press('Escape');
      await page.locator('.pptx-show').waitFor({ state: 'detached' });
    },
  });
  expectNoLeak(r);
});

test('dialogs: open and close Page setup 30 times', async ({ page }) => {
  const cdp = await start(page);
  const dialog = page.getByRole('dialog', { name: 'Page setup' });
  const r = await scenario(page, cdp, {
    name: 'dialogs',
    unit: 'dialog',
    warmup: 2,
    iterations: n(30),
    budget: { heapKB: 50 },
    step: async () => {
      await page.locator('.pptx-menubar__trigger').first().click();
      await page.getByRole('menuitem', { name: 'Page setup…' }).click();
      await dialog.waitFor();
      await page.keyboard.press('Escape');
      await dialog.waitFor({ state: 'hidden' });
    },
  });
  expectNoLeak(r);
});
