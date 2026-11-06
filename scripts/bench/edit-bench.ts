// Editing benchmark: opening a deck, one keystroke, one drag step, and one slide switch on
// synthetic decks of text slides, run headless through the editor facade.
//
// The editor runs in happy-dom with the production HarfBuzz measurer and the packaged
// substitute fonts. The bench stands in for the React chrome the way the adapter drives it:
// after every change it reads the snapshot, asks `toolbarCommandState` for every toolbar
// slot (as `useEditorCommand` does), and calls `renderThumbnail` for each slide summary whose
// object changed (as the rail's memoized `Thumbnail` does). Timings include that work.
//
// Timings depend on the machine; the counters do not. Read the counters first: layout runs,
// `renderThumbnail` calls, thumbnail paints, slide repaints, and the `can` probes compiled per
// operation. A probe compiles a command on a copy of the deck; New Slide and Add Section probes
// cost more as the deck grows, so typing and dragging compile none (`Deck probes` stays 0). The
// `open` row is one load, from the bytes to a rail with every container asked to paint: its
// thumbnail paints are the ones the editor made before it could be used, and its idle paints
// the rest.
// The `open-ready` row is the time until no slide is still opening (`pendingSlides` is 0): a
// long deck shows its first slide first and parses the rest in the background.
//
// Model: DOCX Editor's `scripts/bench/edit-bench.ts` (EigenPal, same owner).
//
// Usage:
//   bun scripts/bench/edit-bench.ts [--sizes 10,60,200,800,5000] [--runs 40] [--warmup 5] [--json]

import { GlobalRegistrator } from '@happy-dom/global-registrator';
import { textDeck } from './text-deck.ts';

if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();

// The engine reads `document` when it registers fonts, so it loads after the DOM exists.
const { chromeSlotIds, createPptxEditor, toolbarCommandState } =
  await import('../../packages/core/src/editor/index.ts');
type Editor = Awaited<ReturnType<typeof createPptxEditor>>;
type Snapshot = ReturnType<Editor['snapshot']>;
type SlideLayout = NonNullable<ReturnType<NonNullable<Editor['session']>['layout']>>;

interface Args {
  sizes: number[];
  runs: number;
  warmup: number;
  json: boolean;
}

function parseArgs(argv: readonly string[]): Args {
  const value = (flag: string): string | undefined => {
    const i = argv.indexOf(flag);
    return i >= 0 ? argv[i + 1] : undefined;
  };
  return {
    sizes: (value('--sizes') ?? '10,60,200,800,5000').split(',').map(Number),
    runs: Number(value('--runs') ?? 40),
    warmup: Number(value('--warmup') ?? 5),
    json: argv.includes('--json'),
  };
}

interface Timing {
  medianMs: number;
  p95Ms: number;
}

interface Counters {
  /** `session.layout` calls per operation. */
  layoutCalls: number;
  /** Slide layouts computed (a `session.layout` result not seen before). */
  layoutsComputed: number;
  /** `renderThumbnail` calls the rail made. */
  thumbnailCalls: number;
  /** Thumbnails repainted during the operation. */
  thumbnailPaints: number;
  /** Thumbnails repainted after the operation, when the editor was idle. */
  idleThumbnailPaints: number;
  /** `can` answers compiled (a probe batch opened), not answered from the cache. */
  probes: number;
  /** Of those, New Slide and Add Section: probes that run a whole-deck command. */
  deckProbes: number;
  /** Times the editing surface painted the slide anew. */
  slidePaints: number;
}

/** Commands whose probe runs a whole-deck command. */
const DECK_COMMANDS: ReadonlySet<string> = new Set(['addSlide', 'addSection']);

interface ScenarioResult {
  name: string;
  slides: number;
  total: Timing;
  toolbar: Timing;
  rail: Timing;
  perOperation: Counters;
}

const median = (xs: number[]): number => {
  const s = [...xs].sort((a, b) => a - b);
  return s[Math.floor(s.length / 2)] ?? 0;
};
const p95 = (xs: number[]): number => {
  const s = [...xs].sort((a, b) => a - b);
  return s[Math.min(s.length - 1, Math.floor(s.length * 0.95))] ?? 0;
};
const timing = (xs: number[]): Timing => ({ medianMs: median(xs), p95Ms: p95(xs) });

/** The chrome the React adapter would run, driven synchronously after each change. */
class Harness {
  readonly editor: Editor;
  readonly host: HTMLDivElement;
  readonly thumbs: HTMLDivElement[] = [];
  #last: Snapshot['slides'] = [];
  #painted = new Map<HTMLElement, Element | null>();
  #slide: Element | null = null;
  #seen = new WeakSet<SlideLayout>();
  counters: Counters = this.#zero();
  toolbarMs = 0;
  railMs = 0;

  constructor(editor: Editor) {
    this.editor = editor;
    this.host = document.createElement('div');
    document.body.append(this.host);
    editor.subscribe(() => this.#onChange());
  }

  #zero(): Counters {
    return {
      layoutCalls: 0,
      layoutsComputed: 0,
      thumbnailCalls: 0,
      thumbnailPaints: 0,
      idleThumbnailPaints: 0,
      probes: 0,
      deckProbes: 0,
      slidePaints: 0,
    };
  }

  /** Count layout calls: wrap the session's `layout` once per session. */
  instrument(): void {
    const session = this.editor.session;
    if (!session) throw new Error('no session');
    const layout = session.layout.bind(session);
    session.layout = (index: number) => {
      this.counters.layoutCalls += 1;
      const r = layout(index);
      if (r && !this.#seen.has(r)) {
        this.#seen.add(r);
        this.counters.layoutsComputed += 1;
      }
      return r;
    };
    // A compiled `can` answer opens one probe batch; a cached one opens none.
    let batches = 0;
    const probeBatch = session.probeBatch.bind(session);
    session.probeBatch = (base) => ((batches += 1), probeBatch(base));
    const can = this.editor.can.bind(this.editor);
    this.editor.can = (command, options) => {
      const before = batches;
      const r = can(command, options);
      if (batches > before) {
        this.counters.probes += 1;
        if (DECK_COMMANDS.has(command.type)) this.counters.deckProbes += 1;
      }
      return r;
    };
  }

  reset(): void {
    this.counters = this.#zero();
    this.toolbarMs = 0;
    this.railMs = 0;
  }

  #onChange(): void {
    const snap = this.editor.snapshot();
    let t0 = performance.now();
    for (const slot of chromeSlotIds()) toolbarCommandState(this.editor, slot);
    this.toolbarMs += performance.now() - t0;
    t0 = performance.now();
    this.syncRail(snap.slides);
    this.railMs += performance.now() - t0;
    this.countSlidePaints();
  }

  /** Whether the surface painted a new slide SVG since the last count. */
  countSlidePaints(): void {
    const svg = this.host.querySelector('.pptx-surface__slide');
    if (svg !== this.#slide) {
      if (this.#slide) this.counters.slidePaints += 1;
      this.#slide = svg;
    }
  }

  /** The rail: one container per slide; repaint a thumbnail when its summary changed. */
  syncRail(slides: Snapshot['slides']): void {
    while (this.thumbs.length < slides.length) {
      const box = document.createElement('div');
      this.host.append(box);
      this.thumbs.push(box);
    }
    slides.forEach((s, i) => {
      if (this.#last[i] === s) return;
      this.counters.thumbnailCalls += 1;
      this.editor.renderThumbnail(i, this.thumbs[i], { defer: true });
    });
    this.#last = slides;
    this.countPaints('thumbnailPaints');
  }

  /** Thumbnails whose painted SVG changed since the last count (a first paint counts too). */
  countPaints(field: 'thumbnailPaints' | 'idleThumbnailPaints'): void {
    for (const box of this.thumbs) {
      const svg = box.firstElementChild;
      if ((this.#painted.get(box) ?? null) !== svg) {
        if (svg) this.counters[field] += 1;
        this.#painted.set(box, svg);
      }
    }
  }
}

const idle = (ms: number): Promise<void> => new Promise((r) => setTimeout(r, ms));

/** The body text box and the shrink-on-overflow box of slide `index`. */
function shapesOf(editor: Editor, index: number) {
  const lay = editor.session!.layout(index)!;
  const named = (name: string) =>
    lay.items.find((i) => i.kind === 'shape' && i.element.name === name);
  return { body: named('Body')!, fit: named('Fit')!, lay };
}

async function measure(
  h: Harness,
  name: string,
  slides: number,
  args: Args,
  step: (i: number) => void
): Promise<ScenarioResult> {
  for (let i = 0; i < args.warmup; i += 1) step(i);
  await idle(400);
  h.countPaints('idleThumbnailPaints');
  h.reset();
  Bun.gc(true);
  const total: number[] = [];
  const toolbar: number[] = [];
  const rail: number[] = [];
  for (let i = 0; i < args.runs; i += 1) {
    const tb = h.toolbarMs;
    const rl = h.railMs;
    const t0 = performance.now();
    step(args.warmup + i);
    total.push(performance.now() - t0);
    toolbar.push(h.toolbarMs - tb);
    rail.push(h.railMs - rl);
  }
  // Repaints the editor defers until it is idle.
  await idle(400);
  h.countPaints('idleThumbnailPaints');
  return result(h, name, slides, args.runs, total, toolbar, rail);
}

function result(
  h: Harness,
  name: string,
  slides: number,
  runs: number,
  total: number[],
  toolbar: number[],
  rail: number[]
): ScenarioResult {
  const per = (v: number): number => Math.round((v / runs) * 100) / 100;
  const c = h.counters;
  return {
    name,
    slides,
    total: timing(total),
    toolbar: timing(toolbar),
    rail: timing(rail),
    perOperation: {
      layoutCalls: per(c.layoutCalls),
      layoutsComputed: per(c.layoutsComputed),
      thumbnailCalls: per(c.thumbnailCalls),
      thumbnailPaints: per(c.thumbnailPaints),
      idleThumbnailPaints: per(c.idleThumbnailPaints),
      probes: per(c.probes),
      deckProbes: per(c.deckProbes),
      slidePaints: per(c.slidePaints),
    },
  };
}

/** Drag the selected shape by pointer events on the surface, as a person does. */
function dragDriver(h: Harness, index: number) {
  const editor = h.editor;
  const stage = h.host.querySelector<HTMLElement>('.pptx-surface__stage');
  if (!stage) throw new Error('no surface');
  const snap = editor.snapshot();
  const W = 1280;
  const H = (W * snap.slideHeight) / snap.slideWidth;
  stage.getBoundingClientRect = () => new DOMRect(0, 0, W, H);
  const client = (x: number, y: number) => ({
    clientX: (x / snap.slideWidth) * W,
    clientY: (y / snap.slideWidth) * W,
  });
  const { body } = shapesOf(editor, index);
  editor.setSelection({ kind: 'shapes', nodeIds: [body.element.nodeId] });
  const [, , , , tx, ty] = body.matrix;
  // The top border band, inside the box: a press there moves the shape.
  const start = client(tx + body.element.xfrm.cx / 2, ty + 2_000);
  const target = stage.querySelector('.pptx-surface__slide') ?? stage;
  target.dispatchEvent(new PointerEvent('pointerdown', { bubbles: true, button: 0, ...start }));
  return (i: number): void => {
    const dx = ((i % 40) + 1) * 3;
    document.dispatchEvent(
      new PointerEvent('pointermove', {
        bubbles: true,
        clientX: start.clientX + dx,
        clientY: start.clientY + dx / 2,
      })
    );
  };
}

async function benchDeck(n: number, args: Args): Promise<ScenarioResult[]> {
  const bytes = textDeck(n);
  Bun.gc(true);
  // (o) Open: load, attach, and mount the rail, until the editor can be used.
  const t0 = performance.now();
  const editor = createPptxEditor({});
  const r = await editor.load(bytes);
  if (!r.ok) throw new Error(`load failed: ${r.reason}`);
  const h = new Harness(editor);
  editor.attach(h.host);
  h.syncRail(editor.snapshot().slides);
  const opened = performance.now() - t0;
  // (o2) Ready: a long deck opens progressively; the rest of its slides parse in the background.
  while ((editor.snapshot().pendingSlides ?? 0) > 0) await idle(1);
  const ready = performance.now() - t0;
  h.counters.thumbnailCalls = n;
  // Thumbnails painted later, when the editor is idle (every one, at the latest after this).
  for (let wait = 0; wait < 600 && h.thumbs.some((b) => !b.firstElementChild); wait += 1)
    await idle(5);
  h.countPaints('idleThumbnailPaints');
  if (h.thumbs.some((b) => !b.firstElementChild))
    throw new Error('a thumbnail never painted after open');
  const out: ScenarioResult[] = [
    result(h, 'open', n, 1, [opened], [0], [0]),
    result(h, 'open-ready', n, 1, [ready], [0], [0]),
  ];
  h.reset();
  h.instrument();
  const middle = Math.floor(n / 2);

  // (a) One keystroke at the end of the body text on the middle slide.
  editor.goToSlide(middle);
  const { body } = shapesOf(editor, middle);
  const end = body.kind === 'shape' ? (body.text?.paragraphLengths[0] ?? 0) : 0;
  const caret = { paragraph: 0, offset: end };
  editor.setSelection({ kind: 'text', nodeId: body.element.nodeId, anchor: caret, focus: caret });
  out.push(await measure(h, 'keystroke', n, args, () => void editor.input('text', 'x')));

  // (a2) One keystroke in the overflowing shrink-on-overflow box (`a:normAutofit`).
  const { fit } = shapesOf(editor, middle);
  const fitEnd = fit.kind === 'shape' ? (fit.text?.paragraphLengths[0] ?? 0) : 0;
  const fitCaret = { paragraph: 0, offset: fitEnd };
  editor.setSelection({
    kind: 'text',
    nodeId: fit.element.nodeId,
    anchor: fitCaret,
    focus: fitCaret,
  });
  out.push(await measure(h, 'keystroke-normAutofit', n, args, () => editor.input('text', 'x')));

  // (b) One drag step (move) of the body text box.
  const from = shapesOf(editor, middle).body.element.xfrm;
  const move = dragDriver(h, middle);
  out.push(await measure(h, 'drag-step', n, args, move));
  document.dispatchEvent(new PointerEvent('pointerup', { bubbles: true }));
  if (!editor.snapshot().canUndo) throw new Error('the drag did not edit the deck');
  // The whole drag is one undo step.
  editor.exec({ type: 'undo' });
  const back = shapesOf(editor, middle).body.element.xfrm;
  if (back.x !== from.x || back.y !== from.y) throw new Error('the drag is not one undo step');
  editor.exec({ type: 'redo' });

  // (c) A slide switch between two neighbors.
  const other = Math.min(n - 1, middle + 1);
  out.push(
    await measure(h, 'slide-switch', n, args, (i) => editor.goToSlide(i % 2 ? other : middle))
  );
  editor.destroy();
  h.host.remove();
  return out;
}

const fmt = (t: Timing): string => `${t.medianMs.toFixed(2)} / ${t.p95Ms.toFixed(2)}`;

async function main(): Promise<void> {
  const args = parseArgs(process.argv.slice(2));
  const results: ScenarioResult[] = [];
  for (const n of args.sizes) results.push(...(await benchDeck(n, args)));
  // An edit on one slide keeps the deck's structure: no whole-deck probe may run again.
  const regressions = results.filter(
    (r) => r.name !== 'slide-switch' && r.perOperation.deckProbes > 0
  );
  for (const r of regressions)
    console.error(
      `${r.name} at ${r.slides} slides compiled ${r.perOperation.deckProbes} deck probes`
    );
  if (regressions.length) process.exitCode = 1;
  if (args.json) {
    console.log(JSON.stringify({ schema: 1, runtime: `bun ${Bun.version}`, args, results }));
    return;
  }
  console.log(
    '| Scenario | Slides | Total ms (median / p95) | Toolbar ms | Rail ms | Layout calls | Layouts computed | Thumbnail calls | Thumbnail paints | Idle paints | Slide paints | Probes | Deck probes |'
  );
  console.log('| --- | --: | --: | --: | --: | --: | --: | --: | --: | --: | --: | --: | --: |');
  for (const r of results) {
    const c = r.perOperation;
    console.log(
      `| ${r.name} | ${r.slides} | ${fmt(r.total)} | ${fmt(r.toolbar)} | ${fmt(r.rail)} | ${c.layoutCalls} | ${c.layoutsComputed} | ${c.thumbnailCalls} | ${c.thumbnailPaints} | ${c.idleThumbnailPaints} | ${c.slidePaints} | ${c.probes} | ${c.deckProbes} |`
    );
  }
}

if (import.meta.main) await main();
