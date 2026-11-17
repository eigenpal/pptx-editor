// Document refresh benchmark: the engine work of one accepted result, without a browser, on
// synthetic decks of text slides (`refresh-deck.ts`).
//
// Each row is one kind of returned file: one changed paragraph, a changed paragraph on 10% of
// the slides, one inserted slide, and one slide on another layout. The columns are the steps of
// `applyUpdate()` in the order it runs them:
//
// - `read`: `readPptx` of the returned bytes (unzip, XML parse, limits).
// - `diff`: `diffPackages`, which includes its own verification: it applies the ops to a copy
//   and compares each edited part by canonical form.
// - `verify`: the share of `diff` that the verification takes (`applyOps` and the canonical
//   comparison, measured on their own).
// - `transact`: `PptxDocumentStore.transact` of the ops, as `replaceDocument` commits them
//   (`exact`, the part checks included).
// - `changes`: `computeRefreshChanges` from the captured package to the new one (Pro).
//
// Layout, paint, and the adapters are not in these numbers: `bun run bench:refresh:browser`
// measures them in Chromium. Each step must stay under the long-task budget of 50 ms on the
// 100-slide deck (openspec/changes/document-refresh): a median over it fails the run.
//
// Usage:
//   bun scripts/bench/refresh-bench.ts [--sizes 20,100,300] [--runs 15] [--warmup 3] [--json]

import {
  PptxDocumentStore,
  applyOps,
  canonicalForm,
  diffPackages,
  readPptx,
  type PptxPackage,
} from '../../packages/core/src/store/index.ts';
import { computeRefreshChanges } from '../../packages/pro/src/review/refresh-changes.ts';
import { refreshDeck, type RefreshEdits } from './refresh-deck.ts';

/** The long-task budget of one step (openspec/changes/document-refresh). */
const BUDGET_MS = 50;
/** The deck size the budget applies to. */
const BUDGET_SLIDES = 100;

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
    sizes: (value('--sizes') ?? '20,100,300').split(',').map(Number),
    runs: Number(value('--runs') ?? 15),
    warmup: Number(value('--warmup') ?? 3),
    json: argv.includes('--json'),
  };
}

interface Timing {
  medianMs: number;
  p95Ms: number;
}

const STEPS = ['read', 'diff', 'verify', 'transact', 'changes'] as const;
type Step = (typeof STEPS)[number];

interface Row {
  scenario: string;
  slides: number;
  ops: number;
  changes: number;
  steps: Record<Step, Timing>;
}

const median = (xs: number[]): number => {
  const s = [...xs].sort((a, b) => a - b);
  return s[Math.floor(s.length / 2)] ?? 0;
};
const p95 = (xs: number[]): number => {
  const s = [...xs].sort((a, b) => a - b);
  return s[Math.min(s.length - 1, Math.floor(s.length * 0.95))] ?? 0;
};

function scenarios(n: number): { name: string; edits: RefreshEdits }[] {
  const tenth = Array.from({ length: Math.max(1, Math.round(n / 10)) }, (_, i) =>
    Math.floor((i * n) / Math.max(1, Math.round(n / 10)))
  );
  return [
    { name: 'one-text', edits: { text: [Math.floor(n / 2)] } },
    { name: 'ten-percent', edits: { text: tenth } },
    { name: 'insert-slide', edits: { insertAfter: Math.floor(n / 2) } },
    { name: 'layout-change', edits: { layout: [Math.floor(n / 2)] } },
  ];
}

function read(bytes: Uint8Array): PptxPackage {
  const r = readPptx(bytes);
  if (!r.ok) throw new Error('the synthetic deck did not open');
  return r.pkg;
}

/** The verification step of `diffPackages`, on its own. */
function verify(open: PptxPackage, next: PptxPackage, ops: Parameters<typeof applyOps>[1]) {
  const applied = applyOps(open, ops);
  if (!applied.ok) throw new Error('the difference does not apply');
  for (const op of ops) {
    if (op.op === 'addPart' || op.op === 'removePart') continue;
    const want = next.parts.get(op.part);
    const got = applied.pkg.parts.get(op.part);
    if (want?.kind === 'xml' && got?.kind === 'xml') canonicalForm(want.xml.root);
    if (got?.kind === 'xml') canonicalForm(got.xml.root);
  }
}

function benchScenario(n: number, name: string, edits: RefreshEdits, args: Args): Row {
  const openBytes = refreshDeck(n);
  const nextBytes = refreshDeck(n, edits);
  const samples: Record<Step, number[]> = {
    read: [],
    diff: [],
    verify: [],
    transact: [],
    changes: [],
  };
  let ops = 0;
  let changes = 0;
  for (let run = 0; run < args.warmup + args.runs; run++) {
    const open = read(openBytes);
    const store = new PptxDocumentStore(open);
    const times: Record<Step, number> = { read: 0, diff: 0, verify: 0, transact: 0, changes: 0 };
    let t = performance.now();
    const next = read(nextBytes);
    times.read = performance.now() - t;
    t = performance.now();
    const diff = diffPackages(store.package, next);
    times.diff = performance.now() - t;
    if (!diff.ok) throw new Error(`no difference as ops: ${diff.reason}`);
    t = performance.now();
    verify(store.package, next, diff.ops);
    times.verify = performance.now() - t;
    const captured = store.package;
    t = performance.now();
    const r = store.transact(diff.ops, { label: 'replaceDocument', exact: true });
    times.transact = performance.now() - t;
    if (!r.ok) throw new Error(`the transaction failed: ${JSON.stringify(r.rejection)}`);
    t = performance.now();
    const computed = computeRefreshChanges(captured, store.package);
    times.changes = performance.now() - t;
    ops = diff.ops.length;
    changes = computed.changes.length;
    if (run < args.warmup) continue;
    for (const step of STEPS) samples[step].push(times[step]);
  }
  const steps = {} as Record<Step, Timing>;
  for (const step of STEPS)
    steps[step] = { medianMs: median(samples[step]), p95Ms: p95(samples[step]) };
  return { scenario: name, slides: n, ops, changes, steps };
}

const fmt = (t: Timing): string => `${t.medianMs.toFixed(2)} / ${t.p95Ms.toFixed(2)}`;

function main(): void {
  const args = parseArgs(process.argv.slice(2));
  const rows: Row[] = [];
  for (const n of args.sizes)
    for (const s of scenarios(n)) rows.push(benchScenario(n, s.name, s.edits, args));
  const over = rows.filter(
    (r) => r.slides <= BUDGET_SLIDES && STEPS.some((step) => r.steps[step].medianMs > BUDGET_MS)
  );
  for (const r of over)
    console.error(
      `${r.scenario} at ${r.slides} slides: a step takes over ${BUDGET_MS} ms ` +
        STEPS.map((s) => `${s} ${r.steps[s].medianMs.toFixed(1)}`).join(', ')
    );
  if (over.length) process.exitCode = 1;
  if (args.json) {
    console.log(JSON.stringify({ schema: 1, runtime: `bun ${Bun.version}`, args, rows }));
    return;
  }
  console.log(
    '| Scenario | Slides | Ops | Changes | read ms | diff ms | verify ms | transact ms | changes ms |'
  );
  console.log('| --- | --: | --: | --: | --: | --: | --: | --: | --: |');
  for (const r of rows)
    console.log(
      `| ${r.scenario} | ${r.slides} | ${r.ops} | ${r.changes} | ` +
        STEPS.map((s) => fmt(r.steps[s])).join(' | ') +
        ' |'
    );
  console.log('\nEach cell: median / p95 over the runs.');
}

if (import.meta.main) main();
