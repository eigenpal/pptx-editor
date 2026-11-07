// Opens each synthetic deck of `hang-decks.ts` in its own process: load, wait for every slide,
// lay out and paint every slide and thumbnail. A deck over the budget is a HANG; over the slow
// mark, SLOW. Each process runs the editor in happy-dom with the production measurer.
//
// Usage:
//   bun scripts/stress/hang-check.ts [--only id,id] [--budget 20000] [--slow 2000] [--json]
//   bun scripts/stress/hang-check.ts --write <dir>   (writes the decks as .pptx files)

import { mkdirSync, writeFileSync } from 'node:fs';
import { join } from 'node:path';
import { FONT_WAIT_MS } from '../../packages/core/src/editor/host-fonts.ts';
import { HANG_DECKS } from './hang-decks.ts';

const argv = process.argv.slice(2);
const value = (flag: string): string | undefined => {
  const i = argv.indexOf(flag);
  return i >= 0 ? argv[i + 1] : undefined;
};

interface Phases {
  build: number;
  load: number;
  ready: number;
  layout: number;
  paint: number;
  slides: number;
}

async function child(id: string): Promise<void> {
  const { GlobalRegistrator } = await import('@happy-dom/global-registrator');
  if (!GlobalRegistrator.isRegistered) GlobalRegistrator.register();
  const { createPptxEditor, defineFontResolver } =
    await import('../../packages/core/src/editor/index.ts');
  const deck = HANG_DECKS.find((d) => d.id === id)!;
  const t: Partial<Phases> = {};
  const mark = (k: keyof Phases, since: number): number => (
    (t[k] = Math.round(performance.now() - since)),
    performance.now()
  );
  const report = (extra: object = {}) =>
    process.stdout.write(`${JSON.stringify({ ...t, ...extra })}\n`);
  let at = performance.now();
  const bytes = deck.bytes();
  at = mark('build', at);
  const fonts = deck.stalledFonts ? defineFontResolver(() => new Promise(() => {})) : undefined;
  const editor = createPptxEditor(fonts ? { fonts } : {});
  const r = await editor.load(bytes);
  at = mark('load', at);
  if (!r.ok) {
    report({ refused: r.detail ?? r.reason });
    return;
  }
  while ((editor.snapshot().pendingSlides ?? 0) > 0) await new Promise((res) => setTimeout(res, 1));
  at = mark('ready', at);
  const n = editor.snapshot().slides.length;
  t.slides = n;
  for (let i = 0; i < n; i += 1) editor.session!.layout(i);
  at = mark('layout', at);
  const host = document.createElement('div');
  document.body.append(host);
  editor.attach(host);
  for (let i = 0; i < Math.min(n, 50); i += 1) {
    const box = document.createElement('div');
    editor.renderThumbnail(i, box);
  }
  mark('paint', at);
  report({ status: editor.snapshot().status });
  editor.destroy();
}

async function main(): Promise<void> {
  if (argv[0] === '--child') return child(argv[1]!);
  const out = value('--write');
  if (out) {
    mkdirSync(out, { recursive: true });
    for (const d of HANG_DECKS) writeFileSync(join(out, `${d.id}.pptx`), d.bytes());
    console.log(`wrote ${HANG_DECKS.length} decks to ${out}`);
    return;
  }
  const budget = Number(value('--budget') ?? 20_000);
  const slow = Number(value('--slow') ?? 2000);
  const only = value('--only')?.split(',');
  const decks = HANG_DECKS.filter((d) => !only || only.includes(d.id));
  const rows: object[] = [];
  let bad = 0;
  for (const d of decks) {
    const start = performance.now();
    const proc = Bun.spawn(['bun', import.meta.path, '--child', d.id], {
      stdout: 'pipe',
      stderr: 'pipe',
    });
    const timer = setTimeout(() => proc.kill(9), budget);
    const [stdout, stderr, code] = await Promise.all([
      new Response(proc.stdout).text(),
      new Response(proc.stderr).text(),
      proc.exited,
    ]);
    clearTimeout(timer);
    const ms = Math.round(performance.now() - start);
    const line = stdout.trim().split('\n').pop() ?? '';
    let phases: Record<string, unknown> = {};
    try {
      phases = JSON.parse(line);
    } catch {
      /* no report: hang or crash */
    }
    // A stalled font origin costs the font wait by design; only time past it is slow.
    const mark = (d.slowMs ?? slow) + (d.stalledFonts ? FONT_WAIT_MS : 0);
    const verdict = ms >= budget ? 'HANG' : code !== 0 ? 'CRASH' : ms >= mark ? 'SLOW' : 'ok';
    if (verdict !== 'ok') bad += 1;
    const row = {
      id: d.id,
      verdict,
      ms,
      ...phases,
      ...(verdict === 'CRASH' ? { error: stderr.trim().split('\n').slice(-3).join(' | ') } : {}),
    };
    rows.push(row);
    if (!argv.includes('--json'))
      console.log(
        `${verdict.padEnd(5)} ${String(ms).padStart(6)} ms  ${d.id.padEnd(26)} ${d.target}${phases.refused ? ` (refused: ${phases.refused})` : ''}  ${JSON.stringify(phases)}${'error' in row ? `  ${row.error}` : ''}`
      );
  }
  if (argv.includes('--json')) console.log(JSON.stringify(rows, null, 2));
  console.log(
    `\n${decks.length - bad}/${decks.length} ok, budget ${budget} ms, slow mark ${slow} ms`
  );
  process.exitCode = bad ? 1 : 0;
}

await main();
