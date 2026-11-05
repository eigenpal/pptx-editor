// Every demo deck and synthetic fixture through the React demo in happy-dom, as
// `scripts/verify-demo.ts` does in Chromium: open it with the demo's Open button, visit every
// slide in the rail, export it with the Export menu, reopen the saved bytes, and fail on any
// console error. A browser also reports an SVG attribute that it cannot parse; here every
// painted attribute is checked for the values that cause it (`NaN`, `Infinity`, `undefined`).

import { afterEach, describe, expect, test } from 'bun:test';
import { readdirSync, readFileSync } from 'node:fs';
import { presentationInfo, readPptx } from '@pptx-editor.dev/core/store';
import {
  click,
  closeDemo,
  openDemo,
  openFile,
  repoFile,
  screen,
  until,
  type Demo,
} from './demo-harness.tsx';

const FIXTURES = new URL('../../../../fixtures/', import.meta.url);
const DECKS = [
  'examples/vite/public/sample.pptx',
  'examples/vite/public/showcase.pptx',
  ...readdirSync(FIXTURES)
    .filter((f) => f.endsWith('.pptx'))
    .sort()
    .map((f) => `fixtures/${f}`),
];

/** Painted attribute values that a browser's SVG parser rejects. */
const BAD_VALUE = /\b(NaN|Infinity|undefined)\b/;

function badAttributes(): string[] {
  const out: string[] = [];
  for (const el of document.querySelectorAll('.pptx-surface__slide, .pptx-surface__slide *'))
    for (const a of el.attributes)
      if (a.name !== 'data-node-id' && BAD_VALUE.test(a.value))
        out.push(`<${el.localName} ${a.name}="${a.value.slice(0, 80)}">`);
  return out;
}

afterEach(closeDemo);

/** Open one deck as `verify-demo.ts` does; returns what went wrong, or nothing. */
async function check(demo: Demo, path: string): Promise<string[]> {
  const problems: string[] = [];
  const before = demo.errors.length;
  const source = readPptx(
    new Uint8Array(readFileSync(new URL(`../../../../${path}`, import.meta.url)))
  );
  const expected = source.ok ? presentationInfo(source.pkg).slides.length : -1;
  await openFile(repoFile(path));
  const rail = () => document.querySelectorAll<HTMLElement>('.pptx-rail__item');
  if (rail().length !== expected) return [`${rail().length} slides in the rail, not ${expected}`];
  for (let i = 0; i < expected; i += 1) {
    await click(rail()[i]);
    if (demo.editor.snapshot().currentSlide !== i) problems.push(`slide ${i + 1} is not shown`);
    if (!document.querySelector('.pptx-surface__slide'))
      problems.push(`slide ${i + 1} is not painted`);
    problems.push(...badAttributes().map((b) => `slide ${i + 1}: ${b}`));
  }
  const saved = demo.downloads.length;
  await click(screen.getByRole('button', { name: 'Export' }));
  await click(screen.getByRole('menuitem', { name: /\.pptx/ }));
  await until(() => expect(demo.downloads.length).toBe(saved + 1));
  const reopened = readPptx(new Uint8Array(await demo.downloads[saved].blob.arrayBuffer()));
  const count = reopened.ok ? presentationInfo(reopened.pkg).slides.length : -1;
  if (count !== expected) problems.push(`the export reopens with ${count} slides`);
  problems.push(...demo.errors.slice(before).map((e) => `console error: ${e.slice(0, 200)}`));
  return problems;
}

describe('demo decks', () => {
  test('every demo deck and fixture opens, shows each slide, exports, and reopens', async () => {
    // One demo for all decks, as in the browser script: each Open replaces the deck.
    const demo = await openDemo();
    const report: Record<string, string[]> = {};
    for (const path of DECKS) {
      const problems = await check(demo, path);
      if (problems.length) report[path] = problems;
    }
    expect(report).toEqual({});
  }, 60_000);
});
