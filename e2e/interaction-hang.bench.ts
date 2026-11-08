// Interaction hang bench: the longest main-thread task of each user action on an open deck, in
// the demo (React chrome, real rail), on the synthetic decks of
// `scripts/stress/interaction-decks.ts`. A browser that runs one task for a long time does not
// paint and does not answer input, so the longest task is what the user feels as a hang.
//
// Each action waits for the page to settle (no long task for 0.5 s) before and after it. The
// longest task comes from a `PerformanceObserver` for `longtask` entries (50 ms and more), so a
// row of 0 had no task of 50 ms or more. Actions run as the user runs them where the browser
// allows (keys, wheel, clipboard events, the browser's paste into the text field), and through
// `window.__pptxEditor` otherwise. A row over 200 ms is marked `slow`; over 1 s, `hang`.
//
// Usage: BENCH_PORT=5299 bun run bench:interaction:browser
// BENCH_SLIDES sets the long deck (default 3000). BENCH_STRICT=1 fails on a `hang` row.

import { expect, test, type Page } from '@playwright/test';
import { bigTable, bigText, manyShapes, manySlides } from '../scripts/stress/interaction-decks.ts';

const SLIDES = Number(process.env.BENCH_SLIDES ?? 3000);
const SLOW_MS = 200;
const HANG_MS = 1_000;
const MOD = process.platform === 'darwin' ? 'Meta' : 'Control';

interface Row {
  readonly deck: string;
  readonly action: string;
  readonly longest: number;
  readonly total: number;
  readonly wall: number;
}

const rows: Row[] = [];

/** Open a deck through the file input; resolves once the editor is ready and idle. */
async function open(page: Page, bytes: Uint8Array): Promise<void> {
  await page.goto('/');
  await page.waitForSelector('.pptx-surface__slide', { timeout: 60_000 });
  await page.locator('input[type=file][accept*=".pptx"]').setInputFiles({
    name: 'bench.pptx',
    mimeType: 'application/vnd.openxmlformats-officedocument.presentationml.presentation',
    buffer: Buffer.from(bytes),
  });
  await page.waitForFunction(
    () => {
      const ed = (window as unknown as { __pptxEditor?: { snapshot(): { status: string } } })
        .__pptxEditor;
      return ed?.snapshot().status === 'ready';
    },
    undefined,
    { timeout: 300_000 }
  );
  await settle(page, 1_000);
}

/** Wait until no long task started for `quiet` ms, at most 60 s. */
async function settle(page: Page, quiet = 500): Promise<void> {
  const t0 = Date.now();
  let seen = -1;
  let since = Date.now();
  while (Date.now() - t0 < 60_000) {
    await page.waitForTimeout(100);
    const n = await page.evaluate(
      () => (window as unknown as { __tasks: number[] }).__tasks.length
    );
    if (n !== seen) {
      seen = n;
      since = Date.now();
    } else if (Date.now() - since >= quiet) return;
  }
}

/** Run one action and record its longest task. */
async function measure(
  page: Page,
  deck: string,
  action: string,
  run: () => Promise<unknown>
): Promise<void> {
  await settle(page);
  await page.evaluate(() => ((window as unknown as { __tasks: number[] }).__tasks = []));
  const t0 = Date.now();
  await run();
  const wall = Date.now() - t0;
  await settle(page);
  const tasks = await page.evaluate(() => (window as unknown as { __tasks: number[] }).__tasks);
  rows.push({
    deck,
    action,
    longest: Math.round(Math.max(0, ...tasks)),
    total: Math.round(tasks.reduce((a, b) => a + b, 0)),
    wall,
  });
}

/** The demo's editor, as the bench reads it. */
interface BenchEditor {
  exec(command: unknown): unknown;
  goToSlide(index: number): void;
  selectSlide(index: number, mode?: 'range'): void;
  setSelection(selection: unknown): void;
  save(): unknown;
  snapshot(): { currentSlide: number; selectedSlides: number[] };
  find: { replaceAll(query: string, replacement: string): unknown };
  session: {
    layout(index: number): {
      items: {
        element: { name?: string; nodeId: string };
        text?: { paragraphLengths: number[] };
      }[];
    };
  };
}

/** Run an editor command. */
const exec = (page: Page, command: unknown) =>
  page.evaluate(
    (c) => void (window as unknown as { __pptxEditor: BenchEditor }).__pptxEditor.exec(c),
    command
  );

/** Select slides `from` to `to` in the slide list. */
const selectSlides = (page: Page, from: number, to: number) =>
  page.evaluate(
    ([a, b]) => {
      const editor = (window as unknown as { __pptxEditor: BenchEditor }).__pptxEditor;
      editor.selectSlide(a!);
      editor.selectSlide(b!, 'range');
    },
    [from, to]
  );

/** Put the caret at the end of the shape `name` of the current slide. */
const caretAtEnd = async (page: Page, name: string) => {
  await page.evaluate((shape) => {
    const editor = (window as unknown as { __pptxEditor: BenchEditor }).__pptxEditor;
    const items = editor.session.layout(editor.snapshot().currentSlide).items;
    const item = items.find((x) => x.element.name === shape) ?? items[0]!;
    const lengths = item.text?.paragraphLengths ?? [0];
    const at = { paragraph: lengths.length - 1, offset: lengths[lengths.length - 1] };
    editor.setSelection({ kind: 'text', nodeId: item.element.nodeId, anchor: at, focus: at });
  }, name);
  await page.locator('.pptx-surface__input').focus();
};

/** A paste event with `data` at the focused element, as the browser sends one. */
const pasteEvent = (page: Page, data: Record<string, string>) =>
  page.evaluate((d) => {
    const transfer = new DataTransfer();
    for (const [type, value] of Object.entries(d)) transfer.setData(type, value);
    const target = document.activeElement ?? document.body;
    target.dispatchEvent(
      new ClipboardEvent('paste', { clipboardData: transfer, bubbles: true, cancelable: true })
    );
  }, data);

/** The browser's own paste of plain text into the text field (`insertFromPaste`). */
const pasteIntoText = (page: Page, text: string) =>
  page.evaluate((t) => {
    document.querySelector('.pptx-surface__input')!.dispatchEvent(
      new InputEvent('beforeinput', {
        inputType: 'insertFromPaste',
        data: t,
        bubbles: true,
        cancelable: true,
      })
    );
  }, text);

const lines = (n: number) =>
  Array.from({ length: n }, (_, i) => `Line ${i} of the pasted text`).join('\n');

test('interaction hang bench', async ({ page }) => {
  await page.addInitScript(() => {
    const w = window as unknown as { __tasks: number[] };
    w.__tasks = [];
    new PerformanceObserver((list) => {
      for (const entry of list.getEntries()) w.__tasks.push(entry.duration);
    }).observe({ type: 'longtask', buffered: true });
  });

  // Paste and typing in a normal deck.
  await open(page, manySlides(20));
  let deck = '20 slides';
  await caretAtEnd(page, 'Body');
  await measure(page, deck, 'paste 3,000 lines into text', () => pasteIntoText(page, lines(3_000)));
  await exec(page, { type: 'undo' });
  await caretAtEnd(page, 'Body');
  await measure(page, deck, 'paste 30,000 lines into text', () =>
    pasteIntoText(page, lines(30_000))
  );
  await exec(page, { type: 'undo' });
  await page.evaluate(() =>
    (window as unknown as { __pptxEditor: BenchEditor }).__pptxEditor.setSelection({
      kind: 'none',
    })
  );
  await page.locator('.pptx-surface__input').focus();
  await measure(page, deck, 'paste 5,000 lines on the slide', () =>
    pasteEvent(page, { 'text/plain': lines(5_000) })
  );
  await exec(page, { type: 'undo' });
  await caretAtEnd(page, 'Body');
  const html = Array.from({ length: 5_000 }, (_, i) => `<p><b>p${i}</b> text</p>`).join('');
  await measure(page, deck, 'paste 5,000 HTML paragraphs', () =>
    pasteEvent(page, { 'text/html': html, 'text/plain': 'x' })
  );
  await exec(page, { type: 'undo' });
  await measure(page, deck, 'paste HTML nested 5,000 deep', () =>
    pasteEvent(page, { 'text/html': '<div>'.repeat(5_000) + 'x', 'text/plain': 'x' })
  );
  await caretAtEnd(page, 'Body');
  await measure(page, deck, 'type 30 characters', () =>
    page.keyboard.type('x'.repeat(30), { delay: 30 })
  );
  await measure(page, deck, 'Enter x20', async () => {
    for (let i = 0; i < 20; i += 1) await page.keyboard.press('Enter');
  });
  await measure(page, deck, 'undo x30', async () => {
    for (let i = 0; i < 30; i += 1) await page.keyboard.press(`${MOD}+z`);
  });

  // A text box of 100,000 characters.
  await open(page, bigText(100_000));
  deck = '100,000 characters';
  await caretAtEnd(page, 'Body');
  await measure(page, deck, 'type 20 characters', () =>
    page.keyboard.type('x'.repeat(20), { delay: 30 })
  );
  await measure(page, deck, 'select all text, bold', async () => {
    await page.keyboard.press(`${MOD}+a`);
    await page.keyboard.press(`${MOD}+b`);
  });
  await measure(page, deck, 'copy all text', () => page.keyboard.press(`${MOD}+c`));
  await measure(page, deck, 'type over all text', () => page.keyboard.type('x'));
  await measure(page, deck, 'undo', () => page.keyboard.press(`${MOD}+z`));

  // 2,500 shapes on one slide.
  await open(page, manyShapes(2_500));
  deck = '2,500 shapes';
  await page.locator('.pptx-surface__input').focus();
  await measure(page, deck, 'select all shapes', () => page.keyboard.press(`${MOD}+a`));
  await measure(page, deck, 'nudge x10', async () => {
    for (let i = 0; i < 10; i += 1) await page.keyboard.press('ArrowRight');
  });
  await measure(page, deck, 'copy all shapes', () => page.keyboard.press(`${MOD}+c`));
  await measure(page, deck, 'group, ungroup', async () => {
    await exec(page, { type: 'group' });
    await exec(page, { type: 'ungroup' });
  });
  await measure(page, deck, 'delete all shapes, undo', async () => {
    await page.keyboard.press('Delete');
    await page.keyboard.press(`${MOD}+z`);
  });

  // A table of 100 rows and 30 columns.
  await open(page, bigTable(100, 30));
  deck = '100x30 table';
  await measure(page, deck, 'double-click a cell, type', async () => {
    const box = (await page.locator('.pptx-surface__slide').boundingBox())!;
    await page.mouse.dblclick(box.x + 200, box.y + 200);
    await page.keyboard.type('hello', { delay: 30 });
  });
  await measure(page, deck, 'select all cells, bold', async () => {
    await page.keyboard.press(`${MOD}+a`);
    await page.keyboard.press(`${MOD}+b`);
  });

  // A long deck.
  await open(page, manySlides(SLIDES));
  deck = `${SLIDES} slides`;
  const last = SLIDES - 1;
  await measure(page, deck, 'go to the last slide', () =>
    page.evaluate(
      (i) => (window as unknown as { __pptxEditor: BenchEditor }).__pptxEditor.goToSlide(i),
      last
    )
  );
  await measure(page, deck, 'select every slide', () => selectSlides(page, 0, last));
  await measure(page, deck, 'hide every slide', () =>
    exec(page, { type: 'setSlidesHidden', hidden: true })
  );
  await measure(page, deck, 'duplicate every slide', () =>
    page.evaluate(() => {
      const editor = (window as unknown as { __pptxEditor: BenchEditor }).__pptxEditor;
      editor.exec({ type: 'duplicateSlides', indices: editor.snapshot().selectedSlides });
    })
  );
  await measure(page, deck, 'undo duplicate', () => exec(page, { type: 'undo' }));
  await measure(page, deck, 'delete all slides but one', async () => {
    await selectSlides(page, 1, last);
    await page.evaluate(() => {
      const editor = (window as unknown as { __pptxEditor: BenchEditor }).__pptxEditor;
      editor.exec({ type: 'deleteSlides', indices: editor.snapshot().selectedSlides });
    });
  });
  await measure(page, deck, 'undo delete', () => exec(page, { type: 'undo' }));
  await measure(page, deck, 'copy one slide in the list', async () => {
    await page.locator('.pptx-rail__item').nth(2).click();
    await page.keyboard.press(`${MOD}+c`);
  });
  await measure(page, deck, 'paste one slide in the list', () => page.keyboard.press(`${MOD}+v`));
  await measure(page, deck, 'add a slide', () => exec(page, { type: 'addSlide' }));
  await measure(page, deck, 'replace all (one word per slide)', () =>
    page.evaluate(
      () =>
        void (window as unknown as { __pptxEditor: BenchEditor }).__pptxEditor.find.replaceAll(
          'Quarterly',
          'Annual'
        )
    )
  );
  await measure(page, deck, 'slide sorter, then normal view', async () => {
    await exec(page, { type: 'setPresentationView', view: 'slideSorter' });
    await exec(page, { type: 'setPresentationView', view: 'normal' });
  });
  await measure(page, deck, 'save', () =>
    page.evaluate(
      () => void (window as unknown as { __pptxEditor: BenchEditor }).__pptxEditor.save()
    )
  );

  const mark = (ms: number) => (ms > HANG_MS ? 'hang' : ms > SLOW_MS ? 'slow' : '');
  console.log(
    [
      '| Deck | Action | Longest task ms | Long tasks ms | Wall ms | |',
      '| --- | --- | --: | --: | --: | --- |',
      ...rows.map(
        (r) =>
          `| ${r.deck} | ${r.action} | ${r.longest} | ${r.total} | ${r.wall} | ${mark(r.longest)} |`
      ),
    ].join('\n')
  );
  if (process.env.BENCH_STRICT)
    expect(rows.filter((r) => r.longest > HANG_MS).map((r) => r.action)).toEqual([]);
});
