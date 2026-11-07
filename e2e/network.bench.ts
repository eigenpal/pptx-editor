// Browser network bench: how the demo ends an open whose download stalls. Each scenario
// redirects one request to a local server that misbehaves, and records how long the demo takes
// to show the deck or an error. The open must never spin for longer than its stall limit.
//
// - deck stall: the deck request never answers. Expect the deck notice within the stall limit.
// - deck mid-body stall: half the deck, then nothing. Same expectation.
// - deck length lie: a 2 GiB `Content-Length`, 1000 bytes, then nothing. Same expectation.
// - deck trickle: 1 KiB every 500 ms for the first 64 KiB, then the rest. Expect the deck.
// - code 404: an editor module answers 404. Expect the boot failure message at once.
// - code stall: an editor module never answers. Expect the slow-load message.
// - late first deck: a file opened while the first deck trickles in stays open.
//
// Usage: BENCH_PORT=<port> bun run bench:network:browser

import { createServer, type Server } from 'node:http';
import { readFileSync } from 'node:fs';
import { fileURLToPath } from 'node:url';
import { expect, test, type Page } from '@playwright/test';

const PUBLIC = fileURLToPath(new URL('../examples/vite/public/', import.meta.url));
/** The demo's deck stall limit (`DECK_STALL_MS`), and its boot stall limit (index.html). */
const DECK_STALL_MS = 10_000;
const BOOT_STALL_MS = 15_000;
/** Time to notice and render, on top of each limit. */
const SLACK_MS = 3_000;

type Mode = 'stall' | 'midstall' | 'lie' | 'trickle';

let server: Server;
let origin = '';

test.beforeAll(async () => {
  server = createServer((req, res) => {
    const [, mode, file] = (req.url ?? '').split('/') as [string, Mode, string];
    const bytes = readFileSync(PUBLIC + file);
    const headers = { 'access-control-allow-origin': '*', 'content-type': 'application/zip' };
    if (mode === 'stall') return;
    if (mode === 'midstall') {
      res.writeHead(200, { ...headers, 'content-length': bytes.length });
      res.write(bytes.subarray(0, bytes.length >> 1));
      return;
    }
    if (mode === 'lie') {
      res.writeHead(200, { ...headers, 'content-length': 2 ** 31 });
      res.write(bytes.subarray(0, 1000));
      return;
    }
    res.writeHead(200, headers);
    let at = 0;
    const tick = (): void => {
      if (res.destroyed) return;
      if (at >= 64 * 1024) return void res.end(bytes.subarray(at));
      res.write(bytes.subarray(at, (at += 1024)));
      setTimeout(tick, 500);
    };
    tick();
  });
  await new Promise<void>((resolve) => server.listen(0, '127.0.0.1', resolve));
  const address = server.address();
  origin = `http://127.0.0.1:${typeof address === 'object' && address ? address.port : 0}`;
});

test.afterAll(async () => {
  server.closeAllConnections();
  await new Promise((resolve) => server.close(resolve));
});

const rows: { scenario: string; ms: number; outcome: string }[] = [];
test.afterAll(() => console.table(rows));

/** Send the first deck's request to the misbehaving server. */
async function misbehave(page: Page, mode: Mode): Promise<void> {
  await page.route('**/showcase.pptx', (route) =>
    route.continue({ url: `${origin}/${mode}/showcase.pptx` })
  );
}

/** Open the demo and wait for the deck, an editor notice, or the boot message. */
async function openDemo(page: Page, cap: number): Promise<{ ms: number; outcome: string }> {
  const t0 = Date.now();
  await page.goto('/', { waitUntil: 'commit' });
  const outcome = await page
    .waitForFunction(
      () => {
        const w = window as unknown as { __pptxEditor?: { snapshot(): { status: string } } };
        const boot = document.querySelector('[data-boot-failure]');
        if (boot) return `boot ${boot.getAttribute('data-boot-failure')}`;
        if (document.querySelector('.pptx-edit-notice__text')) return 'deck notice';
        if (document.querySelector('.pptx-open-failure')) return 'open failure';
        const ready = w.__pptxEditor?.snapshot().status === 'ready';
        return ready && document.querySelector('.pptx-surface__slide') ? 'deck' : false;
      },
      undefined,
      { timeout: cap, polling: 100 }
    )
    .then((handle) => handle.jsonValue() as Promise<string>);
  return { ms: Date.now() - t0, outcome };
}

const DECK_CASES: { mode: Mode; outcome: string; cap: number }[] = [
  { mode: 'stall', outcome: 'deck notice', cap: DECK_STALL_MS + SLACK_MS },
  { mode: 'midstall', outcome: 'deck notice', cap: DECK_STALL_MS + SLACK_MS },
  { mode: 'lie', outcome: 'deck notice', cap: DECK_STALL_MS + SLACK_MS },
  // About 32 seconds of bytes: longer than the stall limit in all, never stalled.
  { mode: 'trickle', outcome: 'deck', cap: 60_000 },
];

for (const { mode, outcome, cap } of DECK_CASES) {
  test(`deck ${mode}`, async ({ page }) => {
    await misbehave(page, mode);
    const result = await openDemo(page, cap);
    rows.push({ scenario: `deck ${mode}`, ...result });
    expect(result.outcome).toBe(outcome);
  });
}

test('code 404', async ({ page }) => {
  await page.route(/document-session/, (route) => route.fulfill({ status: 404, body: '' }));
  const result = await openDemo(page, SLACK_MS * 2);
  rows.push({ scenario: 'code 404', ...result });
  expect(result.outcome).toBe('boot failed');
});

test('code stall', async ({ page }) => {
  await page.route(/document-session/, () => undefined);
  const result = await openDemo(page, BOOT_STALL_MS + SLACK_MS * 2);
  rows.push({ scenario: 'code stall', ...result });
  expect(result.outcome).toBe('boot slow');
});

test('late first deck', async ({ page }) => {
  await misbehave(page, 'trickle');
  const t0 = Date.now();
  await page.goto('/');
  await page.waitForFunction(() => (window as { __pptxEditor?: unknown }).__pptxEditor);
  await page
    .locator('input[type=file][accept*=".pptx"]')
    .first()
    .setInputFiles({
      name: 'Mine.pptx',
      mimeType: 'application/zip',
      buffer: readFileSync(PUBLIC + 'sample.pptx'),
    });
  await page.waitForFunction(() => document.title.startsWith('Mine'));
  // Past the trickle: the first deck would have arrived by now.
  await page.waitForTimeout(40_000);
  const state = await page.evaluate(() => ({
    title: document.title,
    slides: (window as unknown as { __pptxEditor: { session: { info: { slides: unknown[] } } } })
      .__pptxEditor.session.info.slides.length,
    notice: document.querySelector('.pptx-edit-notice__text')?.textContent ?? null,
  }));
  rows.push({ scenario: 'late first deck', ms: Date.now() - t0, outcome: JSON.stringify(state) });
  expect(state.title.startsWith('Mine')).toBe(true);
  expect(state.slides).toBe(5);
  expect(state.notice).toBeNull();
});
