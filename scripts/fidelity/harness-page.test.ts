import { describe, expect, test } from 'bun:test';
import { execFileSync } from 'node:child_process';
import { createHash, randomBytes } from 'node:crypto';
import { mkdtempSync, readFileSync, rmSync, writeFileSync } from 'node:fs';
import { tmpdir } from 'node:os';
import { dirname, join } from 'node:path';
import { chromium, type Page } from '@playwright/test';
import { buildPptx } from '../fixtures/builder.ts';
import {
  fileDigest,
  openDeck,
  renderSlide,
  reopenSaved,
  slideMarkup,
  startHarnessServer,
  withHarnessPage,
} from './harness-page.ts';
import { collectionFace } from './local-fonts.ts';
import { BROWSER_CRASHED, BROWSER_HUNG, scoreSafely } from './score-deck.ts';

const fontsDir = join(
  // The packaged substitutes, a dependency of the engine.
  dirname(
    Bun.resolveSync(
      '@docx-editor.dev/fonts/package.json',
      join(import.meta.dir, '../../packages/core')
    )
  ),
  'assets'
);

describe('the harness server', () => {
  test('two runs get two servers on 127.0.0.1 at ports the system picks', async () => {
    const [a, b] = await Promise.all([startHarnessServer(), startHarnessServer()]);
    try {
      for (const s of [a, b]) expect(s.url).toMatch(/^http:\/\/127\.0\.0\.1:\d+\/$/);
      expect(a.url).not.toBe(b.url);
      const page = await fetch(`${a.url}harness.html`);
      expect(page.status).toBe(200);
      expect(await page.text()).toContain('harness');
    } finally {
      await Promise.all([a.close(), b.close()]);
    }
  }, 30_000);
});

/** The pids of this process's Chromium children (Playwright starts the browser directly). */
function browserPids(): number[] {
  let out = '';
  try {
    out = execFileSync('pgrep', ['-P', String(process.pid)], { encoding: 'utf8' });
  } catch {
    return [];
  }
  return out
    .split('\n')
    .filter(Boolean)
    .map(Number)
    .filter((pid) => {
      const name = execFileSync('ps', ['-p', String(pid), '-o', 'comm='], { encoding: 'utf8' });
      return /chrom|headless/i.test(name);
    });
}

/** The browser process of a page's browser, from CDP. */
async function browserPid(page: Page): Promise<number> {
  const cdp = await page.context().browser()!.newBrowserCDPSession();
  const info = (await cdp.send('SystemInfo.getProcessInfo')) as {
    processInfo: { type: string; id: number }[];
  };
  await cdp.detach();
  return info.processInfo.find((p) => p.type === 'browser')!.id;
}

/** Playwright's Chromium processes (pid, parent, group, state, type), for a failure. */
function chromiumProcesses(): string {
  const all = execFileSync('ps', ['-eo', 'pid=,ppid=,pgid=,stat=,command='], {
    encoding: 'utf8',
  });
  const rows: string[] = [];
  for (const line of all.split('\n')) {
    const [pid, ppid, pgid, stat, binary, ...args] = line.trim().split(/\s+/);
    if (!binary?.includes('ms-playwright')) continue;
    const type = args.find((a) => a.startsWith('--type='))?.slice(7) ?? 'browser';
    rows.push(`${pid} ${ppid} ${pgid} ${stat} ${type}`);
  }
  return rows.join('\n');
}

/** The processes of a process group, by its leader's pid. */
function groupMembers(pgid: number): number[] {
  try {
    return execFileSync('pgrep', ['-g', String(pgid)], { encoding: 'utf8' })
      .split('\n')
      .filter(Boolean)
      .map(Number);
  } catch {
    return [];
  }
}

/**
 * The unit job has no Playwright browser; the fidelity job installs it. The check launches the
 * browser that the harness launches (the headless shell), not only a path to a binary.
 */
const hasChromium = await chromium.launch().then(
  (b) => b.close().then(() => true),
  () => false
);

describe.skipIf(!hasChromium)('a browser that dies in the middle of a deck', () => {
  test('the deck is an error, and a new browser scores the next deck', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'pptx-crash-'));
    try {
      const deck = join(dir, 'deck.pptx');
      writeFileSync(deck, buildPptx({ slides: [{ shapes: '' }, { shapes: '' }] }));
      const input = { file: deck, name: 'd' };
      let stopping = 0;
      const [crashed, next] = await withHarnessPage(async (first, control) => {
        // Kill Chromium with SIGKILL when the deck renders its first slide, as an
        // out-of-memory kill does. The deck has opened by then.
        let killed = false;
        const dying = new Proxy(first, {
          get(target, key, receiver) {
            if (key !== 'evaluate') return Reflect.get(target, key, receiver);
            return async (fn: unknown, arg: unknown) => {
              if (!killed && String(fn).includes('.render(')) {
                killed = true;
                const browser = target.context().browser()!;
                const pid = await browserPid(target);
                const gone = new Promise<void>((done) =>
                  browser.once('disconnected', () => done())
                );
                console.log(`killing browser ${pid}:\n${chromiumProcesses()}`);
                process.kill(pid, 'SIGKILL');
                await gone;
              }
              return target.evaluate(fn as never, arg as never);
            };
          },
        }) as Page;
        const a = await scoreSafely(dying, control, input, { redact: true });
        expect(a.fatal).toBeUndefined();
        console.log(`after the reset (${a.result.error}):\n${chromiumProcesses()}`);
        const b = await scoreSafely(a.page, control, input, { redact: true });
        console.log(
          `the next deck took ${b.result.milliseconds ?? '?'} ms${b.result.error ? `, and failed:\n${chromiumProcesses()}` : ''}`
        );
        stopping = Date.now();
        return [a, b];
      });
      console.log(`the harness stopped in ${Date.now() - stopping} ms`);
      expect(crashed.result.error).toStartWith(
        `${BROWSER_CRASHED} (the next deck gets a new browser): `
      );
      expect(crashed.result.slides).toEqual([]);
      expect(next.result.error).toBeUndefined();
      expect(next.result.slides).toHaveLength(2);
      expect(next.fatal).toBeUndefined();
    } finally {
      rmSync(dir, { recursive: true, force: true });
    }
  }, 120_000);

  test('a browser that stops answering is replaced, and the next deck is scored', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'pptx-hang-'));
    const stopped: number[] = [];
    let stopping = 0;
    try {
      const deck = join(dir, 'deck.pptx');
      writeFileSync(deck, buildPptx({ slides: [{ shapes: '' }] }));
      const input = { file: deck, name: 'd' };
      const [hung, next] = await withHarnessPage(async (first, control) => {
        // SIGSTOP: the browser stays connected but answers nothing, as a busy browser does.
        stopped.push(...browserPids());
        expect(stopped.length).toBeGreaterThan(0);
        for (const pid of stopped) process.kill(pid, 'SIGSTOP');
        const a = await scoreSafely(first, control, input, { redact: true, timeoutMs: 1_000 });
        const b = await scoreSafely(a.page, control, input, { redact: true });
        console.log(`the next deck took ${b.result.milliseconds ?? '?'} ms`);
        stopping = Date.now();
        return [a, b];
      }, 3_000);
      console.log(`the harness stopped in ${Date.now() - stopping} ms`);
      expect(hung.result.error).toBe(
        `${BROWSER_HUNG} (the next deck gets a new browser): the deck did not finish within 1 s`
      );
      expect(next.result.error).toBeUndefined();
      expect(next.result.slides).toHaveLength(1);
      // The hung browser does not keep running: the reset killed its whole process group.
      const alive = stopped.filter((pid) => {
        try {
          process.kill(pid, 0);
          return true;
        } catch {
          return false;
        }
      });
      expect(alive).toEqual([]);
      for (const pid of stopped) expect(groupMembers(pid)).toEqual([]);
    } finally {
      for (const pid of stopped) {
        try {
          process.kill(pid, 'SIGKILL');
        } catch {
          // Playwright already stopped it.
        }
      }
      rmSync(dir, { recursive: true, force: true });
    }
  }, 120_000);
});

describe.skipIf(!hasChromium)('openDeck', () => {
  test('a missing deck or font clears the previous document and returns no local path', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'pptx-missing-'));
    try {
      const deck = join(dir, 'deck.pptx');
      const missing = join(dir, 'private-removed-font.ttf');
      writeFileSync(deck, buildPptx({ slides: [{ shapes: '' }] }));
      await withHarnessPage(async (page) => {
        const responses: number[] = [];
        page.on('response', (response) => {
          if (response.url().includes('/__harness/file/') && response.status() === 404)
            responses.push(response.status());
        });
        for (const fontMissing of [false, true]) {
          await openDeck(page, deck, []);
          await renderSlide(page, 0, 960);
          const attempt = fontMissing
            ? openDeck(page, deck, [
                { family: 'Missing', bold: false, italic: false, file: missing },
              ])
            : openDeck(page, missing, []);
          await expect(attempt).rejects.toThrow(
            `${fontMissing ? 'the font Missing' : 'the deck'} did not arrive from the harness server (HTTP 404)`
          );
          await expect(renderSlide(page, 0, 960)).rejects.toThrow('no document');
          await expect(reopenSaved(page)).rejects.toThrow('no document');
        }
        expect(responses).toEqual([404, 404]);
        expect((await openDeck(page, deck, [])).slides).toBe(1);
      });
    } finally {
      rmSync(dir, { recursive: true, force: true });
    }
  }, 60_000);

  test('the page fetches the deck and each font, and gets every byte', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'pptx-open-'));
    try {
      const deck = join(dir, 'deck.pptx');
      writeFileSync(deck, buildPptx({ slides: [{ shapes: '' }] }));
      // Not fonts: the page only keeps them, and a face that does not load is skipped.
      const font = (family: string, bytes: Buffer) => {
        const file = join(dir, `${family}.ttf`);
        writeFileSync(file, bytes);
        return { family, bold: false, italic: false, file };
      };
      const fonts = [font('A', randomBytes(3 * 1024 * 1024)), font('B', randomBytes(1025))];
      // More than one DevTools message can carry (about 100 MiB): deleted archive issue 82.
      const large = font('Large', Buffer.alloc(110 * 1024 * 1024, 7));
      const [first, second, rendered, after] = await withHarnessPage(async (page) => {
        const fetched: string[] = [];
        page.on('request', (r) => {
          if (r.url().includes('/__harness/file/')) fetched.push(r.url());
        });
        const a = await openDeck(page, deck, [...fonts, large]);
        // The next deck gets only its own fonts.
        const b = await openDeck(page, deck, []);
        const r = await renderSlide(page, 0, 960);
        // After `openDeck`, its paths, and any other path, give nothing.
        const files = fetched.length;
        const urls = [...fetched, '/__harness/file/not-a-token'];
        const statuses = await page.evaluate(
          (list) => Promise.all(list.map(async (u) => (await fetch(u)).status)),
          urls
        );
        return [a, b, r, { files, statuses }];
      });
      const digest = (f: { family: string; file: string }) => ({
        family: f.family,
        ...fileDigest(f.file),
      });
      expect(first).toEqual({
        slides: 1,
        received: { deck: fileDigest(deck), fonts: [...fonts, large].map(digest) },
      });
      expect(second.received.fonts).toEqual([]);
      expect(rendered.png.length).toBeGreaterThan(0);
      // The deck and three fonts, then the deck alone.
      expect(after.files).toBe(5);
      expect(new Set(after.statuses)).toEqual(new Set([404]));
    } finally {
      rmSync(dir, { recursive: true, force: true });
    }
  }, 120_000);

  test('a face of a font collection reaches the page as that face alone', async () => {
    const dir = mkdtempSync(join(tmpdir(), 'pptx-collection-'));
    try {
      const deck = join(dir, 'deck.pptx');
      writeFileSync(deck, buildPptx({ slides: [{ shapes: '' }] }));
      // A collection of one face: a 16-byte `ttcf` header, then the font, its offsets moved.
      const font = new Uint8Array(readFileSync(join(fontsDir, 'LiberationSans-Regular.ttf')));
      const ttc = new Uint8Array(16 + font.length);
      const v = new DataView(ttc.buffer);
      v.setUint32(0, 0x74746366);
      v.setUint32(4, 0x00010000);
      v.setUint32(8, 1);
      v.setUint32(12, 16);
      ttc.set(font, 16);
      for (let i = 0; i < v.getUint16(16 + 4); i += 1)
        v.setUint32(16 + 12 + i * 16 + 8, v.getUint32(16 + 12 + i * 16 + 8) + 16);
      const file = join(dir, 'collection.ttc');
      writeFileSync(file, ttc);
      const opened = await withHarnessPage((page) =>
        openDeck(page, deck, [
          { family: 'Deck Sans', bold: false, italic: false, file, faceIndex: 0 },
        ])
      );
      const face = collectionFace(ttc, 0)!;
      expect(opened.received.fonts).toEqual([
        {
          family: 'Deck Sans',
          bytes: face.length,
          sha256: createHash('sha256').update(face).digest('hex'),
        },
      ]);
    } finally {
      rmSync(dir, { recursive: true, force: true });
    }
  }, 60_000);
});

describe.skipIf(!hasChromium)('slideMarkup', () => {
  test('a saved and reopened deck paints the same markup, with no session-bound names', async () => {
    // Image blob URLs and face ids are new in each session.
    await withHarnessPage(async (page) => {
      for (const [name, named] of [
        ['pictures', 'image:'],
        ['text-runs', 'face:'],
      ]) {
        const { slides } = await openDeck(
          page,
          join(import.meta.dir, `../../fixtures/${name}.pptx`),
          []
        );
        const first: string[] = [];
        for (let i = 0; i < slides; i += 1) first.push(await slideMarkup(page, i, 960));
        expect(first.join('')).toContain(named);
        expect(first.join('')).not.toMatch(/blob:|pptx-f\d/);
        expect((await reopenSaved(page)).slides).toBe(slides);
        for (let i = 0; i < slides; i += 1) expect(await slideMarkup(page, i, 960)).toBe(first[i]);
      }
    });
  }, 60_000);
});
