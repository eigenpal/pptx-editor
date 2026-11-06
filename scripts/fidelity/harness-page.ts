// The fidelity harness page, driven from Node. Every fidelity command renders through these
// helpers: the demo's Vite server serves `examples/vite/harness.html`, Playwright Chromium opens
// it, and the page renders with the editor's own `openPresentation` → `layoutSlide` →
// `paintSlide`. There is no second renderer.

import { chromium, type Browser, type Page } from '@playwright/test';
import { createHash, randomUUID } from 'node:crypto';
import {
  createReadStream,
  mkdirSync,
  mkdtempSync,
  readdirSync,
  readFileSync,
  rmSync,
  statSync,
} from 'node:fs';
import { tmpdir } from 'node:os';
import { join, resolve } from 'node:path';
import { createServer, type Plugin } from 'vite';
import { collectionFace, deckFontFamilies, localFonts, type LocalFont } from './local-fonts.ts';

/** One painted slide (`renderSlide`). */
export interface HarnessRender {
  readonly diagnostics: readonly string[];
  /** The families that the slide's text asks for. */
  readonly fonts: readonly string[];
  /**
   * The painted SVG, with image URLs and face ids replaced by the digests of their bytes, so two
   * sessions that paint the same thing give the same markup.
   */
  readonly markup: string;
}

export interface HarnessServer {
  /** The server's address, for example `http://127.0.0.1:53712/`. */
  readonly url: string;
  close(): Promise<void>;
}

/**
 * The folder of the harness servers' Vite caches, one folder for each server. One root for each
 * user: in a shared temp folder, another user's folders cannot be removed.
 */
const VITE_CACHE_ROOT = join(tmpdir(), `pptx-harness-vite-${process.getuid?.() ?? 'user'}`);

/**
 * A new Vite cache folder for one harness server. Lanes that share the demo's folder write it at
 * the same time and fail with ENOTEMPTY. Each folder starts with its process id, and the folders
 * of processes that no longer run (a lane that was killed) are removed first.
 */
function viteCacheDir(): string {
  mkdirSync(VITE_CACHE_ROOT, { recursive: true });
  for (const entry of readdirSync(VITE_CACHE_ROOT)) {
    const pid = Number(entry.split('-')[0]);
    let running = false;
    try {
      running = Number.isInteger(pid) && pid > 0 && process.kill(pid, 0);
    } catch (e) {
      // EPERM: the process runs under another user.
      running = (e as NodeJS.ErrnoException).code === 'EPERM';
    }
    try {
      if (!running) rmSync(join(VITE_CACHE_ROOT, entry), { recursive: true, force: true });
    } catch {
      // A folder that cannot be removed only takes space.
    }
  }
  return mkdtempSync(join(VITE_CACHE_ROOT, `${process.pid}-`));
}

/**
 * Start the demo's Vite server for the harness page. Harness runs happen in parallel (corpus
 * shards, other checkouts). So the server listens on 127.0.0.1 at a port that the system picks,
 * and the page opens that exact address: with a fixed port and `localhost`, a page could load
 * another checkout's server on `[::1]` and render its code. Without HMR and a file watcher,
 * edits during a run do not reload the page.
 */
export async function startHarnessServer(): Promise<HarnessServer> {
  const cacheDir = viteCacheDir();
  const server = await createServer({
    configFile: resolve('examples/vite/vite.config.ts'),
    root: resolve('examples/vite'),
    plugins: [harnessFiles()],
    server: { host: '127.0.0.1', port: 0, strictPort: true, hmr: false, watch: null },
    // The harness page imports only ES modules, so the server discovers no dependencies. The
    // React plugin still adds React, which the harness page never loads.
    optimizeDeps: { noDiscovery: true, include: [] },
    cacheDir,
    logLevel: 'error',
  });
  const close = async (): Promise<void> => {
    try {
      await server.close();
    } finally {
      rmSync(cacheDir, { recursive: true, force: true });
    }
  };
  await server.listen();
  const url = server.resolvedUrls?.local[0];
  if (!url) {
    await close();
    throw new Error('the harness server has no local address');
  }
  return { url, close };
}

/** The path under which the harness server gives a file to its page (`openDeck`). */
const FILES_PATH = '/__harness/file/';

/**
 * The files that a harness page may fetch now, by a random token: a deck and its fonts while
 * `openDeck` runs. A face of a font collection is served from memory, as a single font. Every harness server of this process reads this one table, so a page fetches
 * from its own server and needs no address. Only the files that `openDeck` adds are served,
 * and only until it removes them.
 */
const servedFiles = new Map<string, string | Uint8Array>();

/** Serve `file` (a path, or bytes) at a new random path until `release` runs. */
function serveFile(file: string | Uint8Array): { readonly url: string; release(): void } {
  const token = randomUUID();
  servedFiles.set(token, file);
  return { url: `${FILES_PATH}${token}`, release: () => servedFiles.delete(token) };
}

/** The path under which the page sends bytes back to Node (`receive`). */
const UPLOAD_PATH = '/__harness/upload/';

/** The uploads that Node waits for, by a random token. */
const uploads = new Map<string, (bytes: Uint8Array<ArrayBuffer>) => void>();

/**
 * A path where the page can send bytes once (HTTP PUT), and the bytes when they arrive. Large
 * pixel arrays go this way: through `page.evaluate` they would be JSON.
 */
function receive(): {
  readonly url: string;
  readonly bytes: Promise<Uint8Array<ArrayBuffer>>;
  release(): void;
} {
  const token = randomUUID();
  const bytes = new Promise<Uint8Array<ArrayBuffer>>((resolve) => uploads.set(token, resolve));
  return { url: `${UPLOAD_PATH}${token}`, bytes, release: () => uploads.delete(token) };
}

/** The Vite middleware for `servedFiles` and `uploads`. */
function harnessFiles(): Plugin {
  return {
    name: 'pptx-harness-files',
    configureServer(server) {
      server.middlewares.use(UPLOAD_PATH, (req, res) => {
        const token = (req.url ?? '').slice(1);
        const done = uploads.get(token);
        if (!done || req.method !== 'PUT') {
          res.statusCode = 404;
          res.end();
          return;
        }
        uploads.delete(token);
        const chunks: Buffer[] = [];
        req.on('data', (c: Buffer) => chunks.push(c));
        req.on('end', () => {
          done(new Uint8Array(Buffer.concat(chunks)));
          res.end();
        });
      });
      server.middlewares.use(FILES_PATH, (req, res) => {
        const file = servedFiles.get((req.url ?? '').slice(1));
        if (file instanceof Uint8Array) {
          res.setHeader('Content-Type', 'application/octet-stream');
          res.setHeader('Content-Length', String(file.length));
          res.setHeader('Cache-Control', 'no-store');
          res.end(file);
          return;
        }
        let size: number | undefined;
        try {
          if (file) size = statSync(file).size;
        } catch {
          // Removed since `openDeck` served it: the same answer as an unknown token, so that no
          // path reaches a log.
        }
        if (!file || size === undefined) {
          res.statusCode = 404;
          res.end();
          return;
        }
        res.setHeader('Content-Type', 'application/octet-stream');
        res.setHeader('Content-Length', String(size));
        res.setHeader('Cache-Control', 'no-store');
        // A read error, or a file that changed size, cuts the response short, and the page's
        // fetch fails with the name of the deck or the font.
        createReadStream(file)
          .on('error', () => res.destroy())
          .pipe(res);
      });
    },
  };
}

/** The result of `HarnessControl.reset`. */
export interface HarnessReset {
  /** The page for the next deck. */
  readonly page: Page;
  /**
   * Why the reset started a new browser: the old one `exited` (it crashed or was killed), or it
   * `hung` (no new page opened in it in time). Unset when the old browser opened the new page,
   * or when the new page failed for another reason (the reset still uses a new browser).
   */
  readonly newBrowser?: 'exited' | 'hung';
}

export interface HarnessControl {
  /**
   * Close the page and open a new one: after a deck crashed or hung the page. When the browser
   * is gone, or no new page opens in it, start a new browser first. Throws only when the new
   * browser cannot open the page either.
   */
  reset(): Promise<HarnessReset>;
  /**
   * Close the page and its browser context, and open the harness page in a new context. A page
   * keeps font and glyph caches across decks, so a deck can paint differently after another
   * deck. A new context for each deck paints it as a run of that deck alone does (`lane.ts`).
   */
  restart(): Promise<Page>;
  /**
   * A second harness page that only draws PDFs (`rasterPdf`). It stays open across decks, so
   * pdf.js loads once for each browser; a new one opens when the browser changed or it closed.
   */
  rasterPage(): Promise<Page>;
}

/** A page reset that takes longer than this fails, so a hung browser cannot stop a run. */
export const RESET_TIMEOUT_MS = 60_000;

/**
 * A browser close that takes longer than this ends with a kill. A close takes about 0.1 s. On
 * Linux (CI, Chromium headless shell 1248), the close of the browser that replaced a crashed one
 * never ended: Playwright waits for the browser's output pipes, and they stayed open. The cause
 * is not known; the kill ends the close.
 */
export const CLOSE_TIMEOUT_MS = 10_000;

/** `within` ran out of time. */
class TimeLimit extends Error {
  override readonly name = 'TimeLimit';
}

/** Reject when `work` does not settle within `ms`. */
function within<T>(work: Promise<T>, ms: number, what: string): Promise<T> {
  let timer: ReturnType<typeof setTimeout> | undefined;
  return Promise.race([
    work,
    new Promise<never>((_, reject) => {
      timer = setTimeout(
        () => reject(new TimeLimit(`${what} did not finish within ${ms / 1000} s`)),
        ms
      );
    }),
  ]).finally(() => clearTimeout(timer));
}

/** A launched Chromium and the pid of its browser process, when CDP gives it. */
interface Launched {
  readonly browser: Browser;
  pid?: number;
  /** Playwright lost the browser: it exited, crashed, or was closed. */
  exited: boolean;
}

async function launch(): Promise<Launched> {
  const browser = await chromium.launch();
  const launched: Launched = { browser, exited: false };
  browser.on('disconnected', () => {
    launched.exited = true;
  });
  try {
    const cdp = await browser.newBrowserCDPSession();
    const info = (await cdp.send('SystemInfo.getProcessInfo')) as {
      processInfo: { type: string; id: number }[];
    };
    await cdp.detach();
    launched.pid = info.processInfo.find((p) => p.type === 'browser')?.id;
  } catch {
    // No pid: a browser that hangs cannot be killed (closeBrowser warns).
  }
  return launched;
}

/**
 * Close a browser within `ms`. Playwright's `close()` never ends for a browser that does not
 * answer, and that browser keeps running. So it is then killed: its process group (Playwright
 * starts the browser as a group leader, so the group holds the browser and its GPU and renderer
 * children, as DOCX Editor's validator kills a stage's group), else its pid.
 */
async function closeBrowser(launched: Launched, ms: number): Promise<void> {
  try {
    await within(launched.browser.close(), ms, 'closing the browser');
  } catch {
    if (launched.pid === undefined) {
      console.warn(
        'fidelity: a browser that does not answer keeps running until the run ends: CDP gave no pid to stop it'
      );
      return;
    }
    for (const target of [-launched.pid, launched.pid]) {
      try {
        process.kill(target, 'SIGKILL');
        return;
      } catch {
        // No such group or process: try the pid, or it exited in the meantime.
      }
    }
  }
}

/** `HarnessControl.reset` could not open a page even in a new browser. */
export class ResetFailed extends Error {
  override readonly name = 'ResetFailed';
  constructor(
    message: string,
    /** Why the reset started a new browser before it failed (see `HarnessReset`). */
    readonly newBrowser: 'exited' | 'hung' | undefined,
    options?: ErrorOptions
  ) {
    super(message, options);
  }
}

/**
 * Start the Vite server and Chromium, open the harness page, run `fn` with it, and shut both
 * down, also when `fn` throws.
 */
export async function withHarnessPage<T>(
  fn: (page: Page, control: HarnessControl) => Promise<T>,
  /** Tests only: a shorter limit for each reset step. */
  resetTimeoutMs = RESET_TIMEOUT_MS
): Promise<T> {
  const server = await startHarnessServer();
  let current: Launched | undefined;
  try {
    current = await launch();
    const open = async (): Promise<Page> => {
      const next = await current!.browser.newPage({ viewport: { width: 2000, height: 1300 } });
      await next.goto(`${server.url}harness.html`);
      await next.waitForSelector('body[data-ready="1"]', { state: 'attached' });
      return next;
    };
    let page = await open();
    let raster: Page | undefined;
    return await fn(page, {
      async rasterPage() {
        if (!raster || raster.isClosed() || raster.context().browser() !== current!.browser)
          raster = await within(open(), resetTimeoutMs, 'opening the raster page');
        return raster;
      },
      async restart() {
        // `browser.newPage` gives each page its own context, and closing the page closes it.
        await within(page.close(), resetTimeoutMs, 'closing the page');
        page = await within(open(), resetTimeoutMs, 'opening a new page');
        return page;
      },
      async reset() {
        // Playwright can report a crashed browser as connected for a moment after the deck's
        // call fails, so the cause is read here, after a new page did not open: the
        // `disconnected` event, or Playwright's "has been closed" when the event comes late.
        let newBrowser: 'exited' | 'hung' | undefined = 'exited';
        if (!current!.exited) {
          await within(page.close(), resetTimeoutMs, 'closing the page').catch(() => undefined);
          try {
            page = await within(open(), resetTimeoutMs, 'opening a new page');
            return { page };
          } catch (e) {
            // "Target crashed" names a crashed page, not the browser, so it does not count.
            const closed = /browser has been closed/i.test(String((e as Error).message ?? e));
            if (!current!.exited && !closed)
              newBrowser = e instanceof TimeLimit ? 'hung' : undefined;
          }
        }
        // A deck can crash or hang the whole browser, and every later page call then fails. As
        // DOCX Editor's validator gives each document a new process, the next deck gets a new
        // browser. Closing a dead browser only frees Playwright's side of it.
        await closeBrowser(current!, Math.min(resetTimeoutMs, CLOSE_TIMEOUT_MS));
        try {
          current = await within(launch(), resetTimeoutMs, 'starting a new browser');
          page = await within(open(), resetTimeoutMs, 'opening a page in the new browser');
        } catch (e) {
          throw new ResetFailed(
            `a new browser did not open the harness page: ${(e as Error).message ?? e}`,
            newBrowser,
            { cause: e }
          );
        }
        return { page, newBrowser };
      },
    });
  } finally {
    if (current) await closeBrowser(current, Math.min(resetTimeoutMs, CLOSE_TIMEOUT_MS));
    await server.close();
  }
}

/** The size and SHA-256 of a file that the harness page got. */
export interface ReceivedFile {
  readonly bytes: number;
  readonly sha256: string;
}

/** The result of `openDeck`. */
export interface OpenedDeck {
  readonly slides: number;
  readonly error?: string;
  /** What the page got: the deck and each font, so a test can check every byte. */
  readonly received: {
    readonly deck: ReceivedFile;
    readonly fonts: readonly (ReceivedFile & { readonly family: string })[];
  };
}

/** The size and SHA-256 of a file on disk, as `OpenedDeck.received` gives them. */
export function fileDigest(file: string): ReceivedFile {
  const bytes = readFileSync(file);
  return { bytes: bytes.length, sha256: createHash('sha256').update(bytes).digest('hex') };
}

/**
 * Open a deck in the harness page with the local reference fonts that the deck names (or `fonts`).
 * The page fetches the deck and the fonts from the harness server as raw bytes. They never go
 * through `page.evaluate`: Chromium ends the DevTools connection on one message of about
 * 100 MiB, and a deck whose theme names large East Asian fonts needs more than that
 * (deleted archive issue 82).
 */
export async function openDeck(
  page: Page,
  file: string,
  fonts?: readonly LocalFont[]
): Promise<OpenedDeck> {
  const faces: LocalFont[] = [];
  const sources: (string | Uint8Array)[] = [];
  for (const f of fonts ?? localFonts(deckFontFamilies(new Uint8Array(readFileSync(file))))) {
    const source =
      f.faceIndex === undefined
        ? f.file
        : collectionFace(new Uint8Array(readFileSync(f.file)), f.faceIndex);
    if (source === undefined) continue;
    faces.push(f);
    sources.push(source);
  }
  const served = [serveFile(file), ...sources.map(serveFile)];
  try {
    const source = {
      deck: served[0].url,
      fonts: faces.map((f, i) => ({
        family: f.family,
        bold: f.bold,
        italic: f.italic,
        url: served[i + 1].url,
      })),
    };
    return await page.evaluate((s) => (window as any).__pptxHarness.open(s), source);
  } finally {
    for (const s of served) s.release();
  }
}

/** Save the open deck, reopen the saved bytes, and make them the open deck. */
export async function reopenSaved(page: Page): Promise<{ slides: number; error?: string }> {
  return page.evaluate(() => (window as any).__pptxHarness.reopenSaved());
}

/** Paint one slide (0-based) of the open deck, and return its diagnostics and markup. */
export async function renderSlide(page: Page, index: number): Promise<HarnessRender> {
  return page.evaluate((n) => (window as any).__pptxHarness.render(n), index);
}

/** A PDF page drawn by the harness page (`rasterPdf`): its pixels (RGBA) and their PNG. */
export interface PageRaster {
  readonly width: number;
  readonly height: number;
  readonly data: Uint8Array;
  readonly png: Uint8Array;
}

/**
 * Draw the first `pages` pages of PDF bytes with pdf.js in the harness page, at `scale` pixels
 * per point, each onto a white canvas of `size`. Both PDFs of a comparison are drawn this way.
 * Glyphs are drawn as paths (`disableFontFace`), so no system font rasterizer takes part.
 */
export async function rasterPdf(
  page: Page,
  pdf: Uint8Array,
  pages: number,
  size: { readonly width: number; readonly height: number },
  scale: number
): Promise<PageRaster[]> {
  const served = serveFile(pdf);
  const targets = Array.from({ length: pages }, () => ({ rgba: receive(), png: receive() }));
  try {
    await page.evaluate((a) => (window as any).__pptxHarness.rasterPdf(a), {
      pdf: served.url,
      ...size,
      scale,
      pages: targets.map((t) => ({ rgba: t.rgba.url, png: t.png.url })),
    });
    return await Promise.all(
      targets.map(async (t) => {
        const data = Bun.gunzipSync(await t.rgba.bytes);
        if (data.length !== size.width * size.height * 4)
          throw new Error(
            `a page raster has ${data.length} bytes, not ${size.width} × ${size.height} pixels`
          );
        return { ...size, data, png: await t.png.bytes };
      })
    );
  } finally {
    served.release();
    for (const t of targets) {
      t.rgba.release();
      t.png.release();
    }
  }
}

/**
 * Print the slides at `indexes` (0-based) of the open deck to PDF with the editor's own paint:
 * one page for each slide, at the slide size.
 */
export async function printDeck(page: Page, indexes: readonly number[]): Promise<Buffer> {
  await page.evaluate((list) => (window as any).__pptxHarness.print(list), [...indexes]);
  try {
    return await page.pdf({
      preferCSSPageSize: true,
      printBackground: true,
      margin: { top: '0', right: '0', bottom: '0', left: '0' },
    });
  } finally {
    await page.evaluate(() => (window as any).__pptxHarness.endPrint());
  }
}
