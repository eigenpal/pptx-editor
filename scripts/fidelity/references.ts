// Reference helpers: rasterize the pages of a deck's reference PDF, and extract its text lines with
// their baselines. Uses Poppler (`pdftoppm`, `pdftotext`) and pdf.js as test tools only; nothing
// here is bundled.

import { execFileSync, spawnSync } from 'node:child_process';
import { createHash } from 'node:crypto';
import {
  mkdirSync,
  readdirSync,
  readFileSync,
  renameSync,
  rmSync,
  statSync,
  utimesSync,
  writeFileSync,
} from 'node:fs';
import { join } from 'node:path';
import { PNG } from 'pngjs';
import { EMU_PER_POINT } from '../../packages/core/src/store/index.ts';
import { downsample, type Raster } from './supersample.ts';
import { THRESHOLDS } from './thresholds.ts';

/** The Poppler version (`pdftoppm -v`), or null when Poppler is not installed. */
export function popplerVersion(): string | null {
  const r = spawnSync('pdftoppm', ['-v'], { encoding: 'utf8' });
  if (r.error) return null;
  // `pdftoppm -v` prints to stderr, and some versions exit 99.
  return /version ([\d.]+)/.exec(`${r.stdout}${r.stderr}`)?.[1] ?? 'unknown';
}

/** Whether Poppler is installed. */
export const hasPoppler = (): boolean => popplerVersion() !== null;

/**
 * Each Poppler call may take this long. It runs synchronously, so without a limit a hung tool
 * would block the harness and its per-deck time limit.
 */
export const POPPLER = { timeout: 120_000, killSignal: 'SIGKILL' } as const;

/**
 * The time limit to rasterize a whole PDF: 120 s, and 5 s more for each page at scale 1. On a
 * busy machine a 51-page deck took 188 s at scale 1 (3.7 s a page). A raster at `scale` has
 * `scale²` times the pixels, and on the fixtures it took about that much longer (16 times at
 * scale 4), so the time for each page grows the same way. The deck's own limit of 10 minutes
 * stays the outer bound (`score-deck.ts`).
 */
export const rasterLimitMs = (pages: number, scale: number = THRESHOLDS.rasterScale): number =>
  POPPLER.timeout + 5_000 * scale * scale * pages;

/** Whether a Poppler call failed because it ran out of time, not because something killed it. */
export const timedOut = (e: unknown): boolean => (e as NodeJS.ErrnoException).code === 'ETIMEDOUT';

/** A slide size (`p:sldSz`) in EMU. */
export interface SlideSize {
  readonly widthEmu: number;
  readonly heightEmu: number;
}

/** A raster size in pixels. */
export interface SizePx {
  readonly width: number;
  readonly height: number;
}

/**
 * The size of a slide render at `widthPx` wide. It computes the height from EMU as
 * `examples/vite/src/harness.ts` does, operation for operation, so both round the same way.
 */
export function renderSizePx(slide: SlideSize, widthPx: number): SizePx {
  return { width: widthPx, height: Math.round((widthPx * slide.heightEmu) / slide.widthEmu) };
}

/** The `pdftoppm -r` resolution that rasterizes the slide at `widthPx` wide, in dots per inch. */
const rasterDpi = (slide: SlideSize, widthPx: number): number =>
  (72 * widthPx) / (slide.widthEmu / EMU_PER_POINT);

/**
 * The raster size that `pdftoppm` gives for one page with `rasterOptions`: the render size, but
 * no more than the page holds. When the reference PDF rounds a page side down, the page ends less
 * than 0.5 pt before the slide, and the raster can be short on that side (1 px on the 400.25 pt
 * wide `slide-size-rounding` fixture). What the page does not hold has no reference, and the
 * comparison leaves it out.
 */
export function expectedRasterSizePx(
  slide: SlideSize,
  page: readonly [number, number],
  widthPx: number
): SizePx {
  const render = renderSizePx(slide, widthPx);
  const dpi = rasterDpi(slide, widthPx);
  // `pdftoppm` makes a page side of w pt `ceil(w × dpi / 72)` px, with the same floating-point
  // operations, as DOCX Editor's `render_size` (`packages/docx-to-pdf/scripts/pdf-visual-diff.py`).
  return {
    width: Math.min(render.width, Math.ceil((page[0] * dpi) / 72)),
    height: Math.min(render.height, Math.ceil((page[1] * dpi) / 72)),
  };
}

/**
 * The `pdftoppm` options that rasterize the slide area of a page at `scale` times the render's
 * pixels per point (`THRESHOLDS.rasterScale`; `rasterize` filters it down to the render size).
 *
 * A reference PDF has each page side as a whole number of points (an A4 slide of 842 × 595.5 pt
 * gets an 842 × 596 pt page, a slide 400.25 pt wide a 400 pt wide one), and draws the slide at
 * scale 1 from the page's top-left corner. Scaling the page to the render width would add the
 * strip to the raster and rescale the slide when the width was rounded. So the page is rendered
 * at the render's pixels per point and cropped to the slide.
 */
export function rasterOptions(slide: SlideSize, widthPx: number, scale = 1): string[] {
  const size = renderSizePx(slide, widthPx);
  return [
    // The crop box, as `pdfinfo` reports the page size; reference PDFs have it equal to
    // the media box.
    '-cropbox',
    '-r',
    String(rasterDpi(slide, widthPx) * scale),
    '-x',
    '0',
    '-y',
    '0',
    // `scale` times the render size, so the filtered raster has the render size.
    '-W',
    String(size.width * scale),
    '-H',
    String(size.height * scale),
  ];
}

/** One rasterized reference page (`rasterizePages`). */
export interface ReferencePage {
  /** The page filtered down to the render width, as a PNG file. */
  readonly path: string;
  /** The pixels of that PNG. */
  readonly image: PNG;
  /** The page before the filter, at the raster scale. Decompressed at each call. */
  device(): Raster;
}

const isSpace = (b: number | undefined): boolean =>
  b === 0x20 || b === 0x0a || b === 0x0d || b === 0x09;

/** The pixels of a binary PPM (`P6`, 8 bits) as RGBA, opaque, as a decoded RGB PNG gives them. */
export function readPpm(bytes: Uint8Array): Raster {
  if (bytes[0] !== 0x50 || bytes[1] !== 0x36) throw new Error('not a binary PPM');
  const fields: number[] = [];
  let i = 2;
  while (fields.length < 3) {
    while (isSpace(bytes[i])) i += 1;
    let n = 0;
    const start = i;
    for (; bytes[i] >= 0x30 && bytes[i] <= 0x39; i += 1) n = n * 10 + bytes[i] - 0x30;
    if (i === start) throw new Error('a PPM header field is not a number');
    fields.push(n);
  }
  i += 1; // The one whitespace byte after the maximum value.
  const [width, height, max] = fields;
  if (max !== 255) throw new Error(`a PPM with maximum value ${max} is not 8 bits`);
  if (bytes.length - i < width * height * 3) throw new Error('the PPM is short');
  const data = new Uint8Array(width * height * 4);
  for (let o = 0, end = data.length; o < end; o += 4, i += 3) {
    data[o] = bytes[i];
    data[o + 1] = bytes[i + 1];
    data[o + 2] = bytes[i + 2];
    data[o + 3] = 255;
  }
  return { width, height, data };
}

/** A device raster kept compressed (zstd) until a comparison needs it. */
interface Packed {
  readonly width: number;
  readonly height: number;
  readonly zst: Uint8Array;
}

const pack = (r: Raster): Packed => ({
  width: r.width,
  height: r.height,
  zst: Bun.zstdCompressSync(r.data, { level: 1 }),
});

function unpack(p: Packed): Raster {
  const data = new Uint8Array(Bun.zstdDecompressSync(p.zst));
  if (data.length !== p.width * p.height * 4) throw new Error('a packed raster has the wrong size');
  return { width: p.width, height: p.height, data };
}

/** Change it when the cache layout, or the way a page is rasterized, changes. */
const RASTER_CACHE_FORMAT = 1;

/**
 * The cache entry of a reference raster: the PDF's SHA-256, the Poppler version, and the exact
 * `pdftoppm` options. A change to any of them rasterizes the PDF again.
 */
export function rasterCacheKey(pdf: string, options: readonly string[]): string {
  return createHash('sha256')
    .update(
      JSON.stringify({
        format: RASTER_CACHE_FORMAT,
        pdf: createHash('sha256').update(readFileSync(pdf)).digest('hex'),
        poppler: popplerVersion(),
        options,
      })
    )
    .digest('hex')
    .slice(0, 32);
}

/** The pages of a cache entry, or undefined when it is missing, incomplete, or damaged. */
function readCached(dir: string): Raster[] | undefined {
  try {
    const meta = JSON.parse(readFileSync(join(dir, 'meta.json'), 'utf8')) as {
      pages: { width: number; height: number }[];
    };
    const pages = meta.pages.map((p, n) =>
      unpack({
        width: p.width,
        height: p.height,
        zst: new Uint8Array(readFileSync(join(dir, `page-${n + 1}.rgba.zst`))),
      })
    );
    // A used entry survives `pruneRasterCache`.
    const now = new Date();
    utimesSync(join(dir, 'meta.json'), now, now);
    return pages;
  } catch {
    return undefined;
  }
}

/** Write a cache entry. `meta.json` comes last, through a rename, so a reader sees it whole. */
function writeCached(dir: string, pages: readonly Packed[]): void {
  try {
    mkdirSync(dir, { recursive: true });
    pages.forEach((p, n) => writeFileSync(join(dir, `page-${n + 1}.rgba.zst`), p.zst));
    const meta = join(dir, `meta.json.${process.pid}`);
    writeFileSync(
      meta,
      JSON.stringify({ pages: pages.map((p) => ({ width: p.width, height: p.height })) })
    );
    renameSync(meta, join(dir, 'meta.json'));
  } catch {
    // A cache that cannot be written only costs time.
  }
}

/**
 * Remove the cache entries that no run used since `since` (ms since the epoch). Only a run over
 * every fixture calls it, so the cache keeps the rasters of the current reference PDFs only.
 */
export function pruneRasterCache(cacheDir: string, since: number): void {
  let entries: string[];
  try {
    entries = readdirSync(cacheDir);
  } catch {
    return;
  }
  for (const entry of entries) {
    const dir = join(cacheDir, entry);
    let used = 0;
    try {
      used = statSync(join(dir, 'meta.json')).mtimeMs;
    } catch {
      // No meta.json: an entry that a stopped run left half written.
    }
    if (used < since) rmSync(dir, { recursive: true, force: true });
  }
}

/**
 * Rasterize the slide area of every page at `scale` times `widthPx` wide (`rasterOptions`), and
 * box-filter each page down to `widthPx` wide (`supersample.ts`) into `outDir` as a PNG. Returns
 * the pages in order. `pdftoppm` writes PPM, the same pixels as its PNG without the encoding.
 * With `cacheDir`, the device rasters are kept there, compressed, under `rasterCacheKey`, and a
 * later run reads them instead of running `pdftoppm`.
 */
export function rasterizePages(
  pdf: string,
  outDir: string,
  slide: SlideSize,
  widthPx: number,
  timeoutMs: number = POPPLER.timeout,
  scale: number = THRESHOLDS.rasterScale,
  cacheDir?: string
): ReferencePage[] {
  mkdirSync(outDir, { recursive: true });
  const options = rasterOptions(slide, widthPx, scale);
  const entry = cacheDir ? join(cacheDir, rasterCacheKey(pdf, options)) : undefined;
  let rasters = entry ? readCached(entry) : undefined;
  if (!rasters) {
    const tmp = join(outDir, 'ppm');
    rmSync(tmp, { recursive: true, force: true });
    mkdirSync(tmp, { recursive: true });
    try {
      execFileSync('pdftoppm', [...options, pdf, join(tmp, 'page')], {
        ...POPPLER,
        timeout: timeoutMs,
      });
      rasters = readdirSync(tmp)
        .filter((f) => /^page-\d+\.ppm$/.test(f))
        .sort((a, b) => Number(/\d+/.exec(a)![0]) - Number(/\d+/.exec(b)![0]))
        .map((f) => readPpm(readFileSync(join(tmp, f))));
    } finally {
      rmSync(tmp, { recursive: true, force: true });
    }
    if (entry && scale > 1) writeCached(entry, rasters.map(pack));
  }
  return rasters.map((device, n) => {
    const image = downsample(device, scale);
    const path = join(outDir, `page-${n + 1}.png`);
    writeFileSync(path, PNG.sync.write(image));
    // At scale 1 the device raster is the image; above it, it stays packed until it is needed.
    if (scale === 1) return { path, image, device: () => image };
    const packed = pack(device);
    return { path, image, device: () => unpack(packed) };
  });
}

/** `rasterizePages` without a cache: the paths of the filtered pages, in page order. */
export function rasterize(
  pdf: string,
  outDir: string,
  slide: SlideSize,
  widthPx: number,
  timeoutMs: number = POPPLER.timeout,
  scale: number = THRESHOLDS.rasterScale
): string[] {
  return rasterizePages(pdf, outDir, slide, widthPx, timeoutMs, scale).map((p) => p.path);
}

/** One text line in the PDF, in points from the page's top-left. */
export interface PdfLine {
  readonly text: string;
  readonly xMin: number;
  readonly xMax: number;
  readonly yMin: number;
  readonly yMax: number;
  /** The baseline from the PDF content stream (`withBaselines`), when a glyph origin was found. */
  readonly baseline?: number;
}

const ENTITIES: Readonly<Record<string, string>> = {
  amp: '&',
  lt: '<',
  gt: '>',
  quot: '"',
  '#39': "'",
};

/** Decode the five XML entities in one pass, so `&amp;lt;` stays `&lt;`. */
export const decode = (s: string): string =>
  s.replace(/&(amp|lt|gt|quot|#39);/g, (_, name: string) => ENTITIES[name]);

/** Text lines per page from `pdftotext -bbox-layout`. */
export function pdfLines(pdf: string, outDir: string): PdfLine[][] {
  mkdirSync(outDir, { recursive: true });
  const html = join(outDir, 'bbox.html');
  let text: string;
  try {
    execFileSync('pdftotext', ['-bbox-layout', pdf, html], POPPLER);
    text = readFileSync(html, 'utf8');
  } finally {
    // The file holds the slide text: remove it on every path, so an images folder keeps no text.
    rmSync(html, { force: true });
  }
  const pages: PdfLine[][] = [];
  for (const page of text.split(/<page /).slice(1)) {
    const lines: PdfLine[] = [];
    for (const m of page.matchAll(
      /<line xMin="([\d.]+)" yMin="([\d.]+)" xMax="([\d.]+)" yMax="([\d.]+)">([\s\S]*?)<\/line>/g
    )) {
      const words = [...m[5].matchAll(/<word[^>]*>([\s\S]*?)<\/word>/g)].map((w) => decode(w[1]));
      lines.push({
        text: words.join(' '),
        xMin: Number(m[1]),
        xMax: Number(m[3]),
        yMin: Number(m[2]),
        yMax: Number(m[4]),
      });
    }
    pages.push(lines);
  }
  return pages;
}

/**
 * The baseline of a PDF line, estimated from its box: 20 % of the box height above the bottom
 * (the descent share of typical Latin faces), or below the top for text turned 180°. Used only
 * when the PDF has no glyph origin for the line (`withBaselines`).
 */
export function estimatedBaseline(box: PdfLine, upsideDown = false): number {
  const descent = (box.yMax - box.yMin) * 0.2;
  return upsideDown ? box.yMin + descent : box.yMax - descent;
}

/** The start of a run of glyphs in the PDF content stream, in points from the page's top-left. */
export interface GlyphOrigin {
  /** The run's text, whitespace-normalized. */
  readonly text: string;
  readonly x: number;
  /** The baseline: the y of the text matrix origin. */
  readonly y: number;
  /** The text is turned 180°, so the run starts at the right. */
  readonly upsideDown: boolean;
}

/**
 * The origin of every text run on every page, from the PDF content stream (pdf.js). A
 * reference PDF holds each line as its own text object with its origin on a 1/300 in grid, so a
 * baseline read here is exact to ±0.12 pt. Runs turned by other angles are left out: the line
 * checks skip rotated text.
 */
export async function glyphOrigins(pdf: string): Promise<GlyphOrigin[][]> {
  const { getDocument } = await import('pdfjs-dist/legacy/build/pdf.mjs');
  const doc = await getDocument({
    data: new Uint8Array(readFileSync(pdf)),
    useSystemFonts: false,
    isEvalSupported: false,
    verbosity: 0,
  }).promise;
  try {
    const pages: GlyphOrigin[][] = [];
    for (let n = 1; n <= doc.numPages; n += 1) {
      const page = await doc.getPage(n);
      const viewport = page.getViewport({ scale: 1 });
      const origins: GlyphOrigin[] = [];
      for (const item of (await page.getTextContent()).items) {
        if (!('str' in item) || !item.str.trim()) continue;
        const [a, b, c, d, e, f] = item.transform as number[];
        if (Math.abs(b) > 1e-6 || Math.abs(c) > 1e-6 || a * d <= 0) continue;
        const [x, y] = viewport.convertToViewportPoint(e, f);
        origins.push({ text: normalize(item.str), x, y, upsideDown: d < 0 });
      }
      pages.push(origins);
      page.cleanup();
    }
    return pages;
  } finally {
    await doc.destroy();
  }
}

/** Origins this close (points) share one baseline: the 1/300 in grid of the reference PDF. */
const SAME_BASELINE_PT = 0.13;

/**
 * Give each `pdftotext` line the baseline of its glyph runs. The line's baseline is the one that
 * carries the most of the line's own text: the runs inside the line box whose text is part of
 * the line's text, grouped by baseline, weighted by their characters. Raised and lowered runs
 * (`a:rPr/@baseline`) are their own runs, so they do not move the line: in `font-effects`, the
 * lowered "2" of "Subscript H2O" makes the box taller, and its origin is nearer the box estimate
 * than the line start. Ties, and boxes with no run of their own text, go to the origin nearest
 * the baseline estimated from the box, then nearest the box's left edge (its right edge for text
 * turned 180°). Boxes of tightly spaced lines overlap, so a box can hold the origin of the line
 * above; that run's text is not part of the line. A line without an origin keeps no baseline;
 * the harness then uses the estimate.
 */
export function withBaselines(
  lines: readonly (readonly PdfLine[])[],
  origins: readonly (readonly GlyphOrigin[])[]
): PdfLine[][] {
  const bare = (s: string): string => s.replace(/\s+/g, '');
  return lines.map((page, i) =>
    page.map((line) => {
      const inside = (origins[i] ?? []).filter(
        (o) =>
          o.x >= line.xMin - 1 &&
          o.x <= line.xMax + 1 &&
          o.y >= line.yMin - 0.5 &&
          o.y <= line.yMax + 0.5
      );
      const text = bare(line.text);
      const weight = (y: number): number =>
        inside
          .filter((o) => Math.abs(o.y - y) <= SAME_BASELINE_PT && text.includes(bare(o.text)))
          .reduce((n, o) => n + bare(o.text).length, 0);
      let best: GlyphOrigin | undefined;
      let most = -1;
      let distance = Infinity;
      for (const o of inside) {
        const w = weight(o.y);
        const d =
          Math.abs(o.y - estimatedBaseline(line, o.upsideDown)) +
          0.01 * Math.abs(o.x - (o.upsideDown ? line.xMax : line.xMin));
        if (w > most || (w === most && d < distance)) {
          most = w;
          distance = d;
          best = o;
        }
      }
      return best ? { ...line, baseline: best.y } : line;
    })
  );
}

/** Whitespace-normalized text for line matching. */
export function normalize(s: string): string {
  return s.replace(/\s+/g, ' ').trim();
}

/**
 * The pages (1-based) of a PDF that hold JPEG or JPEG 2000 images (`pdfimages -list`). The
 * reference PDF stores pictures and rasterized effects as JPEG, so the pixels of such
 * a page include JPEG loss.
 */
export function jpegPages(pdf: string): Set<number> {
  const out = execFileSync('pdfimages', ['-list', pdf], { ...POPPLER, encoding: 'utf8' });
  const pages = new Set<number>();
  for (const line of out.split('\n').slice(2)) {
    const cols = line.trim().split(/\s+/);
    // `jpeg` is DCT, which reference PDFs hold; `jpx` (JPEG 2000) is lossy too.
    if (cols[8] === 'jpeg' || cols[8] === 'jpx') pages.add(Number(cols[0]));
  }
  return pages;
}

/** The fonts of one PDF page (1-based), without the subset prefix (`pdffonts`). */
export function pdfPageFonts(pdf: string, page: number): string[] {
  const out = execFileSync('pdffonts', ['-f', String(page), '-l', String(page), pdf], {
    ...POPPLER,
    encoding: 'utf8',
  });
  return out
    .split('\n')
    .slice(2)
    .map((line) => line.trim().split(/\s+/)[0] ?? '')
    .filter(Boolean)
    .map((name) => name.replace(/^[A-Z]{6}\+/, ''));
}
