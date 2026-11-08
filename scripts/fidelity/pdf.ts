// PDF helpers of the fidelity harness. The same code reads both PDFs of a comparison: the PDF that
// the editor prints and the reference PDF. pdf.js gives each PDF's page sizes, producer, and text
// lines here; the harness page draws the pages with the same pdf.js (`rasterPdf` in
// harness-page.ts). No system tool takes part. Test tools only; nothing here is bundled.

import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { getDocument, OPS, version as pdfjsVersion } from 'pdfjs-dist/legacy/build/pdf.mjs';
import type { PDFDocumentProxy } from 'pdfjs-dist/types/src/display/api';

/** The pdf.js version that draws and reads both PDFs. */
export const PDFJS_VERSION: string = pdfjsVersion;

/** The pdf.js package folder: its CMaps and standard font data. */
export const PDFJS_DIR = dirname(Bun.resolveSync('pdfjs-dist/package.json', import.meta.dir));

/**
 * pdf.js skips an image with more pixels than this. A PDF is untrusted input and can ask for any
 * image size; a skipped image shows as differing pixels.
 */
export const MAX_IMAGE_PIXELS = 2 ** 27;

/** A PDF with more pages than the caller allows (`openPdf`). */
export class PdfPageLimit extends Error {
  constructor(readonly pages: number) {
    super(`the PDF has ${pages} pages, more than allowed`);
  }
}

/** Load PDF bytes with the options that both PDFs of a comparison get. */
function load(bytes: Uint8Array): Promise<PDFDocumentProxy> {
  return getDocument({
    // pdf.js takes the buffer over: give it a copy.
    data: new Uint8Array(bytes),
    useSystemFonts: false,
    // pdf.js sends a font's `name` only with its extra properties (`fonts`).
    fontExtraProperties: true,
    maxImageSize: MAX_IMAGE_PIXELS,
    cMapUrl: `${join(PDFJS_DIR, 'cmaps')}/`,
    cMapPacked: true,
    // No standard font data: a font that the PDF does not embed stays `missingFile`, and the
    // harness refuses it (`fonts`), so no substitute font takes part.
    verbosity: 0,
  }).promise;
}

/** Pixels, RGBA, row by row. */
export interface Raster {
  readonly width: number;
  readonly height: number;
  readonly data: Uint8Array;
}

/** A raster size in pixels. */
export interface SizePx {
  readonly width: number;
  readonly height: number;
}

/** A slide size (`p:sldSz`) in EMU. */
export interface SlideSize {
  readonly widthEmu: number;
  readonly heightEmu: number;
}

/** The size of a slide raster at `widthPx` wide. */
export function rasterSizePx(slide: SlideSize, widthPx: number): SizePx {
  return { width: widthPx, height: Math.round((widthPx * slide.heightEmu) / slide.widthEmu) };
}

/** One text line of a PDF page, in points from the page's top-left. */
export interface PdfLine {
  /** The text, with one space where the PDF leaves a gap between two items. */
  readonly text: string;
  /** The left edge of the first item and the right edge of the last. */
  readonly xMin: number;
  readonly xMax: number;
  /** The baseline: the origin y of the item with the most characters. */
  readonly baseline: number;
  /**
   * The line starts with a one-character item in another font than the next item: a list
   * marker, such as a bullet that a symbol font writes as a letter.
   */
  readonly marker?: boolean;
}

/** A font that a PDF draws with: its name without a subset prefix, and whether it is embedded. */
export interface PdfFont {
  readonly name: string;
  readonly embedded: boolean;
}

/** An open PDF. Call `close` when done. */
export interface Pdf {
  /** The size of each page in points. */
  readonly pages: readonly (readonly [number, number])[];
  /** The PDF's `Creator` and `Producer`, joined with "; " (empty when absent). */
  readonly producer: string;
  /** The text lines of page `index` (0-based): unrotated text on the page only (`groupLines`). */
  lines(index: number): Promise<PdfLine[]>;
  /** Every font that the pages draw with, each name and state once, sorted by name. */
  fonts(): Promise<PdfFont[]>;
  close(): Promise<void>;
}

/**
 * Open PDF bytes with pdf.js. Throws when pdf.js cannot read them, and `PdfPageLimit` when the
 * PDF has more than `maxPages` pages (before any page is read).
 */
export async function openPdf(bytes: Uint8Array, maxPages = Infinity): Promise<Pdf> {
  const doc = await load(bytes);
  if (doc.numPages > maxPages) {
    await doc.loadingTask.destroy();
    throw new PdfPageLimit(doc.numPages);
  }
  const pages: (readonly [number, number])[] = [];
  for (let n = 1; n <= doc.numPages; n += 1) {
    const page = await doc.getPage(n);
    const { width, height } = page.getViewport({ scale: 1 });
    pages.push([width, height]);
  }
  const meta = (await doc.getMetadata().catch(() => undefined))?.info as
    { Creator?: string; Producer?: string } | undefined;
  const producer = [meta?.Creator, meta?.Producer]
    .filter((s): s is string => typeof s === 'string' && s.trim() !== '')
    .map((s) => s.trim())
    .join('; ');
  return {
    pages,
    producer,
    async lines(index) {
      const page = await doc.getPage(index + 1);
      const viewport = page.getViewport({ scale: 1 });
      const items: TextItem[] = [];
      for (const item of (await page.getTextContent()).items) {
        // Whitespace items carry no glyph; pdf.js also adds them to bridge gaps, which would join
        // columns. Spaces come from the gaps instead (`groupLines`).
        if (!('str' in item) || !item.str.trim()) continue;
        // An item with no advance is a mark that sits on another glyph, such as an Arabic dot
        // that a font draws apart from its letter. It holds no text of a line, and a text layer
        // maps it to any character: the pixels check it.
        if (!(item.width > 0)) continue;
        const [a, b, c, d, e, f] = item.transform as number[];
        // Rotated, mirrored, and vertical text is checked by pixels only. Text that is only
        // sheared (`c`) stays: its baseline is level. A synthetic oblique face shears its
        // glyphs, and pdf.js can merge upright text next to it into the same sheared item.
        if (Math.abs(b) > 1e-6 || a <= 0 || d <= 0 || Math.abs(c) >= d) continue;
        const [x, y] = viewport.convertToViewportPoint(e, f);
        items.push({ text: item.str, x, y, width: item.width, size: d, font: item.fontName });
      }
      page.cleanup();
      return groupLines(items, viewport.width, viewport.height);
    },
    async fonts() {
      const out = new Map<string, PdfFont>();
      for (let n = 1; n <= doc.numPages; n += 1) {
        const page = await doc.getPage(n);
        const ops = await page.getOperatorList();
        ops.fnArray.forEach((fn, i) => {
          if (fn !== OPS.setFont) return;
          const font = page.commonObjs.get(ops.argsArray[i][0]) as
            { name?: unknown; missingFile?: boolean } | undefined;
          const name = String(font?.name ?? 'unnamed').replace(/^[A-Z]{6}\+/, '');
          const embedded = font?.missingFile !== true;
          out.set(`${name}\0${embedded}`, { name, embedded });
        });
        page.cleanup();
      }
      return [...out.values()].sort((a, b) => a.name.localeCompare(b.name));
    },
    async close() {
      await doc.loadingTask.destroy();
    },
  };
}

/** Open a PDF file (`openPdf`). */
export const openPdfFile = (path: string, maxPages?: number): Promise<Pdf> =>
  openPdf(new Uint8Array(readFileSync(path)), maxPages);

/** One text item of a page: its text, origin, advance width, and font size, in points. */
export interface TextItem {
  readonly text: string;
  readonly x: number;
  /** The baseline, from the page's top. */
  readonly y: number;
  readonly width: number;
  readonly size: number;
  /** The PDF font of the item (pdf.js `fontName`). */
  readonly font?: string;
}

/** Items whose baselines are this close (points) share a row. */
const SAME_ROW_PT = 0.25;
/** A gap wider than this share of the font size starts a new line on the same row. */
const LINE_GAP_EM = 1;
/** A gap wider than this share of the font size is a space between two items. */
const SPACE_GAP_EM = 0.2;

/**
 * Group text items into lines. Items with an origin outside the page are left out. Items whose
 * baselines are within `SAME_ROW_PT` form a row; a row splits where the gap between two items is
 * wider than `LINE_GAP_EM` of the font size (columns, tabs, two shapes side by side). Both PDFs
 * of a comparison go through this one function, so the same layout gives the same lines.
 */
export function groupLines(
  items: readonly TextItem[],
  widthPt: number,
  heightPt: number
): PdfLine[] {
  const inside = items
    .filter((i) => i.x >= 0 && i.x <= widthPt && i.y >= 0 && i.y <= heightPt)
    .sort((p, q) => p.y - q.y || p.x - q.x);
  const rows: TextItem[][] = [];
  for (const item of inside) {
    const row = rows.at(-1);
    if (row && Math.abs(item.y - row[0].y) <= SAME_ROW_PT) row.push(item);
    else rows.push([item]);
  }
  const lines: PdfLine[] = [];
  for (const row of rows) {
    row.sort((p, q) => p.x - q.x);
    let run: TextItem[] = [];
    const flush = (): void => {
      if (run.length) lines.push(lineOf(run));
      run = [];
    };
    for (const item of row) {
      const last = run.at(-1);
      if (last && item.x - (last.x + last.width) > LINE_GAP_EM * Math.max(item.size, last.size))
        flush();
      run.push(item);
    }
    flush();
  }
  return lines.filter((l) => l.text.trim() !== '');
}

function lineOf(run: readonly TextItem[]): PdfLine {
  let text = '';
  let main = run[0];
  run.forEach((item, i) => {
    const last = run[i - 1];
    if (
      last &&
      item.x - (last.x + last.width) > SPACE_GAP_EM * Math.max(item.size, last.size) &&
      !/\s$/.test(text) &&
      !/^\s/.test(item.text)
    )
      text += ' ';
    text += item.text;
    if (item.text.trim().length > main.text.trim().length) main = item;
  });
  const end = run.at(-1)!;
  const marker = run.length > 1 && run[0].text.trim().length === 1 && run[0].font !== run[1].font;
  return {
    text: text.replace(/\s+/g, ' ').trim(),
    xMin: run[0].x,
    xMax: end.x + end.width,
    baseline: main.y,
    ...(marker ? { marker } : {}),
  };
}

/**
 * CJK radicals that a PDF text layer can write for an ideograph: a font maps both code points
 * to one glyph, and the PDF's `ToUnicode` map names the radical. The Kangxi radicals
 * (U+2F00 to U+2FD5) decompose under NFKC. These CJK Radicals Supplement characters do not,
 * so each maps to its equivalent unified ideograph (Unicode `EquivalentUnifiedIdeograph`).
 */
const RADICAL_IDEOGRAPHS: ReadonlyMap<string, string> = new Map([['⻑', '長']]);

/**
 * Whitespace-normalized text for line matching. Radicals that stand for an ideograph read as
 * the ideograph, so the same glyphs match on both sides.
 */
export function normalize(s: string): string {
  const text = /[⺀-⿟]/.test(s)
    ? s.replace(/[⺀-⿟]/g, (c) => RADICAL_IDEOGRAPHS.get(c) ?? c.normalize('NFKC'))
    : s;
  return text.replace(/\s+/g, ' ').trim();
}

/** The start of a run of glyphs in a PDF content stream, in points from the page's top-left. */
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
 * The origin of every text run on every page (pdf.js), for `calibrate.ts`. Runs turned by other
 * angles than 0° and 180° are left out.
 */
export async function glyphOrigins(pdf: string): Promise<GlyphOrigin[][]> {
  const doc = await load(new Uint8Array(readFileSync(pdf)));
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
    await doc.loadingTask.destroy();
  }
}
