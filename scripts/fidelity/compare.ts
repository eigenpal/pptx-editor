// The per-slide measurements of the fidelity harness: pixels against the PDF raster, and laid-out
// lines against the PDF's text lines. Pure functions; `run.ts` drives them.

import pixelmatch from 'pixelmatch';
import { PNG } from 'pngjs';
import { estimatedBaseline, normalize, type PdfLine, type SizePx } from './references.ts';
import { deviceAntialiased, encodePng, pixels, type Raster } from './supersample.ts';
import { THRESHOLDS } from './thresholds.ts';
import type { HarnessLine } from './harness-page.ts';

export interface PixelScore {
  /**
   * Fraction of pixels that differ (pixelmatch, anti-aliasing tolerant). On a size mismatch it is
   * 1 and not a score: read `sizeMismatch` first (a slide then has no `pixelMismatch`).
   */
  readonly mismatch: number;
  readonly sizeMismatch: boolean;
  /** The size of each image, and the size that each must have when it was given. */
  readonly sizes: {
    readonly render: SizePx;
    readonly reference: SizePx;
    readonly expected?: ExpectedSizes;
  };
  /** The diff overlay, when asked for and some pixels differ. */
  readonly diff?: Buffer;
  /** The box around every differing pixel, `[x0, y0, x1, y1]` in pixels, or undefined. */
  readonly diffBox?: readonly [number, number, number, number];
}

/** The size each image must have: `renderSizePx` and `expectedRasterSizePx` (references.ts). */
export interface ExpectedSizes {
  readonly render: SizePx;
  readonly reference: SizePx;
}

export interface PixelOptions {
  /** Return the diff overlay. */
  readonly diff?: boolean;
  /** The size each image must have. Without it, the two images must have the same size. */
  readonly expected?: ExpectedSizes;
  /**
   * The same two images before the box filter, at `scale` device pixels for each pixel
   * (`supersample.ts`). With them, anti-aliasing is judged at the device scale
   * (`deviceAntialiased`), not on the filtered images.
   */
  readonly device?: {
    readonly render: Buffer | Raster;
    readonly reference: Buffer | Raster;
    readonly scale: number;
  };
}

/**
 * Compare two images, `render` and `reference`, each PNG bytes or decoded pixels. Each may differ from its expected size by
 * `maxSizeDeltaPx`; a larger difference is a size mismatch. The overlap of the two images is
 * compared. Pixels that differ only by anti-aliasing do not count: pixelmatch judges that on the
 * images themselves, or at the device scale when `options.device` holds the unfiltered images.
 */
export function comparePixels(
  render: Buffer | Raster,
  reference: Buffer | Raster,
  options: PixelOptions = {}
): PixelScore {
  const pa = pixels(render);
  const pb = pixels(reference);
  const off = (image: Raster, size: SizePx): boolean =>
    Math.abs(image.width - size.width) > THRESHOLDS.maxSizeDeltaPx ||
    Math.abs(image.height - size.height) > THRESHOLDS.maxSizeDeltaPx;
  const { expected } = options;
  const sizes = {
    render: { width: pa.width, height: pa.height },
    reference: { width: pb.width, height: pb.height },
    ...(expected ? { expected } : {}),
  };
  if (expected ? off(pa, expected.render) || off(pb, expected.reference) : off(pb, pa))
    return { mismatch: 1, sizeMismatch: true, sizes };
  const w = Math.min(pa.width, pb.width);
  const h = Math.min(pa.height, pb.height);
  const crop = (p: Raster): Uint8Array => {
    if (p.width === w && p.height === h) return p.data;
    const out = new Uint8Array(w * h * 4);
    for (let y = 0; y < h; y += 1)
      out.set(p.data.subarray(y * p.width * 4, y * p.width * 4 + w * 4), y * w * 4);
    return out;
  };
  const diff = new PNG({ width: w, height: h });
  const device = options.device && options.device.scale > 1 ? options.device : undefined;
  let n = pixelmatch(crop(pa), crop(pb), diff.data, w, h, {
    threshold: THRESHOLDS.pixelThreshold,
    // A box-filtered edge looks like anti-aliasing to pixelmatch at any position, so with device
    // images every differing pixel is found here and the device scale decides which count.
    includeAA: device !== undefined,
  });
  const differs = n > 0 || marked(diff);
  if (device && n) n = deviceAntialiased(diff, device.render, device.reference, device.scale);
  return {
    mismatch: n / (w * h),
    sizeMismatch: false,
    sizes,
    // An overlay shows counted (red) and excused (yellow) pixels. With neither, it is all gray.
    diff: options.diff && differs ? encodePng(diff) : undefined,
    diffBox: n ? differingBox(diff) : undefined,
  };
}

/**
 * Whether pixelmatch marked any pixel of its overlay: red for a difference, yellow for
 * anti-aliasing. It paints every other pixel gray, with equal channels.
 */
function marked(diff: PNG): boolean {
  const d = diff.data;
  for (let i = 0; i < d.length; i += 4) if (d[i] === 255 && d[i + 2] === 0) return true;
  return false;
}

/** Whether a size equals another. */
const same = (a: SizePx, b: SizePx): boolean => a.width === b.width && a.height === b.height;

/**
 * The cause of a size mismatch: a fixed text that says which image is wrong, with no numbers and
 * no ";", so that causes group. It starts with "slide size differs", which `eigenpal/pptx-eval`
 * matches, and it splits causes on ";" (`CONTRIBUTING.md`, "Fidelity harness"). The sizes are in
 * `sizeMismatchDetail` and in the slide's `renderSizePx` and `referenceSizePx`.
 */
export function sizeMismatchCause(sizes: PixelScore['sizes']): string {
  const { expected } = sizes;
  if (!expected) return 'slide size differs: the two images have different sizes';
  const render = !same(sizes.render, expected.render);
  const reference = !same(sizes.reference, expected.reference);
  if (render && reference)
    return 'slide size differs: neither the render nor the PDF raster has its expected size';
  return `slide size differs: the ${render ? 'render' : 'PDF raster'} does not have its expected size`;
}

/**
 * The sizes of a size mismatch, for the gate and the report, for example "the render is
 * 1920 × 1366 px, 1920 × 1358 px expected, and the PDF raster is 1920 × 1358 px, as expected".
 */
export function sizeMismatchDetail(sizes: PixelScore['sizes']): string {
  const px = (s: SizePx): string => `${s.width} × ${s.height} px`;
  const part = (name: string, actual: SizePx, want: SizePx | undefined): string =>
    `the ${name} is ${px(actual)}${
      !want ? '' : same(actual, want) ? ', as expected' : `, ${px(want)} expected`
    }`;
  return `${part('render', sizes.render, sizes.expected?.render)}, and ${part(
    'PDF raster',
    sizes.reference,
    sizes.expected?.reference
  )}`;
}

/** The size limit as the report and the gate write it: "exact", or "within N px". */
export function sizeLimitText(): string {
  const delta: number = THRESHOLDS.maxSizeDeltaPx;
  return delta ? `within ${delta} px` : 'exact';
}

/** The box around the pixels that pixelmatch marks as different (pure red in its output). */
function differingBox(diff: PNG): [number, number, number, number] | undefined {
  let x0 = Infinity;
  let y0 = Infinity;
  let x1 = -1;
  let y1 = -1;
  for (let y = 0; y < diff.height; y += 1)
    for (let x = 0; x < diff.width; x += 1) {
      const i = (y * diff.width + x) * 4;
      if (diff.data[i] === 255 && diff.data[i + 1] === 0 && diff.data[i + 2] === 0) {
        if (x < x0) x0 = x;
        if (x > x1) x1 = x;
        if (y < y0) y0 = y;
        if (y > y1) y1 = y;
      }
    }
  return x1 < 0 ? undefined : [x0, y0, x1 + 1, y1 + 1];
}

/** A font name without spaces, hyphens, or case, for a loose match of a family and a PDF name. */
const fontKey = (name: string): string => name.replace(/[^a-z0-9]/gi, '').toLowerCase();

/**
 * The requested families that a PDF page does not hold: the reference PDF drew that text with
 * another font. PDF names are PostScript names (`Calibri-Bold`, `ArialMT`), so a family matches a
 * name that starts with it.
 */
export function missingFonts(requested: readonly string[], present: readonly string[]): string[] {
  const keys = present.map(fontKey);
  return [...new Set(requested)].filter((f) => f && !keys.some((k) => k.startsWith(fontKey(f))));
}

/**
 * Whether a text is a bullet or a number that the reference PDF holds before a list line, with
 * spaces removed: one bullet or dash symbol (`•`, `–`, `▪`, or a private-use symbol-font
 * character), or a number, a letter, or a Roman numeral with `.` or `)` (`1.`, `a)`, `(iv)`). A
 * lost word (`The` before `quick brown fox`) is none of these.
 */
export const isListMarker = (prefix: string): boolean =>
  /^(?:[•◦▪▫■□●○◆◇►▸‣⁃∙·–—\-*✓✔➢➤→\uE000-\uF8FF]|\(?(?:\d{1,3}|[a-zA-Z]|[ivx]{1,5}|[IVX]{1,5})[.)])$/u.test(
    prefix
  );

export interface LineScore {
  readonly matched: number;
  readonly total: number;
  /** Largest baseline distance of a matched line, in points. */
  readonly worstBaseline: number;
  /** Matched lines whose PDF baseline is an estimate from the line box (no glyph origin). */
  readonly estimated: number;
  /** Our lines that start outside the slide and that no PDF line holds: not checked. */
  readonly skipped: number;
  readonly unmatched: readonly string[];
  /**
   * PDF lines with text that no laid-out line matched: text that we did not lay out, lay out
   * differently, or do not check (rotated and vertical text, tables). Reported, not gated.
   */
  readonly pdfUnmatched: number;
  /** The texts of those PDF lines (first 60 characters each), for the fixture report. */
  readonly pdfUnmatchedTexts: readonly string[];
  /**
   * Every unmatched laid-out line is the same text as unmatched PDF pieces on one row, joined in
   * x order with spaces removed: `pdftotext` split the line into other words (letter spacing,
   * sub- and superscripts, two font sizes), and the line did not break elsewhere.
   */
  readonly splitOnly: boolean;
  /**
   * Every unmatched laid-out line is the end of an unmatched PDF line whose extra start is a
   * bullet or a number (`isListMarker`). Read from the full texts.
   */
  readonly markersOnly: boolean;
  /** The unmatched lines of both sides hold the same characters: the text breaks elsewhere. */
  readonly sameText: boolean;
}

/** The baseline of a PDF line: its glyph origin, else the estimate from its box. */
export function pdfBaseline(box: PdfLine, upsideDown = false): number {
  return box.baseline ?? estimatedBaseline(box, upsideDown);
}

/** Text without any whitespace: `pdftotext` adds spaces inside letter-spaced and raised text. */
const bare = (s: string): string => s.replace(/\s+/g, '');

/**
 * Whether PDF text holds our line: the same characters, ignoring whitespace, after the bullet or
 * number when the line has one. A number must match exactly. A character bullet can map to
 * another character in the PDF (symbol fonts), so a prefix up to one character longer than the
 * bullet is accepted in its place; the bullet glyph itself is checked by pixels.
 */
function holds(pdfText: string, text: string, line: BulletOf): boolean {
  const p = bare(pdfText);
  const t = bare(text);
  if (p === t) return true;
  if (!p.endsWith(t)) return false;
  return isBullet(p.slice(0, p.length - t.length), line);
}

type BulletOf = Pick<HarnessLine, 'bullet' | 'numbered'>;

/** Whether PDF text (without whitespace) can stand for the line's bullet or number (`holds`). */
function isBullet(prefix: string, line: BulletOf): boolean {
  if (line.bullet === undefined) return false;
  return line.numbered
    ? prefix === bare(line.bullet)
    : prefix.length <= bare(line.bullet).length + 1;
}

/** Whether PDF text starts with `word`, after the line's bullet or number by the rule of `holds`. */
function startsWithWord(pdfText: string, word: string, line: BulletOf): boolean {
  const p = bare(pdfText);
  if (p.startsWith(word)) return true;
  if (line.bullet === undefined) return false;
  const longest = bare(line.bullet).length + 1;
  for (let k = 1; k <= longest && k < p.length; k += 1)
    if (p.startsWith(word, k) && isBullet(p.slice(0, k), line)) return true;
  return false;
}

/** PDF pieces on one baseline that `pdftotext` split, for example at a change of font size. */
const PIECE_BASELINE_PT = 1;
/** Joined pieces start at our line's left edge and leave no wider gap than this between them. */
const PIECE_GAP_PT = 12;

/**
 * Compare our lines with the PDF's lines. A line "matches" when the PDF has an unused line with
 * the same text (`holds`), or unused pieces on one baseline that together hold it. When several
 * PDF lines hold it, the nearest one is taken (by baseline and left edge), so repeated text on
 * one slide does not pair across shapes. The PDF baseline is the glyph origin from the content
 * stream (`pdfBaseline`); lines without one are counted in `estimated`, and the report labels
 * them.
 */
export function compareLines(ours: readonly HarnessLine[], pdf: readonly PdfLine[]): LineScore {
  let matched = 0;
  let worst = 0;
  let estimated = 0;
  let skipped = 0;
  const unmatched: string[] = [];
  const fullUnmatched: string[] = [];
  const pool = pdf.map((l) => ({ ...l, n: normalize(l.text), used: false }));
  type Entry = (typeof pool)[number];
  for (const line of ours) {
    const n = normalize(line.text);
    if (!n) continue;
    let hit: Entry[] | undefined;
    let nearest = Infinity;
    const consider = (pieces: Entry[]): void => {
      const dy = pdfBaseline(pieces[0], line.upsideDown) - line.baselinePt;
      // Text turned 180° starts at the right, so its left edges do not compare.
      const dx = line.upsideDown ? 0 : pieces[0].xMin - line.leftPt;
      const distance = Math.hypot(dx, dy);
      if (distance < nearest) {
        nearest = distance;
        hit = pieces;
      }
    };
    for (const p of pool) if (!p.used && holds(p.n, n, line)) consider([p]);
    if (!hit && !line.upsideDown) {
      const row = pool
        .filter((p) => !p.used && Math.abs(pdfBaseline(p) - line.baselinePt) <= PIECE_BASELINE_PT)
        .sort((a, b) => a.xMin - b.xMin);
      for (let i = 0; i < row.length; i += 1) {
        if (Math.abs(row[i].xMin - line.leftPt) > PIECE_GAP_PT) continue;
        for (let j = i + 1; j < row.length; j += 1) {
          if (row[j].xMin - row[j - 1].xMax > PIECE_GAP_PT) break;
          const pieces = row.slice(i, j + 1);
          if (holds(pieces.map((p) => p.n).join(' '), n, line)) consider(pieces);
        }
      }
    }
    if (!hit) {
      // A line off the slide is skipped only when no PDF line starts with its first word (after
      // its bullet): a different line break there still fails.
      const first = bare(n.split(' ')[0] ?? '');
      if (line.offSlide && !pool.some((p) => !p.used && startsWithWord(p.n, first, line)))
        skipped += 1;
      else {
        unmatched.push(n.slice(0, 60));
        fullUnmatched.push(n);
      }
      continue;
    }
    for (const p of hit) p.used = true;
    matched += 1;
    if (hit[0].baseline === undefined) estimated += 1;
    worst = Math.max(worst, Math.abs(pdfBaseline(hit[0], line.upsideDown) - line.baselinePt));
  }
  const rest = pool.filter((p) => p.n && !p.used);
  // Pieces on one row: their boxes overlap in y.
  const sameRow = (a: PdfLine, b: PdfLine): boolean =>
    Math.max(a.yMin, b.yMin) < Math.min(a.yMax, b.yMax);
  const split = (text: string): boolean =>
    rest.some((first) => {
      const row = rest.filter((p) => sameRow(first, p)).sort((a, b) => a.xMin - b.xMin);
      return row.some((_, i) => {
        let joined = '';
        for (let j = i; j < row.length && joined.length < text.length; j += 1)
          joined += bare(row[j].n);
        return joined === text;
      });
    });
  const restBare = rest.map((p) => bare(p.n));
  const marker = (text: string): boolean =>
    restBare.some(
      (p) => p.length > text.length && p.endsWith(text) && isListMarker(p.slice(0, -text.length))
    );
  const chars = (texts: readonly string[]): string => [...texts.join('')].sort().join('');
  return {
    splitOnly: fullUnmatched.length > 0 && fullUnmatched.every((t) => split(bare(t))),
    markersOnly: fullUnmatched.length > 0 && fullUnmatched.every((t) => marker(bare(t))),
    sameText: fullUnmatched.length > 0 && chars(fullUnmatched.map(bare)) === chars(restBare),
    matched,
    total: ours.filter((l) => normalize(l.text)).length - skipped,
    worstBaseline: worst,
    estimated,
    skipped,
    unmatched,
    pdfUnmatched: pool.filter((p) => p.n && !p.used).length,
    pdfUnmatchedTexts: pool.filter((p) => p.n && !p.used).map((p) => p.n.slice(0, 60)),
  };
}

/** Whether a slide passes every threshold in `thresholds.ts`. */
export function slidePasses(pixels: PixelScore, lines: LineScore): boolean {
  const lineMatch = lines.total ? lines.matched / lines.total : 1;
  return (
    !pixels.sizeMismatch &&
    pixels.mismatch <= THRESHOLDS.maxPixelMismatch &&
    lineMatch >= THRESHOLDS.minLineBreakMatch &&
    lines.worstBaseline <= THRESHOLDS.maxBaselineDeltaPt
  );
}
