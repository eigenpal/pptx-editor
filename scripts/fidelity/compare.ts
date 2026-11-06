// The per-slide measurements of the fidelity harness: the pixels of our PDF page against the
// reference page, and the text lines of both pages. Both sides come from the same code
// (`pdf.ts`), so every rule applies to both alike. Pure functions; `score-deck.ts` drives them.

import pixelmatch from 'pixelmatch';
import { PNG } from 'pngjs';
import { normalize, type PdfLine, type Raster } from './pdf.ts';
import { THRESHOLDS } from './thresholds.ts';

export interface PixelScore {
  /** Fraction of pixels that differ (pixelmatch, anti-aliasing tolerant). */
  readonly mismatch: number;
  /** The diff overlay as PNG, when asked for and some pixel differs. */
  readonly diff?: Buffer;
  /** The box around every differing pixel, `[x0, y0, x1, y1]` in pixels, or undefined. */
  readonly diffBox?: readonly [number, number, number, number];
}

/** Compare two rasters of the same size. */
export function comparePixels(ours: Raster, reference: Raster, withDiff = false): PixelScore {
  if (ours.width !== reference.width || ours.height !== reference.height)
    throw new Error('the two rasters have different sizes');
  const { width: w, height: h } = ours;
  const diff = new PNG({ width: w, height: h });
  const n = pixelmatch(ours.data, reference.data, diff.data, w, h, {
    threshold: THRESHOLDS.pixelThreshold,
  });
  return {
    mismatch: n / (w * h),
    diff: withDiff && n > 0 ? PNG.sync.write(diff, { filterType: 2 }) : undefined,
    diffBox: n ? differingBox(diff) : undefined,
  };
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

/**
 * A bullet at the start of a line: symbols, punctuation, and private-use characters. A symbol
 * font can map one bullet glyph to other characters in each PDF, so line matching leaves bullets
 * out on both sides; the pixels check the glyph.
 */
const LEADING_BULLET = /^[\p{S}\p{P}\p{Co}]+\s*/u;

/**
 * Right-to-left scripts. A PDF text layer of shaped Arabic or Hebrew text holds glyph forms or
 * guessed characters, not the text, so these lines are checked by pixels only.
 */
const RTL = /[֐-ࣿיִ-﷿ﹰ-﻿]/u;

/** Control characters: a glyph that the PDF text layer cannot map. */
const UNMAPPED = /[\u0000-\u0008\u000E-\u001F�]/u;

/**
 * The text that line matching compares: normalized, without a leading bullet or whitespace.
 * Empty for a line that is not checked: right-to-left text, unmapped glyphs, and lines without
 * two letters in a row (numbers, and the one glyph per line of vertical or stacked text).
 */
export function lineKey(text: string): string {
  const n = normalize(text);
  if (RTL.test(n) || UNMAPPED.test(n)) return '';
  return /\p{L}{2}/u.test(n) ? n.replace(LEADING_BULLET, '').replace(/\s+/g, '') : '';
}

/**
 * The key without a first one-character word, for a line that starts with a list marker
 * (`PdfLine.marker`): a bullet that a symbol font writes as a letter (`q Hollow square`). Empty
 * for any other line, so a one-letter word of the text must match exactly. Used only when the
 * full keys do not match.
 */
function bulletless(line: PdfLine): string {
  if (!line.marker) return '';
  const n = normalize(line.text).replace(LEADING_BULLET, '');
  return /^\S\s/u.test(n) ? lineKey(n.slice(2)) : '';
}

export interface LineScore {
  /** Reference lines that our page holds. */
  readonly matched: number;
  /** Reference lines that are checked (`lineKey`). */
  readonly total: number;
  /** Largest baseline distance of a matched line, in points. */
  readonly worstBaseline: number;
  /**
   * Largest horizontal distance of a matched line, in points: the nearer of the left-edge and
   * the right-edge distances.
   */
  readonly worstShift: number;
  /** Reference lines that our page does not hold (first 60 characters each). */
  readonly missing: readonly string[];
  /**
   * Our checked lines (`lineKey`) that the reference page does not hold (first 60 characters
   * each). Each one fails the slide (`THRESHOLDS.maxExtraLines`).
   */
  readonly extra: readonly string[];
}

/** Lines on one baseline within this distance (points) can join to hold a line of the other page. */
const SAME_BASELINE_PT = 0.25;

/**
 * Lines pair first only when their baselines, and their nearer edges, are this close (points):
 * 8 times the larger line tolerance of `THRESHOLDS`, so 4 pt. pdf.js lines carry no shape, so a
 * line of one shape could otherwise pair with a line of the same text in another shape, at any
 * distance. A line of the same text in another shape is at least one line pitch away (about 1.2
 * times the text size: 4 pt is less than half the pitch of 6.5 pt text), or in another column.
 * A pair past the tolerance but within this limit is one line out of place: it fails the slide
 * by its distance. Lines of the same text that no pair within the limit took then pair at any
 * distance, nearest first: a line that moved far fails the slide by its distance, and is not
 * missing and extra.
 */
export const LINE_PAIR_LIMIT_PT =
  8 * Math.max(THRESHOLDS.maxBaselineDeltaPt, THRESHOLDS.maxLineShiftPt);

/** One line of either page, as the line matching reads it. */
interface Entry {
  readonly line: PdfLine;
  readonly key: string;
  readonly alt: string;
  readonly bare: string;
  used: boolean;
}

const entry = (line: PdfLine): Entry => ({
  line,
  key: lineKey(line.text),
  alt: bulletless(line),
  bare: normalize(line.text).replace(/\s+/g, ''),
  used: false,
});

/**
 * A candidate pair: reference lines and our lines that hold the same text. One side is one line;
 * the other is one line, or several lines on one baseline that join to its text.
 */
interface Pair {
  readonly ref: readonly Entry[];
  readonly ours: readonly Entry[];
  readonly baseline: number;
  /** The nearer of the left-edge and the right-edge distances. */
  readonly shift: number;
  /** The distance that orders the pairs: of the left edges and the baselines. */
  readonly distance: number;
}

/** The pair of `ref` and `ours` (each in left-to-right order), or null past `limit`. */
function pairOf(
  ref: readonly Entry[],
  ours: readonly Entry[],
  limit = LINE_PAIR_LIMIT_PT
): Pair | null {
  const [r0, r1] = [ref[0].line, ref[ref.length - 1].line];
  const [o0, o1] = [ours[0].line, ours[ours.length - 1].line];
  const baseline = Math.abs(o0.baseline - r0.baseline);
  const shift = Math.min(Math.abs(o0.xMin - r0.xMin), Math.abs(o1.xMax - r1.xMax));
  if (baseline > limit || shift > limit) return null;
  return { ref, ours, baseline, shift, distance: Math.hypot(o0.xMin - r0.xMin, baseline) };
}

/**
 * Take the pairs nearest first, each line at most once. The order depends only on where the
 * lines are, not on which page holds a line or where it comes in its page.
 */
function assign(pairs: Pair[], take: (p: Pair) => void): void {
  const order = (p: Pair) => [p.distance, p.ref[0].line.baseline, p.ref[0].line.xMin] as const;
  pairs.sort((p, q) => {
    const [a, b] = [order(p), order(q)];
    return a[0] - b[0] || a[1] - b[1] || a[2] - b[2];
  });
  for (const p of pairs) {
    if (p.ref.some((e) => e.used) || p.ours.some((e) => e.used)) continue;
    for (const e of [...p.ref, ...p.ours]) e.used = true;
    take(p);
  }
}

/**
 * The unused lines of `pool` on the baseline of `line` that, joined in left-to-right order, hold
 * its text: a PDF can place each word of a justified line apart.
 */
function joinings(line: Entry, pool: readonly Entry[]): Entry[][] {
  const row = pool
    .filter((p) => !p.used && Math.abs(p.line.baseline - line.line.baseline) <= SAME_BASELINE_PT)
    .sort((p, q) => p.line.xMin - q.line.xMin);
  const out: Entry[][] = [];
  for (let i = 0; i < row.length; i += 1) {
    let joined = row[i].bare;
    for (let j = i + 1; j < row.length && joined.length < line.bare.length; j += 1) {
      joined += row[j].bare;
      if (joined === line.bare) out.push(row.slice(i, j + 1));
    }
  }
  return out;
}

/**
 * Compare the lines of our page with the lines of the reference page. Both pages follow the same
 * rules, so the pairs do not depend on which page is which. A line pairs with a line of the other
 * page with the same text (`lineKey`); failing that, the same text after a one-letter bullet;
 * failing that, several lines of the other page on its baseline that together hold its text.
 * Only pairs within `LINE_PAIR_LIMIT_PT` count, and the nearest are taken first, so repeated
 * text pairs within its shape, never with a line of another shape. Last, the lines of the same
 * text that are left pair at any distance, nearest first: a line that moved is out of place.
 */
export function compareLines(ours: readonly PdfLine[], reference: readonly PdfLine[]): LineScore {
  const mine = ours.map(entry);
  const refs = reference.map(entry);
  const checked = refs.filter((r) => r.key);
  let matched = 0;
  let worst = 0;
  let worstShift = 0;
  const take = (p: Pair): void => {
    matched += p.ref.filter((r) => r.key).length;
    worst = Math.max(worst, p.baseline);
    worstShift = Math.max(worstShift, p.shift);
  };
  /** The single-line pairs of unused lines that `same` accepts. */
  const pairs = (same: (r: Entry, o: Entry) => boolean, limit?: number): Pair[] => {
    const out: Pair[] = [];
    for (const r of checked)
      for (const o of mine) {
        if (r.used || o.used || !same(r, o)) continue;
        const p = pairOf([r], [o], limit);
        if (p) out.push(p);
      }
    return out;
  };
  // The same text.
  assign(
    pairs((r, o) => r.key === o.key),
    take
  );
  // A bullet that one PDF writes as a letter and the other as a symbol.
  assign(
    pairs(
      (r, o) => !!((r.alt && (o.alt === r.alt || o.key === r.alt)) || (o.alt && o.alt === r.key))
    ),
    take
  );
  // Several lines of either page that join to hold one line of the other.
  const joined: Pair[] = [];
  for (const r of checked)
    if (!r.used)
      for (const pieces of joinings(r, mine)) {
        const p = pairOf([r], pieces);
        if (p) joined.push(p);
      }
  for (const o of mine)
    if (!o.used && o.key)
      for (const pieces of joinings(o, refs)) {
        const p = pairOf(pieces, [o]);
        if (p) joined.push(p);
      }
  assign(joined, take);
  // The same text left over, at any distance: a line out of place, not missing and extra.
  assign(
    pairs((r, o) => r.key === o.key, Infinity),
    take
  );
  return {
    matched,
    total: checked.length,
    worstBaseline: worst,
    worstShift,
    missing: checked.filter((r) => !r.used).map((r) => normalize(r.line.text).slice(0, 60)),
    extra: mine.filter((o) => o.key && !o.used).map((o) => normalize(o.line.text).slice(0, 60)),
  };
}

/** Whether a slide passes every threshold in `thresholds.ts`. */
export function slidePasses(pixels: PixelScore, lines: LineScore): boolean {
  const lineMatch = lines.total ? lines.matched / lines.total : 1;
  return (
    pixels.mismatch <= THRESHOLDS.maxPixelMismatch &&
    lineMatch >= THRESHOLDS.minLineBreakMatch &&
    lines.worstBaseline <= THRESHOLDS.maxBaselineDeltaPt &&
    lines.worstShift <= THRESHOLDS.maxLineShiftPt &&
    lines.extra.length <= THRESHOLDS.maxExtraLines
  );
}
