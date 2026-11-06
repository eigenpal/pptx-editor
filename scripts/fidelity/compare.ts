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
  /** Reference lines that our page does not hold (first 60 characters each). */
  readonly missing: readonly string[];
  /**
   * Our checked lines (`lineKey`) that the reference page does not hold (first 60 characters
   * each). Each one fails the slide (`THRESHOLDS.maxExtraLines`).
   */
  readonly extra: readonly string[];
}

/** Our lines on one baseline within this distance (points) can join to hold a reference line. */
const SAME_BASELINE_PT = 0.25;

/**
 * Compare the lines of our page with the lines of the reference page. A reference line matches
 * an unused line of ours with the same text (`lineKey`); failing that, the same text after a
 * one-letter bullet; failing that, several of our lines on its baseline that together hold its
 * text (a PDF can place each word of a justified line apart). When several candidates hold it,
 * the nearest one (by baseline and left edge) is taken, so repeated text does not pair across
 * shapes.
 */
export function compareLines(ours: readonly PdfLine[], reference: readonly PdfLine[]): LineScore {
  const pool = ours.map((l) => ({
    ...l,
    key: lineKey(l.text),
    alt: bulletless(l),
    bare: normalize(l.text).replace(/\s+/g, ''),
    used: false,
  }));
  type Entry = (typeof pool)[number];
  let matched = 0;
  let total = 0;
  let worst = 0;
  const missing: string[] = [];
  for (const line of reference) {
    const key = lineKey(line.text);
    if (!key) continue;
    total += 1;
    let hit: Entry[] | undefined;
    let nearest = Infinity;
    const consider = (pieces: Entry[]): void => {
      const d = Math.hypot(pieces[0].xMin - line.xMin, pieces[0].baseline - line.baseline);
      if (d < nearest) {
        nearest = d;
        hit = pieces;
      }
    };
    for (const p of pool) if (!p.used && p.key === key) consider([p]);
    // A bullet that one PDF writes as a letter and the other as a symbol.
    const alt = hit ? '' : bulletless(line);
    if (!hit)
      for (const p of pool)
        if (!p.used && ((alt && (p.alt === alt || p.key === alt)) || (p.alt && p.alt === key)))
          consider([p]);
    if (!hit) {
      const bare = normalize(line.text).replace(/\s+/g, '');
      const row = pool
        .filter((p) => !p.used && Math.abs(p.baseline - line.baseline) <= SAME_BASELINE_PT)
        .sort((p, q) => p.xMin - q.xMin);
      for (let i = 0; i < row.length; i += 1) {
        let joined = row[i].bare;
        for (let j = i + 1; j < row.length && joined.length < bare.length; j += 1) {
          joined += row[j].bare;
          if (joined === bare) consider(row.slice(i, j + 1));
        }
      }
    }
    if (!hit) {
      missing.push(normalize(line.text).slice(0, 60));
      continue;
    }
    for (const p of hit) p.used = true;
    matched += 1;
    worst = Math.max(worst, Math.abs(hit[0].baseline - line.baseline));
  }
  return {
    matched,
    total,
    worstBaseline: worst,
    missing,
    extra: pool.filter((p) => p.key && !p.used).map((p) => normalize(p.text).slice(0, 60)),
  };
}

/** Whether a slide passes every threshold in `thresholds.ts`. */
export function slidePasses(pixels: PixelScore, lines: LineScore): boolean {
  const lineMatch = lines.total ? lines.matched / lines.total : 1;
  return (
    pixels.mismatch <= THRESHOLDS.maxPixelMismatch &&
    lineMatch >= THRESHOLDS.minLineBreakMatch &&
    lines.worstBaseline <= THRESHOLDS.maxBaselineDeltaPt &&
    lines.extra.length <= THRESHOLDS.maxExtraLines
  );
}
