/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Ported from DOCX Editor `packages/pro/src/drawing/chart-legend.ts` at commit ba742d7d1.
// The chart legend: its entries, the room it takes from the plot, and its painting.

import { categoryLabel } from './chart-number-format.ts';
import type { ChartLegend, ChartModel, ChartPlot } from './chart-model.ts';
import { ellipse, markerPath, markerStrokes, type ChartScene } from './chart-scene.ts';
import {
  isLineKind,
  seriesLook,
  variesByPoint,
  type LookContext,
  type SeriesLook,
} from './chart-series.ts';
import {
  baselineDrop,
  measureTextWidth,
  lineHeight,
  type ChartLabel,
  type ChartTextStyle,
} from './chart-text.ts';

const MAX_ENTRIES = 256;

export interface LegendEntry {
  readonly text: string;
  readonly look: SeriesLook;
  /** Line plots show a stroke with the marker; the others a filled square. */
  readonly line: boolean;
  /** Bubble plots fill a circle in the square's place. */
  readonly round?: boolean;
  /**
   * The entry names a category of a series coloured point by point. In a row, such an entry
   * leaves 0.035 sizes less after its text than a series entry (`keyMetrics`).
   */
  readonly point?: boolean;
}

/** Legend entries in the order the legend lists them, deleted entries removed. */
export function legendEntries(
  model: ChartModel,
  legend: ChartLegend,
  context: LookContext
): LegendEntry[] {
  const entries: LegendEntry[] = [];
  const vertical = legend.position === 'l' || legend.position === 'r' || legend.position === 'tr';
  for (const plot of model.plots) {
    const plotEntries: LegendEntry[] = [];
    const line = isLineKind(plot);
    // A lone series whose points are coloured one by one lists its categories.
    const lone = model.plots.length === 1 && plot.series.length === 1;
    const varied = variesByPoint(plot);
    if (varied || (lone && plot.series[0]!.points.size > 0 && !line)) {
      const series = plot.series[0]!;
      const count = series.values.length;
      for (let index = 0; index < count && index < MAX_ENTRIES; index += 1) {
        const text = categoryLabel(series, index, model.date1904) ?? String(index + 1);
        const point = varied || series.points.has(index) ? { index, count } : undefined;
        plotEntries.push({
          text,
          look: seriesLook(plot, series, context, point),
          line,
          point: true,
        });
      }
    } else {
      for (const series of plot.series) {
        if (plotEntries.length >= MAX_ENTRIES) break;
        plotEntries.push({
          text: series.name ?? `Series${series.index + 1}`,
          look: seriesLook(plot, series, context),
          line,
          round: plot.kind === 'bubble',
        });
      }
      if (reversesInLegend(plot, vertical)) plotEntries.reverse();
    }
    for (const entry of plotEntries) entries.push(entry);
  }
  return entries.filter((_, index) => !legend.deleted.has(index));
}

/** Series that stack or cluster upward list from the top in a vertical legend. */
function reversesInLegend(plot: ChartPlot, vertical: boolean): boolean {
  const stacked = plot.grouping === 'stacked' || plot.grouping === 'percentStacked';
  if (plot.kind === 'bar' && plot.horizontal) return plot.grouping === 'clustered';
  if (!vertical) return false;
  return stacked && (plot.kind === 'bar' || plot.kind === 'area' || plot.kind === 'line');
}

/** Geometry of the legend's entries, in chart points. */
export interface LegendLayout {
  readonly items: readonly {
    readonly entry: LegendEntry;
    readonly x: number;
    readonly baseline: number;
  }[];
}

/**
 * The key's width, the offset from the key to its text, and the room a row leaves after the
 * text, for one entry of `size` points. Measured on the `charts` and `chart-sizes` reference
 * PDFs at 9 to 20 pt:
 * - A square key is 0.55 sizes wide. Its text starts 0.275 sizes less 0.4 pt past the key (at
 *   the key's end for text under 1.45 pt). The key, its gap to the text, and the room after
 *   the text take 1.65 sizes together (1.615 for a point entry).
 * - A line key is 19.2 pt long at every size, with its text 2.03 pt past it. A row leaves
 *   6.43 pt and 0.386 sizes after the text.
 * - A row of entries centres 0.18 sizes right of the chart's centre (1.76 pt for line keys).
 * - A right legend keeps `gap` between the plot and its keys: 13.25 pt and half a square key,
 *   or 15.66 pt for line keys. A left legend keeps `LEFT_GAP` after its text.
 * The reference PDFs have no legend that mixes square and line keys. Such a legend takes the
 * `shift` of its first entry for every row, and the larger `gap`: both are unmeasured.
 * A line entry whose line is hidden keys only its marker, as narrow as a square key.
 */
export function keyMetrics(
  entry: LegendEntry,
  size: number
): {
  readonly key: number;
  readonly text: number;
  readonly after: number;
  readonly shift: number;
  readonly gap: number;
} {
  if (entry.line && entry.look.line) {
    const after = 6.43 + 0.386 * size;
    return { key: LINE_KEY, text: LINE_KEY + 2.03, after, shift: 1.76, gap: 15.66 };
  }
  const key = 0.55 * size;
  const text = key + Math.max(0, 0.275 * size - 0.4);
  const after = (entry.point ? 1.615 : 1.65) * size - text;
  return { key, text, after, shift: 0.18 * size, gap: 13.25 + key / 2 };
}

/** A line key's length in points, the same at every text size. */
const LINE_KEY = 19.2;

/** Rows of a vertical legend are one line and 5.85 pt apart, at every text size. */
const ROW_GAP = 5.85;
/** A key's centre sits this many sizes above its entry's baseline (`paintLegend`). */
const KEY_RISE = 0.27;
/**
 * A row legend centres its keys 9 pt and 0.6 sizes inside the chart edge (below the title for a
 * top legend), and the plot ends 14.2 pt and 0.6 sizes past that centre line.
 */
const ROW_EDGE = 9;
const ROW_HALF = 0.6;
const ROW_PLOT_GAP = 14.2;
/** The margin of an automatic chart layout on each side, in points (`chart-render.ts`). */
export const CHART_MARGIN = 11;
/**
 * The room after the text of a left legend. It is measured at 12 pt only: the reference PDFs
 * have no left legend at other sizes.
 */
const LEFT_GAP = 16.5;

export interface LegendRoom {
  readonly layout: LegendLayout;
  /** The plot's outer box after the legend took its room. */
  readonly plot: { l: number; t: number; r: number; b: number };
}

/**
 * Place the entries and shrink the plot box. Vertical legends sit against the chart edge,
 * centred on the room below the title; horizontal ones centre on the chart and wrap rows.
 */
export function placeLegend(
  legend: ChartLegend,
  entries: readonly LegendEntry[],
  style: ChartTextStyle,
  width: number,
  height: number,
  box: { l: number; t: number; r: number; b: number }
): LegendRoom {
  const size = style.sizePt;
  const line = lineHeight(style);
  const plot = { ...box };
  const items: { entry: LegendEntry; x: number; baseline: number }[] = [];
  const baselineOffset = baselineDrop(style);
  if (entries.length === 0) return { layout: { items }, plot };
  const manual = legend.layout;
  if (legend.position === 'b' || legend.position === 't') {
    const available = width - 20;
    const rows: { entry: LegendEntry; width: number; after: number }[][] = [[]];
    let rowWidth = 0;
    for (const entry of entries) {
      const metrics = keyMetrics(entry, size);
      const entryWidth = metrics.text + measureTextWidth(entry.text, style);
      const row = rows[rows.length - 1]!;
      const gap = row.length > 0 ? row[row.length - 1]!.after : 0;
      if (rowWidth + gap + entryWidth > available && row.length > 0) {
        rows.push([]);
        rowWidth = 0;
      } else rowWidth += gap;
      rowWidth += entryWidth;
      rows[rows.length - 1]!.push({ entry, width: entryWidth, after: metrics.after });
    }
    const pitch = 1.84 * size;
    const edge = ROW_EDGE + ROW_HALF * size;
    // The key centre line of the first row. A top legend measures from `CHART_MARGIN` above
    // `box.t`: the top of the chart, or 1.5 pt below the title's lines (`chart-render.ts`).
    const firstCenter =
      legend.position === 'b'
        ? height - edge - (rows.length - 1) * pitch
        : box.t - CHART_MARGIN + edge;
    const shift = keyMetrics(entries[0]!, size).shift;
    const lastCenter = firstCenter + (rows.length - 1) * pitch;
    const top = manual ? manual.y * height : null;
    for (const [rowIndex, row] of rows.entries()) {
      const total = row.reduce(
        (sum, item, index) => sum + item.width + (index > 0 ? row[index - 1]!.after : 0),
        0
      );
      let x = manual ? manual.x * width : width / 2 + shift - total / 2;
      const baseline =
        top === null
          ? firstCenter + rowIndex * pitch + KEY_RISE * size
          : top + rowIndex * pitch + baselineOffset;
      for (const item of row) {
        items.push({ entry: item.entry, x, baseline });
        x += item.width + item.after;
      }
    }
    if (!legend.overlay && !manual) {
      const reach = ROW_HALF * size + ROW_PLOT_GAP;
      if (legend.position === 'b') plot.b = Math.min(plot.b, firstCenter - reach);
      else plot.t = Math.max(plot.t, lastCenter + reach);
    }
    return { layout: { items }, plot };
  }
  // Vertical: one entry per row.
  const pitch = line + ROW_GAP;
  const textWidth = Math.max(...entries.map((entry) => measureTextWidth(entry.text, style)));
  const keyText = Math.max(...entries.map((entry) => keyMetrics(entry, size).text));
  // A placed legend shows the entries its box holds, centred in the box.
  const room = manual && manual.h > 0 ? manual.h * height : height - 20;
  const fits = Math.max(1, Math.floor((room - line) / pitch) + 1);
  const shown = entries.slice(0, fits);
  const blockHeight = line + (shown.length - 1) * pitch;
  const left = manual
    ? manual.x * width
    : legend.position === 'l'
      ? 10.7
      : width - 10 - textWidth - keyText;
  const center = (box.t + height - CHART_MARGIN) / 2;
  const top = manual
    ? manual.y * height + Math.max(0, (room - blockHeight) / 2)
    : legend.position === 'tr'
      ? box.t
      : center - blockHeight / 2;
  // An automatic legend centres its keys: the middle key sits on the centre line.
  const centred = !manual && legend.position !== 'tr';
  for (const [index, entry] of shown.entries()) {
    const baseline = centred
      ? center + (index - (shown.length - 1) / 2) * pitch + KEY_RISE * size
      : top + index * pitch + baselineOffset;
    items.push({ entry, x: left, baseline });
  }
  if (!legend.overlay && !manual) {
    if (legend.position === 'l') plot.l = Math.max(plot.l, left + keyText + textWidth + LEFT_GAP);
    else {
      const gap = Math.max(...entries.map((entry) => keyMetrics(entry, size).gap));
      plot.r = Math.min(plot.r, left - gap);
    }
  }
  return { layout: { items }, plot };
}

/** Paint the keys and queue the entry labels. */
export function paintLegend(
  layout: LegendLayout,
  style: ChartTextStyle,
  scene: ChartScene,
  labels: ChartLabel[]
): void {
  const size = style.sizePt;
  const line = lineHeight(style);
  for (const item of layout.items) {
    const { entry } = item;
    const metrics = keyMetrics(entry, size);
    const centerY = item.baseline - KEY_RISE * size;
    if (entry.line) {
      if (entry.look.line) {
        scene.line(
          [
            { x: item.x, y: centerY },
            { x: item.x + metrics.key, y: centerY },
          ],
          entry.look.line,
          'round'
        );
      }
      const marker = entry.look.marker;
      if (marker) {
        const cx = item.x + metrics.key / 2;
        scene.path(
          markerPath(marker.symbol, cx, centerY, marker.size * 0.96),
          true,
          marker.fill,
          marker.line
        );
        for (const stroke of markerStrokes(marker.symbol, cx, centerY, marker.size * 0.96)) {
          scene.line(stroke, marker.line);
        }
      }
    } else if (entry.round) {
      const half = metrics.key / 2;
      scene.path(
        [ellipse(item.x + half, centerY, half, half)],
        true,
        entry.look.fill,
        entry.look.line
      );
    } else {
      const half = metrics.key / 2;
      scene.rect(
        item.x,
        centerY - half,
        metrics.key,
        metrics.key,
        entry.look.fill,
        entry.look.line
      );
    }
    const textX = item.x + metrics.text;
    const estimate = measureTextWidth(entry.text, style);
    labels.push({
      paragraphs: [[{ text: entry.text, style }]],
      align: 'l',
      x: textX,
      y: item.baseline - baselineDrop(style),
      width: estimate * 1.3 + size * 2,
      height: line,
      rotation: 0,
    });
  }
}
