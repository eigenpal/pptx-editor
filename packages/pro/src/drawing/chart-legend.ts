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
   * The entry names a category of a series coloured point by point. In a row, the reference
   * leaves 0.825 sizes after such an entry, against 0.86 after a series (the `charts` PDF).
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

/** The key's width and the offset from the key to its text, for one entry. */
function keyMetrics(
  entry: LegendEntry,
  size: number
): { readonly key: number; readonly text: number; readonly after: number } {
  // A line entry whose line is hidden keys only its marker, as narrow as a square key.
  // Measured on the `charts` reference PDF: a line key is 1.6 sizes long, a square key 0.55
  // sizes wide, and a row leaves 0.92 or 0.86 sizes after each entry's text (0.825 after a
  // point entry).
  return entry.line && entry.look.line
    ? { key: 1.6 * size, text: 1.77 * size, after: 0.92 * size }
    : { key: 0.55 * size, text: 0.79 * size, after: (entry.point ? 0.825 : 0.86) * size };
}

/** Rows of a vertical legend are 1.4 lines apart (the `charts` reference PDF). */
const ROW_PITCH = 1.4;
/** A key's centre sits this many sizes above its entry's baseline (`paintLegend`). */
const KEY_RISE = 0.27;

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
    const blockHeight = line + (rows.length - 1) * pitch;
    const top = manual
      ? manual.y * height
      : legend.position === 'b'
        ? height - 9.6 - blockHeight
        : box.t - 2.9;
    for (const [rowIndex, row] of rows.entries()) {
      const total = row.reduce(
        (sum, item, index) => sum + item.width + (index > 0 ? row[index - 1]!.after : 0),
        0
      );
      let x = manual ? manual.x * width : width / 2 + 2.15 - total / 2;
      for (const item of row) {
        items.push({ entry: item.entry, x, baseline: top + rowIndex * pitch + baselineOffset });
        x += item.width + item.after;
      }
    }
    if (!legend.overlay && !manual) {
      if (legend.position === 'b') plot.b = Math.min(plot.b, top - 13.2);
      else plot.t = Math.max(plot.t, top + blockHeight + 14.6);
    }
    return { layout: { items }, plot };
  }
  // Vertical: one entry per row.
  const pitch = ROW_PITCH * line;
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
  const center = (box.t + height - 11) / 2;
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
    if (legend.position === 'l') plot.l = Math.max(plot.l, left + keyText + textWidth + 16.5);
    else plot.r = Math.min(plot.r, left - 16.5);
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
