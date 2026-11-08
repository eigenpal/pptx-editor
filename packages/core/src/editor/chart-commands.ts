// The chart commands on the editor's selection (`insertChart`, `setChartData`, `setChartType`,
// `setChartFormat`, `setChartColor`), and the selected chart as the snapshot reports it
// (`EditorSnapshot.chart`) for the chart controls and the Chart tab. The edits themselves are
// DOM-free (`automation/commands/chart-*.ts`).

import type {
  ChartCommand,
  ChartType,
  NewChartData,
  ChartDataRead,
  ChartFormatRead,
  EditorSelection,
  SelectedChartState,
} from '../contracts/index.ts';
import type { PptxPackage } from '../store/index.ts';
import { chartEditable, readChartData, setChartData } from '../automation/commands/chart-data.ts';
import {
  readChartFormat,
  setChartColor,
  setChartFormat,
  setChartType,
} from '../automation/commands/chart-format.ts';
import { insertChartFrame, sampleChartData } from '../automation/commands/chart-insert.ts';
import { isChartType } from '../automation/commands/chart-type.ts';
import { createTranslator } from '@pptx-editor.dev/i18n';
import { placeholderFor } from '../automation/commands/placeholder-fill.ts';
import { CommandRefusal, type OpBatch } from '../automation/commands/op-batch.ts';
import type { CommandContext, CommandEffect } from './commands.ts';
import { slidePart } from './text-targets.ts';

/** The one selected element, or undefined. */
function selectedOne(selection: EditorSelection): string | undefined {
  return selection.kind === 'shapes' && selection.nodeIds.length === 1
    ? selection.nodeIds[0]
    : undefined;
}

/** The selected chart, while exactly one chart frame is selected; else null. */
export function chartStateAt(
  pkg: PptxPackage | undefined,
  part: string | undefined,
  selection: EditorSelection
): SelectedChartState | null {
  const nodeId = selectedOne(selection);
  const state = pkg && part && nodeId ? chartEditable(pkg, part, nodeId) : undefined;
  if (!state || !nodeId) return null;
  return { nodeId, editable: state.editable, ...(state.reason ? { reason: state.reason } : {}) };
}

/** The data of the chart frame `nodeId`, or of the one selected chart frame. */
export function chartDataAt(
  pkg: PptxPackage | undefined,
  part: string | undefined,
  selection: EditorSelection,
  nodeId?: string
): ChartDataRead {
  const id = nodeId ?? selectedOne(selection);
  if (!pkg || !part || !id) return { ok: false, reason: 'errors.chartData.noChart' };
  return readChartData(pkg, part, id);
}

/** The type and elements of the chart frame `nodeId`, or of the one selected chart frame. */
export function chartFormatAt(
  pkg: PptxPackage | undefined,
  part: string | undefined,
  selection: EditorSelection,
  nodeId?: string
): ChartFormatRead {
  const id = nodeId ?? selectedOne(selection);
  if (!pkg || !part || !id) return { ok: false, reason: 'errors.chartData.noChart' };
  return readChartFormat(pkg, part, id);
}

const english = createTranslator();

/**
 * The sample data of a new chart, with the names in the host's language: `Category 1`,
 * `Series 1`, … (the Chart tab's names for a new category and series).
 */
function localSample(ctx: CommandContext, type: ChartType): NewChartData | undefined {
  const sample = sampleChartData(type);
  const name = (key: 'chartData.newCategory' | 'chartData.newSeries', n: number): string => {
    const text = ctx.translate?.(key);
    return text && text !== key
      ? text.replaceAll('{number}', String(n))
      : english(key, { number: n });
  };
  // A scatter chart's categories are X values: numbers, in any language.
  const numbered = sample.categories.every((c) => /^\d+$/.test(c));
  return {
    categories: numbered
      ? sample.categories
      : sample.categories.map((_, i) => name('chartData.newCategory', i + 1)),
    series: sample.series.map((s, i) => ({ ...s, name: name('chartData.newSeries', i + 1) })),
  };
}

/**
 * The facade's chart reads (`chartData`, `chartFormat`) over the editor's current package, slide
 * part, and selection.
 */
export function chartReads(
  state: () => readonly [PptxPackage | undefined, string | undefined, EditorSelection]
): {
  chartData(nodeId?: string): ChartDataRead;
  chartFormat(nodeId?: string): ChartFormatRead;
} {
  return {
    chartData: (id) => chartDataAt(...state(), id),
    chartFormat: (id) => chartFormatAt(...state(), id),
  };
}

/** The keys each chart command takes. */
const KEYS: Readonly<Record<ChartCommand['type'], ReadonlySet<string>>> = {
  setChartData: new Set(['type', 'nodeId', 'data']),
  insertChart: new Set(['type', 'chartType', 'data']),
  setChartType: new Set(['type', 'nodeId', 'chartType']),
  setChartFormat: new Set(['type', 'nodeId', 'format']),
  setChartColor: new Set(['type', 'nodeId', 'index', 'color']),
};

/** Whether `type` names a chart command. */
export const isChartCommandType = (type: string): type is ChartCommand['type'] =>
  Object.hasOwn(KEYS, type);

/** Compile a chart command. Throws CommandRefusal. */
export function compileChartCommand(
  ctx: CommandContext,
  batch: OpBatch,
  command: ChartCommand
): CommandEffect {
  const keys = KEYS[command.type];
  for (const key of Object.keys(command))
    if (!keys.has(key)) throw new CommandRefusal('errors.unsupported', key);
  const part = slidePart(ctx);
  if (command.type === 'insertChart') {
    if (!isChartType(command.chartType))
      throw new CommandRefusal('errors.unsupported', 'chartType');
    const { width, height } = ctx.session.info;
    const into = placeholderFor(batch, part, ctx.slide, ctx.selection, 'chart');
    const data = command.data ?? localSample(ctx, command.chartType);
    const id = insertChartFrame(batch, part, command.chartType, data, width, height, into);
    return { label: 'insertChart', selection: { kind: 'shapes', nodeIds: [id] } };
  }
  if (command.nodeId !== undefined && typeof command.nodeId !== 'string')
    throw new CommandRefusal('errors.unsupported', 'nodeId');
  const nodeId = command.nodeId ?? selectedOne(ctx.selection);
  if (!nodeId) throw new CommandRefusal('disabled.noSelection');
  switch (command.type) {
    case 'setChartData':
      setChartData(batch, part, nodeId, command.data);
      return { label: 'chartData' };
    case 'setChartType':
      setChartType(batch, part, nodeId, command.chartType);
      return { label: 'chartType' };
    case 'setChartFormat':
      setChartFormat(batch, part, nodeId, command.format);
      return { label: 'chartFormat' };
    default:
      setChartColor(batch, part, nodeId, command.index, command.color);
      return { label: 'chartColor' };
  }
}
