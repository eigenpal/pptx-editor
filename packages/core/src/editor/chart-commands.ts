// The chart data command (`setChartData`) on the editor's selection, and the selected chart as
// the snapshot reports it (`EditorSnapshot.chart`) for the chart controls and the Chart Data
// tab. The edit itself is DOM-free (`automation/commands/chart-data.ts`).

import type {
  ChartCommand,
  ChartDataRead,
  EditorSelection,
  SelectedChartState,
} from '../contracts/index.ts';
import type { PptxPackage } from '../store/index.ts';
import { chartEditable, readChartData, setChartData } from '../automation/commands/chart-data.ts';
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

const KEYS = new Set(['type', 'nodeId', 'data']);

/** Compile `setChartData`. Throws CommandRefusal. */
export function compileChartCommand(
  ctx: CommandContext,
  batch: OpBatch,
  command: ChartCommand
): CommandEffect {
  for (const key of Object.keys(command))
    if (!KEYS.has(key)) throw new CommandRefusal('errors.unsupported', key);
  if (command.nodeId !== undefined && typeof command.nodeId !== 'string')
    throw new CommandRefusal('errors.unsupported', 'nodeId');
  const nodeId = command.nodeId ?? selectedOne(ctx.selection);
  if (!nodeId) throw new CommandRefusal('disabled.noSelection');
  setChartData(batch, slidePart(ctx), nodeId, command.data);
  return { label: 'chartData' };
}
