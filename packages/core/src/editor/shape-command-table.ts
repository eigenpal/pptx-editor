// Shape commands that `commands.ts` sends here: the adjust values of a preset shape, which its
// adjust handles drag (`surface-adjust.ts`), the pattern and picture fills, and the text warp.
// The equation commands go on to `equation-command-table.ts`. Each compiles to one transaction,
// so one undo step.

import type { EditorCommand } from '../contracts/index.ts';
import { child, type Element } from '../store/index.ts';
import { presetAdjustments } from '../layout/index.ts';
import { findItem } from '../automation/commands/placement.ts';
import { writeAdjustments } from '../automation/commands/adjust-commands.ts';
import { setShapePattern, setShapePicture } from '../automation/commands/style-commands.ts';
import { setTextWarp } from '../automation/commands/text-warp.ts';
import { CommandRefusal, type OpBatch } from '../automation/commands/op-batch.ts';
import type { CommandContext, CommandEffect } from './commands.ts';
import { selectedIds } from './commands.ts';
import { slidePart } from './text-targets.ts';
import { compileEquationCommand, isEquationCommand } from './equation-command-table.ts';

/** The commands this table compiles. */
export type ShapeCommand = Extract<
  EditorCommand,
  {
    readonly type:
      | 'setAdjustments'
      | 'setShapePattern'
      | 'setShapePicture'
      | 'setTextWarp'
      | 'insertEquation'
      | 'setEquation';
  }
>;

const TYPES: ReadonlySet<string> = new Set([
  'setAdjustments',
  'setShapePattern',
  'setShapePicture',
  'setTextWarp',
  'insertEquation',
  'setEquation',
]);

export function isShapeCommand(command: EditorCommand): command is ShapeCommand {
  return TYPES.has(command.type);
}

export function compileShapeCommand(
  ctx: CommandContext,
  batch: OpBatch,
  command: ShapeCommand
): CommandEffect {
  if (isEquationCommand(command)) return compileEquationCommand(ctx, batch, command);
  if (command.type === 'setShapePattern') {
    setShapePattern(batch, slidePart(ctx), selectedIds(ctx.selection), command.pattern);
    return { label: 'shapeFill' };
  }
  if (command.type === 'setShapePicture') {
    const ids = selectedIds(ctx.selection);
    setShapePicture(batch, slidePart(ctx), ids, command.bytes, command.tile === true);
    return { label: 'shapeFill' };
  }
  if (command.type === 'setTextWarp') {
    const items = ctx.session.layout(ctx.slide)?.items ?? [];
    const sizeOf = (id: string) => findItem(items, id)?.element.xfrm;
    const ids = selectedIds(ctx.selection);
    setTextWarp(batch, slidePart(ctx), ids, command.preset, command.adjust, sizeOf);
    return { label: 'textWarp' };
  }
  const ids = selectedIds(ctx.selection);
  if (ids.length !== 1) throw new CommandRefusal('disabled.selectOne');
  const id = ids[0]!;
  const item = findItem(ctx.session.layout(ctx.slide)?.items ?? [], id);
  const el = item?.element;
  if (!el || (el.kind !== 'shape' && el.kind !== 'connector' && el.kind !== 'picture'))
    throw new CommandRefusal('disabled.noAdjustHandles');
  if (el.geometry.kind !== 'preset') throw new CommandRefusal('disabled.noAdjustHandles');
  const { preset, adjust } = el.geometry;
  const slots = presetAdjustments(preset, el.xfrm.cx, el.xfrm.cy, adjust);
  if (!slots || slots.length === 0) throw new CommandRefusal('disabled.noAdjustHandles');
  const values = new Map<string, number>();
  for (const [name, raw] of Object.entries(command.values)) {
    // The names written are the preset table's, never the caller's text.
    const slot = slots.find((s) => s.name === name);
    if (!slot || typeof raw !== 'number' || !Number.isFinite(raw))
      throw new CommandRefusal('errors.unsupported', 'adjust value');
    let v = Math.round(raw);
    if (slot.min !== null && v < slot.min) v = Math.ceil(slot.min);
    if (slot.max !== null && v > slot.max) v = Math.floor(slot.max);
    values.set(slot.name, v);
  }
  if (values.size === 0) throw new CommandRefusal('disabled.noAdjustChange');
  const part = slidePart(ctx);
  const node: Element = batch.element(part, id);
  const spPr = child(node, 'p:spPr');
  if (!spPr || child(spPr, 'a:custGeom')) throw new CommandRefusal('disabled.noAdjustHandles');
  const shown = new Map(slots.map((s) => [s.name, Math.round(s.value)]));
  if ([...values].every(([name, v]) => shown.get(name) === v))
    throw new CommandRefusal('disabled.noAdjustChange');
  writeAdjustments(
    batch,
    { part, spPr, own: child(spPr, 'a:prstGeom'), preset, adjust: shown },
    values
  );
  return { label: 'adjustments' };
}
