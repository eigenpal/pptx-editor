// Equation commands that `commands.ts` sends here: Insert Equation and Edit Equation
// (`contracts/equation.ts`). Each compiles to one transaction, so one undo step.

import type { EditorCommand, EquationCommand } from '../contracts/index.ts';
import {
  insertEquation,
  matchFallbackBox,
  setEquation,
} from '../automation/commands/equation-commands.ts';
import type { OpBatch } from '../automation/commands/op-batch.ts';
import type { CommandContext, CommandEffect } from './commands.ts';
import { refitTargets, slidePart } from './text-targets.ts';

export function isEquationCommand(command: EditorCommand): command is EquationCommand {
  return command.type === 'insertEquation' || command.type === 'setEquation';
}

export function compileEquationCommand(
  ctx: CommandContext,
  batch: OpBatch,
  command: EquationCommand
): CommandEffect {
  const part = slidePart(ctx);
  if (command.type === 'insertEquation') {
    const { width, height } = ctx.session.info;
    const id = insertEquation(batch, part, command.linear, width, height, command.font);
    refitTargets(batch, ctx, part, [{ shapeId: id }]);
    matchFallbackBox(batch, part, id);
    return { label: 'insertEquation', selection: { kind: 'shapes', nodeIds: [id] } };
  }
  // Fit the shape of the edited equation, which need not be selected.
  const shapeId = setEquation(batch, part, command.nodeId, command.linear);
  refitTargets(batch, ctx, part, [{ shapeId }]);
  return { label: 'setEquation' };
}
