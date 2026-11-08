// The commands that have their own modules for one kind of slide content: comments, animations,
// charts, and diagrams. `compile` in `commands.ts` asks here first.

import type {
  ChartCommand,
  CommentCommand,
  DiagramCommand,
  EditorCommand,
} from '../contracts/index.ts';
import type { OpBatch } from '../automation/commands/op-batch.ts';
import {
  compileAnimationCommand,
  isAnimationCommand,
  type AnimationEditCommand,
} from './animation-command-table.ts';
import { compileChartCommand, isChartCommandType } from './chart-commands.ts';
import { compileCommentCommand, isCommentCommand } from './comment-commands.ts';
import type { CommandContext, CommandEffect } from './commands.ts';
import { compileDiagramCommand } from './diagram-commands.ts';

/** A command that one of these modules owns. */
export type ObjectCommand = CommentCommand | AnimationEditCommand | ChartCommand | DiagramCommand;

/** Whether one of these modules owns `command`. */
export function isObjectCommand(command: EditorCommand): command is ObjectCommand {
  return (
    isCommentCommand(command) ||
    isAnimationCommand(command) ||
    isChartCommandType(command.type) ||
    command.type === 'addDiagramShape' ||
    command.type === 'deleteDiagramShape' ||
    command.type === 'setDiagramText'
  );
}

/** The effect of `command`, from the module that owns it. */
export function compileObjectCommand(
  ctx: CommandContext,
  batch: OpBatch,
  command: ObjectCommand
): CommandEffect {
  if (isCommentCommand(command)) return compileCommentCommand(ctx, batch, command);
  if (isAnimationCommand(command)) return compileAnimationCommand(ctx, batch, command);
  if (isChartCommandType(command.type))
    return compileChartCommand(ctx, batch, command as ChartCommand);
  return compileDiagramCommand(ctx, batch, command as DiagramCommand);
}
