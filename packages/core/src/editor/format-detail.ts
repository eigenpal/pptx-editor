import type { PatchShapeFormatCommand } from '../contracts/index.ts';
import type { CommandContext, CommandEffect } from './commands.ts';
import { slidePart } from './text-targets.ts';
import type { EditorSelection, ShapeFormatRead } from '../contracts/index.ts';
import type { DocumentSession } from './document-session.ts';
import { OpBatch, CommandRefusal } from '../automation/commands/op-batch.ts';
import {
  patchShapeFormatProperties,
  readShapeFormatProperties,
} from '../automation/commands/format-detail.ts';

/** Detailed properties of one selected shape, or an explicit node on the active part. @internal */
export function shapeFormatAt(
  session: DocumentSession | null,
  part: string | undefined,
  selection: EditorSelection,
  nodeId?: string
): ShapeFormatRead {
  const id =
    nodeId ??
    (selection.kind === 'shapes' && selection.nodeIds.length === 1
      ? selection.nodeIds[0]
      : selection.kind === 'text' && !selection.cell
        ? selection.nodeId
        : undefined);
  if (!session || !part || !id) return { ok: false, reason: 'disabled.noSelection' };
  try {
    return {
      ok: true,
      format: {
        target: { documentId: session.id, part, nodeId: id },
        ...readShapeFormatProperties(new OpBatch(session.store.package), part, id),
      },
    };
  } catch (error) {
    return {
      ok: false,
      reason: error instanceof CommandRefusal ? error.reason : 'errors.unsupported',
    };
  }
}

/** Compile a partial edit only for the document and active part that supplied its target. @internal */
export function compileShapeFormatPatch(
  ctx: CommandContext,
  batch: OpBatch,
  command: PatchShapeFormatCommand
): CommandEffect {
  const target = command.target;
  if (!target || target.documentId !== ctx.session.id || target.part !== slidePart(ctx))
    throw new CommandRefusal('errors.unsupported', 'stale format target');
  patchShapeFormatProperties(batch, target.part, target.nodeId, command.patch);
  return { label: 'shapeFormat' };
}
