// The diagram commands (`addDiagramShape`, `deleteDiagramShape`), typing in a diagram shape,
// and the selected diagram as the snapshot reports it (`EditorSnapshot.diagram`) for the
// Diagram Design controls. The edits themselves are DOM-free (`automation/commands/diagram-*`).

import type {
  DiagramCommand,
  EditorSelection,
  SelectedDiagramState,
  TextPoint,
} from '../contracts/index.ts';
import { child, intAttr, type Element, type PptxPackage } from '../store/index.ts';
import { readDiagramModel } from '../cascade/diagram-model.ts';
import {
  diagramFrameOf,
  requireDiagram,
  writeDiagram,
  type DiagramFrame,
} from '../automation/commands/diagram-edit.ts';
import { editDiagramText } from '../automation/commands/diagram-text.ts';
import { relayoutDrawing } from '../automation/commands/diagram-relayout.ts';
import { deleteDiagramNode, insertDiagramNode } from '../automation/commands/diagram-structure.ts';
import { CommandRefusal, type OpBatch } from '../automation/commands/op-batch.ts';
import type { CommandContext, CommandEffect } from './commands.ts';
import { editText, type TextInputKind } from './text-input.ts';
import type { ShapeItem } from './table-cells.ts';
import { slidePart } from './text-targets.ts';

type TextSelection = Extract<EditorSelection, { kind: 'text' }>;

/** Type into the diagram shape of a text selection (`diagramNode`); the caret after it. */
export function typeDiagram(
  batch: OpBatch,
  part: string,
  sel: TextSelection,
  item: ShapeItem,
  kind: TextInputKind,
  data: string | undefined
): TextPoint {
  return editDiagramText(batch, part, sel.nodeId, sel.diagramNode ?? '', (b, p, body) =>
    editText(b, p, body, item, sel, kind, data)
  );
}

/** The frame's extent, the drawing's space. */
function extentOf(frame: Element): { cx: number; cy: number } {
  const ext = child(child(frame, 'p:xfrm'), 'a:ext');
  return { cx: intAttr(ext, 'cx') ?? 0, cy: intAttr(ext, 'cy') ?? 0 };
}

/** Whether the drawing lays out again: a dry run of a structure edit that changes nothing. */
const structureCache = new WeakMap<Element, WeakMap<Element, string | null>>();
function structureRefusal(frame: DiagramFrame): string | null {
  let byData = structureCache.get(frame.drawing);
  if (!byData) structureCache.set(frame.drawing, (byData = new WeakMap()));
  const known = byData.get(frame.dataModel);
  if (known !== undefined) return known;
  let reason: string | null = null;
  try {
    const model = readDiagramModel(frame.dataModel);
    relayoutDrawing({
      drawing: frame.drawing,
      before: model,
      after: model,
      extent: extentOf(frame.frame),
    });
  } catch (e) {
    reason = e instanceof CommandRefusal ? e.reason : 'errors.diagram.layoutUnsupported';
  }
  byData.set(frame.dataModel, reason);
  return reason;
}

/** The diagram frame of a selection: one selected frame, or the frame that holds the caret. */
function frameIdOf(selection: EditorSelection): string | undefined {
  if (selection.kind === 'shapes' && selection.nodeIds.length === 1) return selection.nodeIds[0];
  if (selection.kind === 'text' && selection.diagramNode !== undefined) return selection.nodeId;
  return undefined;
}

/** The selected diagram, while one diagram frame is selected or holds the caret; else null. */
export function diagramStateAt(
  pkg: PptxPackage | undefined,
  part: string | undefined,
  selection: EditorSelection
): SelectedDiagramState | null {
  const id = frameIdOf(selection);
  const frame = pkg && part && id ? diagramFrameOf(pkg, part, id) : undefined;
  if (!frame || !id) return null;
  const reason = structureRefusal(frame);
  return {
    nodeId: id,
    node: selection.kind === 'text' ? (selection.diagramNode ?? null) : null,
    structure: reason === null,
    ...(reason ? { reason } : {}),
  };
}

const KEYS = new Set(['type', 'position', 'nodeId', 'node']);
const POSITIONS = new Set(['after', 'before', 'below']);

/** The target node of a command: its own, the caret's, or the first or last top-level node. */
function targetOf(
  frame: DiagramFrame,
  command: DiagramCommand,
  selection: EditorSelection
): string {
  if (command.node !== undefined) {
    if (typeof command.node !== 'string') throw new CommandRefusal('errors.unsupported', 'node');
    return command.node;
  }
  if (selection.kind === 'text' && selection.nodeId === frame.frameId && selection.diagramNode)
    return selection.diagramNode;
  const model = readDiagramModel(frame.dataModel);
  const top = model.root ? model.childrenOf(model.root) : [];
  const edge =
    command.type === 'addDiagramShape' && command.position === 'after'
      ? top[top.length - 1]
      : command.type === 'addDiagramShape' && command.position === 'before'
        ? top[0]
        : undefined;
  if (!edge) throw new CommandRefusal('errors.diagram.noNode');
  return edge;
}

/** Compile `addDiagramShape` or `deleteDiagramShape`. Throws CommandRefusal. */
export function compileDiagramCommand(
  ctx: CommandContext,
  batch: OpBatch,
  command: DiagramCommand
): CommandEffect {
  for (const key of Object.keys(command))
    if (!KEYS.has(key)) throw new CommandRefusal('errors.unsupported', key);
  if (command.nodeId !== undefined && typeof command.nodeId !== 'string')
    throw new CommandRefusal('errors.unsupported', 'nodeId');
  if (command.type === 'addDiagramShape' && !POSITIONS.has(command.position))
    throw new CommandRefusal('errors.unsupported', 'position');
  const nodeId = command.nodeId ?? frameIdOf(ctx.selection);
  if (!nodeId) throw new CommandRefusal('disabled.noSelection');
  const part = slidePart(ctx);
  const frame = requireDiagram(batch.pkg, part, nodeId);
  const target = targetOf(frame, command, ctx.selection);
  const before = readDiagramModel(frame.dataModel);
  const extent = extentOf(frame.frame);
  if (command.type === 'addDiagramShape') {
    const { dataModel, node } = insertDiagramNode(frame.dataModel, target, command.position);
    const after = readDiagramModel(dataModel);
    const drawing = relayoutDrawing({
      drawing: frame.drawing,
      before,
      after,
      extent,
      inserted: { node, target },
    });
    // A probe (`Editor.can`) only asks whether the edit runs: the parts are not written.
    if (!ctx.probe) writeDiagram(batch, frame, dataModel, drawing);
    const caret = { paragraph: 0, offset: 0 };
    return {
      label: 'addDiagramShape',
      selection: { kind: 'text', nodeId, diagramNode: node, anchor: caret, focus: caret },
    };
  }
  const { dataModel } = deleteDiagramNode(frame.dataModel, target);
  const after = readDiagramModel(dataModel);
  const drawing = relayoutDrawing({ drawing: frame.drawing, before, after, extent });
  if (!ctx.probe) writeDiagram(batch, frame, dataModel, drawing);
  return { label: 'deleteDiagramShape', selection: { kind: 'shapes', nodeIds: [nodeId] } };
}
