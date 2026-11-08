// The diagram commands (`addDiagramShape`, `deleteDiagramShape`, `setDiagramText`), typing in
// a diagram shape, and the selected diagram as the snapshot reports it (`EditorSnapshot.diagram`)
// for the Diagram Design controls and the Diagram tab of the side panel. The edits themselves are DOM-free (`automation/commands/diagram-*`).

import type {
  DiagramCommand,
  DiagramNodeState,
  EditorSelection,
  SelectedDiagramState,
  TextPoint,
} from '../contracts/index.ts';
import {
  attr,
  child,
  elements,
  intAttr,
  NS,
  type Element,
  type PptxPackage,
} from '../store/index.ts';
import {
  diagramTextNodes,
  isContentPoint,
  readDiagramModel,
  type DiagramModel,
} from '../cascade/diagram-model.ts';
import { cleanText } from '../automation/commands/text-units.ts';
import {
  deleteRange,
  insertText,
  paragraphsOf,
  paragraphText,
  splitParagraph,
} from '../automation/commands/text-model.ts';
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

/** The plain text of a text body, one line per paragraph. */
function textOf(body: Element | undefined): string {
  return body
    ? paragraphsOf(body)
        .map((p) => paragraphText(p))
        .join('\n')
    : '';
}

/** The text body of the drawing shape that shows `node` alone, if it has a paragraph. */
function shownBody(frame: DiagramFrame, node: string): Element | undefined {
  const textNodes = diagramTextNodes(readDiagramModel(frame.dataModel));
  for (const [modelId, shape] of drawnShapes(child(frame.drawing, 'dsp:spTree'))) {
    if (textNodes.get(modelId) !== node) continue;
    const body = child(shape, 'dsp:txBody');
    return body && paragraphsOf(body).length > 0 ? body : undefined;
  }
  return undefined;
}

/** The drawing shapes (`dsp:sp`) under `tree`, in groups too, by their `modelId`. */
function drawnShapes(
  tree: Element | undefined,
  out = new Map<string, Element>()
): Map<string, Element> {
  for (const el of elements(tree)) {
    if (el.ns !== NS.dsp) continue;
    const id = el.local === 'sp' ? attr(el, 'modelId') : undefined;
    if (id && !out.has(id)) out.set(id, el);
    if (el.local === 'grpSp') drawnShapes(el, out);
  }
  return out;
}

/** The content nodes of a model in outline order, with their text and whether it is editable. */
function nodesOf(model: DiagramModel, drawing: Element): readonly DiagramNodeState[] {
  // The node each drawing shape shows: a node that one shape shows alone takes typing.
  const shown = new Map<string, Element>();
  const textNodes = diagramTextNodes(model);
  for (const [modelId, shape] of drawnShapes(child(drawing, 'dsp:spTree'))) {
    const node = textNodes.get(modelId);
    if (node !== undefined && !shown.has(node)) shown.set(node, shape);
  }
  const out: DiagramNodeState[] = [];
  const seen = new Set<string>();
  const walk = (id: string, level: number): void => {
    for (const kid of model.childrenOf(id)) {
      // A model with a loop in its connections lists each node once.
      if (seen.has(kid) || out.length >= MAX_NODES) continue;
      seen.add(kid);
      const point = model.points.get(kid);
      if (!isContentPoint(point)) continue;
      const shape = shown.get(kid);
      const drawn = shape ? child(shape, 'dsp:txBody') : undefined;
      // A body with no paragraph has no point to type at.
      const body = drawn && paragraphsOf(drawn).length > 0 ? drawn : undefined;
      out.push({
        node: kid,
        text: body ? textOf(body) : textOf(child(point?.el, 'dgm:t')),
        level,
        editable: !!body,
      });
      walk(kid, level + 1);
    }
  };
  if (model.root) walk(model.root, 0);
  return out;
}

/** The most nodes the snapshot lists: a data model from a file has no other limit. */
const MAX_NODES = 1000;

const nodesCache = new WeakMap<Element, WeakMap<Element, readonly DiagramNodeState[]>>();
function cachedNodes(frame: DiagramFrame): readonly DiagramNodeState[] {
  let byData = nodesCache.get(frame.drawing);
  if (!byData) nodesCache.set(frame.drawing, (byData = new WeakMap()));
  let nodes = byData.get(frame.dataModel);
  if (!nodes)
    byData.set(
      frame.dataModel,
      (nodes = nodesOf(readDiagramModel(frame.dataModel), frame.drawing))
    );
  return nodes;
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
    nodes: cachedNodes(frame),
  };
}

/** The point at `offset` of the text `texts` makes, one newline between paragraphs. */
function pointAt(texts: readonly string[], offset: number): TextPoint {
  let rest = offset;
  for (let i = 0; i < texts.length - 1; i++) {
    const length = texts[i]!.length;
    if (rest <= length) return { paragraph: i, offset: rest };
    rest -= length + 1;
  }
  return { paragraph: texts.length - 1, offset: rest };
}

/**
 * Replace the text of a body with `text`, a newline between paragraphs: the caret after the
 * new text. Only the changed middle is replaced, so runs, fields, line breaks, and paragraph
 * properties outside the change stay as typing keeps them. The new text goes in at the end of
 * the change before the old text goes, so it takes the formatting of the run there.
 */
function replaceBodyText(batch: OpBatch, part: string, bodyId: string, text: string): TextPoint {
  const paragraphs = paragraphsOf(batch.element(part, bodyId));
  if (paragraphs.length === 0) throw new CommandRefusal('errors.diagram.noText');
  const texts = paragraphs.map((p) => paragraphText(p));
  const old = texts.join('\n');
  let head = 0;
  const most = Math.min(old.length, text.length);
  while (head < most && old[head] === text[head]) head += 1;
  let tail = 0;
  while (tail < most - head && old[old.length - 1 - tail] === text[text.length - 1 - tail])
    tail += 1;
  const start = pointAt(texts, head);
  let at = pointAt(texts, old.length - tail);
  text
    .slice(head, text.length - tail)
    .split('\n')
    .forEach((line, i) => {
      if (i > 0) at = splitParagraph(batch, part, bodyId, at);
      if (line) at = insertText(batch, part, bodyId, at, line);
    });
  return old.length - tail > head
    ? deleteRange(batch, part, bodyId, start, pointAt(texts, old.length - tail))
    : at;
}

/** The most characters `setDiagramText` writes, as a text field holds. */
const MAX_TEXT = 10_000;

/** Compile `setDiagramText`. Throws CommandRefusal. */
function compileSetText(
  ctx: CommandContext,
  batch: OpBatch,
  command: Extract<DiagramCommand, { type: 'setDiagramText' }>
): CommandEffect {
  for (const key of Object.keys(command))
    if (!TEXT_KEYS.has(key)) throw new CommandRefusal('errors.unsupported', key);
  if (command.nodeId !== undefined && typeof command.nodeId !== 'string')
    throw new CommandRefusal('errors.unsupported', 'nodeId');
  if (typeof command.node !== 'string') throw new CommandRefusal('errors.unsupported', 'node');
  if (typeof command.text !== 'string') throw new CommandRefusal('errors.invalidText');
  if (command.text.length > MAX_TEXT) throw new CommandRefusal('errors.diagram.textTooLong');
  const nodeId = command.nodeId ?? frameIdOf(ctx.selection);
  if (!nodeId) throw new CommandRefusal('disabled.noSelection');
  const part = slidePart(ctx);
  const frame = requireDiagram(batch.pkg, part, nodeId);
  // Tabs and line breaks become spaces, as in a pasted single-line field; newlines stay.
  const lines = command.text
    .replace(/\r\n?/g, '\n')
    .split('\n')
    .map((line) => cleanText(line).replace(/[\t\u000B]/g, ' '));
  if (lines.length > 1000) throw new CommandRefusal('errors.diagram.textTooLong');
  if (ctx.probe) {
    // A probe only asks whether the edit runs: one shape shows the node's text, with a
    // paragraph to type in, as the edit itself checks.
    if (!shownBody(frame, command.node)) throw new CommandRefusal('errors.diagram.noText');
    return { label: 'setDiagramText' };
  }
  editDiagramText(batch, part, nodeId, command.node, (b, p, body) =>
    replaceBodyText(b, p, body, lines.join('\n'))
  );
  // A caret in the node's shape may now point past its text: the diagram is selected instead.
  // Another selection stays, so a field of the side panel that applies on blur keeps the click.
  const sel = ctx.selection;
  const inNode = sel.kind === 'text' && sel.nodeId === nodeId && sel.diagramNode === command.node;
  return inNode
    ? { label: 'setDiagramText', selection: { kind: 'shapes', nodeIds: [nodeId] } }
    : { label: 'setDiagramText' };
}

const TEXT_KEYS = new Set(['type', 'node', 'text', 'nodeId']);
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
  if (command.type === 'setDiagramText') return compileSetText(ctx, batch, command);
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
