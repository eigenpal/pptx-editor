// Text edits in a diagram shape. The text of a diagram node lives twice: in the data model point
// (`dgm:pt/dgm:t`), which a presentation application lays the diagram out from, and in the saved
// drawing shape (`dsp:sp/dsp:txBody`) that shows it, which the editor draws. One edit writes
// both, so the two never part. The edit runs the editor's own text operations on each body, so
// run formatting stays as it does in any other shape.

import {
  attr,
  child,
  children,
  elements,
  NS,
  specOf,
  type Element,
  type NodeSpec,
} from '../../store/index.ts';
import type { TextPoint } from '../../contracts/neutral.ts';
import { diagramTextNodes, readDiagramModel } from '../../cascade/diagram-model.ts';
import { openDiagram, requireDiagram, writeDiagram } from './diagram-edit.ts';
import { CommandRefusal, type OpBatch } from './op-batch.ts';
import { paragraphsOf, paragraphText } from './text-model.ts';

/** One text edit on a body of the scratch copy: the caret after it. */
export type DiagramTextEdit = (batch: OpBatch, part: string, txBodyId: string) => TextPoint;

/** The drawing shape (`dsp:sp`) under `tree` that shows the text of `node`. */
function shapeOf(
  tree: Element | undefined,
  nodes: ReadonlyMap<string, string>,
  node: string
): Element | undefined {
  for (const el of elements(tree)) {
    if (el.ns !== NS.dsp) continue;
    if (el.local === 'sp' && nodes.get(attr(el, 'modelId') ?? '') === node) return el;
    if (el.local === 'grpSp') {
      const found = shapeOf(el, nodes, node);
      if (found) return found;
    }
  }
  return undefined;
}

/** The plain text of a body, one entry per paragraph. */
function bodyText(body: Element | undefined): string {
  return body
    ? paragraphsOf(body)
        .map((p) => paragraphText(p))
        .join('\n')
    : '';
}

/** Attributes of run properties that the drawing's layout sets, not the author. */
const LAID_OUT = new Set(['sz', 'kern']);

/** A paragraph of the drawing as data model text: no laid-out paragraph or run sizes. */
function asModelParagraph(p: Element): NodeSpec {
  const strip = (spec: NodeSpec): NodeSpec => ({
    ...spec,
    ...(spec.local === 'rPr' || spec.local === 'endParaRPr'
      ? { attrs: (spec.attrs ?? []).filter((a) => a.ns !== '' || !LAID_OUT.has(a.local)) }
      : {}),
    children: (spec.children ?? [])
      .filter((c) => 'text' in c || c.local !== 'pPr')
      .map((c) => ('text' in c ? c : strip(c))),
  });
  return strip(specOf(p));
}

const a = (local: string, kids: NodeSpec[] = []): NodeSpec => ({
  ns: NS.a,
  local,
  prefix: 'a',
  children: kids,
});

/**
 * Give the data point the drawing body's text when the two did not hold the same text before
 * the edit, or the point has no text body (`dgm:t` is optional).
 */
function copyText(batch: OpBatch, part: string, point: Element, drawn: Element): void {
  const paragraphs = paragraphsOf(drawn).map(asModelParagraph);
  const t = child(point, 'dgm:t');
  if (t) {
    for (const p of children(t, 'a:p')) batch.push({ op: 'removeNode', part, node: p.id });
    const at = batch.element(part, t.id).children.length;
    paragraphs.forEach((spec, i) =>
      batch.push({ op: 'insertNode', part, parent: t.id, index: at + i, spec })
    );
    return;
  }
  // Schema order: prSet, spPr, t, extLst.
  const ext = point.children.findIndex((c) => c.type === 'element' && c.local === 'extLst');
  batch.push({
    op: 'insertNode',
    part,
    parent: point.id,
    index: ext < 0 ? point.children.length : ext,
    spec: {
      ns: NS.dgm,
      local: 't',
      prefix: 'dgm',
      children: [a('bodyPr'), a('lstStyle'), ...paragraphs],
    },
  });
}

/**
 * Edit the text of diagram node `node` of the diagram frame `frameId`: `edit` runs on the
 * drawing shape's body, then on the data point's body when both held the same text; otherwise
 * the point takes the shape's new text. Returns the caret in the drawing shape's body.
 */
export function editDiagramText(
  batch: OpBatch,
  slidePart: string,
  frameId: string,
  node: string,
  edit: DiagramTextEdit
): TextPoint {
  const frame = requireDiagram(batch.pkg, slidePart, frameId);
  const scratch = openDiagram(batch.pkg, frame);
  const model = readDiagramModel(scratch.dataModel());
  const shape = shapeOf(child(scratch.drawing(), 'dsp:spTree'), diagramTextNodes(model), node);
  const point = model.points.get(node)?.el;
  const drawnBody = child(shape, 'dsp:txBody');
  if (!shape || !point || !drawnBody) throw new CommandRefusal('errors.diagram.noText');
  const pointBody = child(point, 'dgm:t');
  const same = !!pointBody && bodyText(pointBody) === bodyText(drawnBody);
  const at = edit(scratch.batch, frame.drawingPart, drawnBody.id);
  if (same) edit(scratch.batch, frame.dataPart, pointBody.id);
  else {
    const drawn = scratch.batch.element(frame.drawingPart, drawnBody.id);
    copyText(scratch.batch, frame.dataPart, point, drawn);
  }
  writeDiagram(batch, frame, scratch.dataModel(), scratch.drawing());
  return at;
}
