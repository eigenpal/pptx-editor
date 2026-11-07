// The parts behind a diagram frame, opened for an edit, and written back. A diagram edit writes
// its data model part (`dgm:dataModel`) and its saved drawing part (`dsp:drawing`) again whole,
// as bytes, in the edit's transaction (`removePart` and `addPart`), as the reader keeps them: a
// collaboration peer and a reopened copy then hold the same parts. The text editor's ops need
// node ids, so a text edit runs on a scratch copy of the package that holds both parts as XML
// trees (`openDiagram`); the scratch ops are never committed.

import {
  attr,
  child,
  nodeIndex,
  serializeXmlPart,
  xmlPartOf,
  type Element,
  type PptxPackage,
} from '../../store/index.ts';
import { diagramParts, type DiagramParts } from '../../cascade/diagram.ts';
import { toBase64 } from './base64.ts';
import { CommandRefusal, OpBatch } from './op-batch.ts';

const DIAGRAM_URI = 'http://schemas.openxmlformats.org/drawingml/2006/diagram';

/** A diagram graphic frame and the parts it shows. */
export interface DiagramFrame extends DiagramParts {
  readonly frameId: string;
  readonly frame: Element;
}

/** The diagram behind the graphic frame `nodeId` of `slidePart`, or undefined. */
export function diagramFrameOf(
  pkg: PptxPackage,
  slidePart: string,
  nodeId: string
): DiagramFrame | undefined {
  const xml = xmlPartOf(pkg, slidePart);
  const node = xml ? nodeIndex(xml).get(nodeId)?.node : undefined;
  if (node?.type !== 'element' || node.kind !== 'graphicFrame') return undefined;
  const data = child(child(node, 'a:graphic'), 'a:graphicData');
  if (attr(data, 'uri') !== DIAGRAM_URI) return undefined;
  const parts = diagramParts(pkg, slidePart, data);
  return parts ? { ...parts, frameId: nodeId, frame: node } : undefined;
}

/** The diagram frame `nodeId`, or a refusal. */
export function requireDiagram(pkg: PptxPackage, slidePart: string, nodeId: string): DiagramFrame {
  const frame = diagramFrameOf(pkg, slidePart, nodeId);
  if (!frame) throw new CommandRefusal('errors.diagram.noDiagram');
  return frame;
}

/** Serialize a part root, or refuse an edit that leaves text XML cannot hold. */
function xmlOf(partName: string, root: Element): string {
  try {
    return serializeXmlPart({ partName, root, nextId: 0 });
  } catch {
    throw new CommandRefusal('errors.invalidText');
  }
}

/** Write a part again whole, as bytes, in this transaction. */
function replacePart(batch: OpBatch, part: string, root: Element): void {
  const bytes = new TextEncoder().encode(xmlOf(part, root));
  batch.push({ op: 'removePart', part });
  batch.push({ op: 'addPart', part, base64: toBase64(bytes) });
}

/** Write the data model and the saved drawing of a diagram back, each only when it changed. */
export function writeDiagram(
  batch: OpBatch,
  frame: DiagramFrame,
  dataModel: Element,
  drawing: Element
): void {
  if (dataModel !== frame.dataModel) replacePart(batch, frame.dataPart, dataModel);
  if (drawing !== frame.drawing) replacePart(batch, frame.drawingPart, drawing);
}

/** A scratch copy of the package in which both diagram parts are XML trees with node ids. */
export interface DiagramScratch {
  readonly batch: OpBatch;
  dataModel(): Element;
  drawing(): Element;
}

/** Open a diagram's parts on a scratch batch, for edits that need node ids. */
export function openDiagram(pkg: PptxPackage, frame: DiagramFrame): DiagramScratch {
  const scratch = new OpBatch(pkg);
  for (const [part, root] of [
    [frame.dataPart, frame.dataModel],
    [frame.drawingPart, frame.drawing],
  ] as const) {
    scratch.push({ op: 'removePart', part });
    scratch.push({ op: 'addPart', part, xml: xmlOf(part, root) });
  }
  const rootOf = (part: string): Element => {
    const root = xmlPartOf(scratch.pkg, part)?.root;
    if (!root) throw new CommandRefusal('errors.diagram.noDiagram');
    return root;
  };
  return {
    batch: scratch,
    dataModel: () => rootOf(frame.dataPart),
    drawing: () => rootOf(frame.drawingPart),
  };
}
