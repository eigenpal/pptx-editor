// The package parts behind one diagram frame (ECMA-376 Part 1, 21.4 DrawingML Diagrams).
//
// A diagram `a:graphicData` holds `dgm:relIds`, whose `r:dm` names the data model part. The
// data model carries an extension, `dsp:dataModelExt/@relId`, that names the saved drawing
// part (`dsp:drawing`): the shapes last laid out, in the frame's coordinate space. That
// relationship is on the slide part, so it resolves there first, then on the data part.
//
// The reader keeps these parts as bytes, so a save writes them back unchanged. They are parsed
// here with the same bounded XML reader as every other part. Only internal parts of the
// expected relationship type resolve; an external target is never fetched. The approach
// follows DOCX Editor (EigenPal), `packages/core/src/store/package/diagram-parts.ts`.

import {
  attr,
  binaryPartOf,
  child,
  elements,
  NS,
  parseXmlPart,
  REL,
  relationshipsOf,
  xmlPartOf,
  type Element,
  type PptxPackage,
  type XmlPart,
} from '../store/index.ts';

const DIAGRAM_DATA =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData';

/** The most shapes one diagram drawing may hold. A larger drawing shows its placeholder. */
export const MAX_DIAGRAM_SHAPES = 512;
/** The deepest group nesting in one diagram drawing. */
export const MAX_DIAGRAM_GROUP_DEPTH = 8;

/** The parts one diagram renders from. */
export interface DiagramParts {
  /** The data model root (`dgm:dataModel`). */
  readonly dataModel: Element;
  readonly dataPart: string;
  /** The saved drawing root (`dsp:drawing`). */
  readonly drawing: Element;
  readonly drawingPart: string;
}

function decode(bytes: Uint8Array): string {
  if (bytes[0] === 0xff && bytes[1] === 0xfe) return new TextDecoder('utf-16le').decode(bytes);
  if (bytes[0] === 0xfe && bytes[1] === 0xff) return new TextDecoder('utf-16be').decode(bytes);
  const text = new TextDecoder('utf-8').decode(bytes);
  return text.charCodeAt(0) === 0xfeff ? text.slice(1) : text;
}

/** Parsed trees of parts held as bytes. Part bytes never change, so one parse serves them all. */
const parsed = new WeakMap<Uint8Array, XmlPart | null>();

/** The root of a part, parsed with the bounded reader when the package holds it as bytes. */
function partRoot(pkg: PptxPackage, partName: string | undefined): Element | undefined {
  const xml = xmlPartOf(pkg, partName);
  if (xml) return xml.root;
  const bytes = binaryPartOf(pkg, partName);
  if (!bytes || !partName) return undefined;
  let tree = parsed.get(bytes);
  if (tree === undefined) {
    const r = parseXmlPart(partName, decode(bytes));
    tree = r.ok ? r.part : null;
    parsed.set(bytes, tree);
  }
  return tree?.root;
}

/** The internal target of relationship `id` of `source`, only when its type is `type`. */
function target(
  pkg: PptxPackage,
  source: string | undefined,
  id: string,
  type: string
): string | undefined {
  if (!source) return undefined;
  const rel = relationshipsOf(pkg, source).find((r) => r.id === id);
  return rel && !rel.external && rel.type === type ? rel.target : undefined;
}

/** `dsp:dataModelExt/@relId` under the data model's `dgm:extLst/a:ext`. */
function drawingRelId(dataModel: Element): string | undefined {
  for (const ext of elements(child(dataModel, 'dgm:extLst'))) {
    const id = attr(child(ext, 'dsp:dataModelExt'), 'relId');
    if (id) return id;
  }
  return undefined;
}

/**
 * The data model and saved drawing behind a diagram frame's `a:graphicData`, or undefined
 * when either does not resolve to the expected root element.
 */
export function diagramParts(
  pkg: PptxPackage,
  slidePart: string,
  graphicData: Element | undefined
): DiagramParts | undefined {
  const dm = attr(child(graphicData, 'dgm:relIds'), 'r:dm');
  const dataPart = dm ? target(pkg, slidePart, dm, DIAGRAM_DATA) : undefined;
  const dataModel = partRoot(pkg, dataPart);
  if (!dataPart || dataModel?.ns !== NS.dgm || dataModel.local !== 'dataModel') return undefined;
  const id = drawingRelId(dataModel);
  if (!id) return undefined;
  const drawingPart =
    target(pkg, slidePart, id, REL.diagramDrawing) ?? target(pkg, dataPart, id, REL.diagramDrawing);
  const drawing = partRoot(pkg, drawingPart);
  if (!drawingPart || drawing?.ns !== NS.dsp || drawing.local !== 'drawing') return undefined;
  return { dataModel, dataPart, drawing, drawingPart };
}
