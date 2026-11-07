// The shape clipboard (design D6 of `group-and-multi-edit`). Copy serializes the selected
// shapes and the parts their relationships point to; paste validates the payload as untrusted
// input, renumbers shape ids, and rewires relationships into the target slide.
//
// Pictures and media travel as bytes and are shared with an equal part on paste. A chart, a
// diagram, or an embedded object travels as its part graph (the part, the parts it owns, and
// their relationships), and paste gives it new parts, as a duplicated slide does. A web link
// travels as its address and passes `sanitizeHref` again on paste. A link to a slide is kept
// only in the presentation it came from. Other external relationships are never carried, so a
// pasted shape can never make the editor fetch anything.

import {
  isOnlyStructureMarks,
  isStructureMarkSpec,
  withoutStructureMarks,
} from '../automation/commands/table-edit-marks.ts';
import { asSavedForCopy } from '../automation/commands/table-repair.ts';
import {
  attr,
  contentTypeOf,
  elements,
  NS,
  nodeIndex,
  normalizePartName,
  presentationInfo,
  REL,
  relationshipsOf,
  relsPartName,
  sanitizeHref,
  serializeXmlPart,
  specOf,
  xmlPartOf,
  type Relationship,
  type Element,
  type NodeSpec,
  type PptxPackage,
} from '../store/index.ts';
import type { ClipboardNode, ClipboardPayload, ClipboardRelationship } from '../contracts/index.ts';
import { fromBase64, toBase64 } from '../automation/commands/base64.ts';
import { CommandRefusal, OpBatch } from '../automation/commands/op-batch.ts';
import { scopeItems } from '../automation/commands/placement.ts';
import type { LayoutItem } from '../layout/index.ts';
import {
  addExternalRelationship,
  addOverride,
  addRelationship,
  coordinate,
  ensureContentType,
  freePartName,
  relativeTarget,
  shapeIds,
  renumberShapes,
  spTreeOf,
  xmlAttr,
  XML_DECL,
} from '../automation/commands/package-edits.ts';
import { treeOf } from '../automation/commands/slide-structure.ts';
import { alternateBranch, isAlternateContent } from '../cascade/index.ts';

/** The clipboard MIME type for shapes. */
export const CLIPBOARD_MIME = 'application/x-pptx-editor+json';

const MAX_SHAPES = 500;
const MAX_NODES = 200_000;
const MAX_DEPTH = 64;
const MAX_BYTES = 64 * 1024 * 1024;
/**
 * The longest clipboard text parsed as a payload: the carried bytes in base64 plus room for the
 * shapes. Longer text is refused before `JSON.parse`, so a hostile page cannot freeze the tab.
 */
export const MAX_PAYLOAD_CHARS = Math.ceil(MAX_BYTES / 3) * 4 + 8 * 1024 * 1024;

/** Parse clipboard text in this editor's format; null when it is too long or not JSON. */
export function parseClipboardText(raw: string): unknown {
  if (!raw || raw.length > MAX_PAYLOAD_CHARS) return null;
  try {
    return JSON.parse(raw) as unknown;
  } catch {
    return null;
  }
}

/**
 * What a paste holds for this editor: shapes in its own format, the last copy when a browser
 * dropped that format but kept the matching (non-empty) text, a refusal for a payload that is
 * too long or not JSON, or null to leave the paste to the browser. `parse` caches by text.
 */
export function readClipboard(
  data: Pick<DataTransfer, 'getData'>,
  lastCopied: ClipboardPayload | null,
  parse: (raw: string) => unknown = parseClipboardText
): { readonly payload: unknown } | { readonly refused: true } | null {
  const raw = data.getData(CLIPBOARD_MIME);
  if (raw) {
    const payload = parse(raw);
    return payload ? { payload } : { refused: true };
  }
  // An empty text (an image from another app) never matches a copy without text.
  const text = data.getData('text/plain');
  return lastCopied && text && text === clipboardText(lastCopied) ? { payload: lastCopied } : null;
}
const NAME = /^[A-Za-z_][\w.-]*$/;
const SHAPES = new Set(['sp', 'pic', 'grpSp', 'cxnSp', 'graphicFrame']);
const HYPERLINKS = new Set(['hlinkClick', 'hlinkHover', 'hlinkMouseOver']);
const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';
const MS2011 = 'http://schemas.microsoft.com/office/2011/relationships/';
/** Relationship types whose binary targets travel with copied shapes. */
const CARRIED = new Set([
  `${R}image`,
  `${R}video`,
  `${R}audio`,
  'http://schemas.microsoft.com/office/2007/relationships/media',
]);

/** How a carried part graph is named on paste, and the content types it may have. */
interface GraphRule {
  readonly stem: string;
  readonly types: RegExp;
}
const DML = 'application\\/vnd\\.openxmlformats-officedocument\\.drawingml\\.';
const xmlType = (pattern: string): RegExp => new RegExp(`^${pattern}\\+xml$`);
/** An embedded package or OLE object: any `application/*` type, kept as bytes and never run. */
const EMBEDDED = /^application\/[A-Za-z0-9.+-]{1,200}$/;
/** The relationship types whose targets travel as part graphs (charts, diagrams, OLE). */
const GRAPHS: ReadonlyMap<string, GraphRule> = new Map([
  [`${R}chart`, { stem: '/ppt/charts/chart', types: xmlType(`${DML}chart`) }],
  [
    `${MS2011}chartStyle`,
    { stem: '/ppt/charts/style', types: /^application\/vnd\.ms-office\.chartstyle\+xml$/ },
  ],
  [
    `${MS2011}chartColorStyle`,
    { stem: '/ppt/charts/colors', types: /^application\/vnd\.ms-office\.chartcolorstyle\+xml$/ },
  ],
  [`${R}chartUserShapes`, { stem: '/ppt/drawings/drawing', types: xmlType(`${DML}chartshapes`) }],
  [
    `${R}themeOverride`,
    {
      stem: '/ppt/theme/themeOverride',
      types: xmlType('application\\/vnd\\.openxmlformats-officedocument\\.themeOverride'),
    },
  ],
  [`${R}diagramData`, { stem: '/ppt/diagrams/data', types: xmlType(`${DML}diagramData`) }],
  [`${R}diagramLayout`, { stem: '/ppt/diagrams/layout', types: xmlType(`${DML}diagramLayout`) }],
  [
    `${R}diagramQuickStyle`,
    { stem: '/ppt/diagrams/quickStyle', types: xmlType(`${DML}diagramStyle`) },
  ],
  [`${R}diagramColors`, { stem: '/ppt/diagrams/colors', types: xmlType(`${DML}diagramColors`) }],
  [
    REL.diagramDrawing,
    {
      stem: '/ppt/diagrams/drawing',
      types: /^application\/vnd\.ms-office\.drawingml\.diagramDrawing\+xml$/,
    },
  ],
  [`${R}oleObject`, { stem: '/ppt/embeddings/oleObject', types: EMBEDDED }],
  [`${R}package`, { stem: '/ppt/embeddings/package', types: EMBEDDED }],
]);
/** The most parts that one copy carries. */
const MAX_PARTS = 256;
/**
 * The namespace of the diagram data extension (`dsp:dataModelExt/@relId`) that names the saved
 * drawing of a diagram. The id is a relationship of the slide, not of the data part.
 */
const DSP = 'http://schemas.microsoft.com/office/drawing/2008/diagram';

/** Relationship ids a subtree uses (`r:*` attributes). */
function relIds(el: Element, out: Set<string>, depth = 0): Set<string> {
  if (depth > MAX_DEPTH) return out;
  for (const a of el.attrs) if (a.ns === NS.r) out.add(a.value);
  for (const c of el.children) if (c.type === 'element') relIds(c, out, depth + 1);
  return out;
}

/** The first element `ns`:`local` at or below `el`. */
function descendant(el: Element, ns: string, local: string, depth = 0): Element | undefined {
  if (el.ns === ns && el.local === local) return el;
  if (depth > MAX_DEPTH) return undefined;
  for (const c of el.children)
    if (c.type === 'element') {
      const hit = descendant(c, ns, local, depth + 1);
      if (hit) return hit;
    }
  return undefined;
}

/** The slide relationship that names the saved drawing of the diagram whose data is `part`. */
function diagramDrawingId(pkg: PptxPackage, part: string): string | undefined {
  const root = treeOf(pkg, part)?.root;
  const ext = root && descendant(root, DSP, 'dataModelExt');
  return ext ? attr(ext, 'relId') : undefined;
}

/** The bytes of a part, as a save writes them. */
function bytesOf(pkg: PptxPackage, name: string): Uint8Array | undefined {
  const part = pkg.parts.get(name);
  if (!part) return undefined;
  return part.kind === 'binary' ? part.bytes : new TextEncoder().encode(serializeXmlPart(part.xml));
}

/** The carried relationships of one copy, within the byte and part limits. */
class Carrier {
  bytes = 0;
  parts = 0;
  constructor(readonly pkg: PptxPackage) {}

  /**
   * A relationship as the clipboard carries it. Undefined drops it with the attributes that
   * name it; null means that the shape cannot travel without it.
   */
  carry(rel: Relationship, nested: boolean, depth = 0): ClipboardRelationship | null | undefined {
    const plain = { id: rel.id, type: rel.type, contentType: '', base64: '' };
    if (rel.type === REL.hyperlink) {
      const safe = rel.external && rel.externalTarget ? sanitizeHref(rel.externalTarget) : null;
      if (safe?.ok) return { ...plain, part: '', target: safe.href };
      return nested ? null : undefined;
    }
    if (rel.external || !rel.target) return nested ? null : undefined;
    if (rel.type === REL.slide && !nested) return { ...plain, part: rel.target };
    const graph = GRAPHS.get(rel.type);
    if (!CARRIED.has(rel.type) && !graph) return nested ? null : undefined;
    const part = this.pkg.parts.get(rel.target);
    if (!part || (!graph && part.kind !== 'binary')) return nested ? null : undefined;
    const contentType = contentTypeOf(this.pkg, rel.target) ?? 'application/octet-stream';
    if (graph && !graph.types.test(contentType)) return null;
    const bytes = bytesOf(this.pkg, rel.target)!;
    this.bytes += bytes.length;
    this.parts += 1;
    if (this.bytes > MAX_BYTES || this.parts > MAX_PARTS || depth > MAX_DEPTH) return null;
    const out = { ...plain, part: rel.target, contentType, base64: toBase64(bytes) };
    if (!graph) return out;
    const relationships: ClipboardRelationship[] = [];
    for (const r of relationshipsOf(this.pkg, rel.target)) {
      const c = this.carry(r, true, depth + 1);
      if (!c) return null;
      relationships.push(c);
    }
    return { ...out, relationships };
  }
}

const isBranchEl = (el: Element): boolean =>
  el.ns === NS.mc && (el.local === 'Choice' || el.local === 'Fallback');

/**
 * The `mc:AlternateContent` that holds `shape` as a copy in one of its branches, or `shape`
 * itself. A copy travels with every copy of the other branches (an equation keeps its math with
 * its picture), and the wrapper declares each namespace that a branch requires.
 */
function withAlternates(pkg: PptxPackage, slidePart: string, shape: Element): Element {
  const xml = xmlPartOf(pkg, slidePart);
  if (!xml) return shape;
  const index = nodeIndex(xml);
  const up = (id: string | undefined): Element | undefined => {
    const parent = id === undefined ? undefined : index.get(id)?.parent;
    const node = parent == null ? undefined : index.get(parent)?.node;
    return node?.type === 'element' ? node : undefined;
  };
  const branch = up(shape.id);
  const wrapper = branch && isBranchEl(branch) ? up(branch.id) : undefined;
  if (!wrapper || !isAlternateContent(wrapper)) return shape;
  // The namespaces in scope at the wrapper, nearest first.
  const scope = new Map<string, string>();
  for (let el: Element | undefined = wrapper; el; el = up(el.id))
    for (const d of el.nsDecls) if (!scope.has(d.prefix)) scope.set(d.prefix, d.uri);
  const own = new Set(wrapper.nsDecls.map((d) => d.prefix));
  const extra: { prefix: string; uri: string }[] = [];
  for (const b of elements(wrapper)) {
    if (!isBranchEl(b)) continue;
    const local = new Set(b.nsDecls.map((d) => d.prefix));
    for (const prefix of (attr(b, 'Requires') ?? '').split(/\s+/).filter(Boolean)) {
      const uri = scope.get(prefix);
      if (uri === undefined || local.has(prefix) || own.has(prefix)) continue;
      own.add(prefix);
      extra.push({ prefix, uri });
    }
  }
  return extra.length ? { ...wrapper, nsDecls: [...wrapper.nsDecls, ...extra] } : wrapper;
}

/**
 * Serialize the selected shapes of one slide in z-order (the order of their scope: the slide,
 * or the entered `group`). Null when nothing can be copied.
 */
export function copySelection(
  pkg: PptxPackage,
  slidePart: string,
  items: readonly LayoutItem[],
  selected: { readonly nodeIds: readonly string[]; readonly group?: string },
  document: string
): ClipboardPayload | null {
  const order = scopeItems(items, selected.group).map((i) => i.element.nodeId);
  const ids = [...selected.nodeIds].sort((a, b) => order.indexOf(a) - order.indexOf(b));
  const batch = new OpBatch(pkg);
  try {
    const shapes = new Map<string, Element>();
    for (const id of ids) {
      const el = withAlternates(pkg, slidePart, batch.element(slidePart, id));
      if (!shapes.has(el.id)) shapes.set(el.id, el);
    }
    return copyShapes(pkg, slidePart, [...shapes.values()], document);
  } catch (e) {
    if (e instanceof CommandRefusal) return null;
    throw e;
  }
}

/**
 * Serialize shapes of one slide. A shape whose chart, diagram, or embedded object cannot travel
 * (an external or unknown relationship in its part graph, or a part past the limits) is left
 * out; null when nothing is left.
 */
export function copyShapes(
  pkg: PptxPackage,
  slidePart: string,
  shapes: readonly Element[],
  document: string
): ClipboardPayload | null {
  const rels = new Map(relationshipsOf(pkg, slidePart).map((r) => [r.id, r]));
  const carried = new Map<string, ClipboardRelationship>();
  const kept: ClipboardNode[] = [];
  const carrier = new Carrier(pkg);
  for (const shape of shapes) {
    const pending: ClipboardRelationship[] = [];
    const ids = relIds(shape, new Set());
    // A diagram's saved drawing is a relationship of the slide that its data part names.
    for (const id of [...ids]) {
      const rel = rels.get(id);
      if (rel?.type === `${R}diagramData` && rel.target) {
        const drawing = diagramDrawingId(pkg, rel.target);
        if (drawing) ids.add(drawing);
      }
    }
    const before = { bytes: carrier.bytes, parts: carrier.parts };
    let ok = true;
    for (const id of ids) {
      const rel = rels.get(id);
      if (!rel || carried.has(id) || pending.some((p) => p.id === id)) continue;
      const c = carrier.carry(rel, false);
      if (c === null) ok = false;
      if (!c) continue;
      pending.push(c);
    }
    // A table that the save leaves out is not copied; a marked table goes as the canvas draws
    // it and a save writes it, without its marks.
    const saved = ok ? asSavedForCopy(pkg, slidePart, shape) : null;
    if (!saved) {
      carrier.bytes = before.bytes;
      carrier.parts = before.parts;
      continue;
    }
    for (const r of pending) carried.set(r.id, r);
    kept.push(withoutStructureMarks(specOf(saved)));
  }
  if (kept.length === 0) return null;
  return { version: 1, document, slidePart, shapes: kept, relationships: [...carried.values()] };
}

/** Plain text of copied shapes, one paragraph per line. */
export function clipboardText(payload: ClipboardPayload): string {
  const lines: string[] = [];
  const walk = (n: ClipboardNode, line: string[]): void => {
    for (const c of n.children ?? []) {
      if ('text' in c) {
        if (n.local === 't') line.push(c.text);
      } else if (c.local === 'p' && c.ns === NS.a) {
        const p: string[] = [];
        walk(c, p);
        lines.push(p.join(''));
      } else if (c.local === 'br' && c.ns === NS.a) line.push('\n');
      else walk(c, line);
    }
  };
  // One branch of an `mc:AlternateContent` gives its text: the first one.
  const drawn = (s: ClipboardNode): ClipboardNode =>
    s.ns === NS.mc && s.local === 'AlternateContent'
      ? { ...s, children: (s.children ?? []).filter((c) => !('text' in c)).slice(0, 1) }
      : s;
  for (const s of payload.shapes) walk(drawn(s), []);
  return lines.join('\n').trim();
}

// ---- Paste ----

const isObj = (v: unknown): v is Record<string, unknown> =>
  typeof v === 'object' && v !== null && !Array.isArray(v);

/** Validate an untrusted node: names, strings, depth, and size. */
export function sanitizeNode(v: unknown, depth: number, budget: { n: number }): NodeSpec | null {
  if (!isObj(v) || depth > MAX_DEPTH || (budget.n -= 1) < 0) return null;
  const { ns, local, prefix, attrs, nsDecls, children } = v;
  if (typeof ns !== 'string' || typeof local !== 'string' || !NAME.test(local)) return null;
  if (prefix !== undefined && (typeof prefix !== 'string' || !NAME.test(prefix))) return null;
  const outAttrs: { ns: string; local: string; prefix?: string; value: string }[] = [];
  if (attrs !== undefined) {
    if (!Array.isArray(attrs)) return null;
    for (const a of attrs) {
      if (!isObj(a)) return null;
      if (typeof a.ns !== 'string' || typeof a.local !== 'string' || !NAME.test(a.local))
        return null;
      if (typeof a.value !== 'string') return null;
      if (a.prefix !== undefined && (typeof a.prefix !== 'string' || !NAME.test(a.prefix)))
        return null;
      outAttrs.push({
        ns: a.ns,
        local: a.local,
        value: a.value,
        ...(a.prefix ? { prefix: a.prefix } : {}),
      });
    }
  }
  const outDecls: { prefix: string; uri: string }[] = [];
  if (nsDecls !== undefined) {
    if (!Array.isArray(nsDecls)) return null;
    for (const d of nsDecls) {
      if (!isObj(d) || typeof d.prefix !== 'string' || typeof d.uri !== 'string') return null;
      if (d.prefix !== '' && !NAME.test(d.prefix)) return null;
      outDecls.push({ prefix: d.prefix, uri: d.uri });
    }
  }
  const outKids: (NodeSpec | { text: string })[] = [];
  if (children !== undefined) {
    if (!Array.isArray(children)) return null;
    for (const c of children) {
      if (isObj(c) && 'text' in c) {
        if (typeof c.text !== 'string') return null;
        outKids.push({ text: c.text });
        continue;
      }
      const k = sanitizeNode(c, depth + 1, budget);
      if (!k) return null;
      outKids.push(k);
    }
  }
  return {
    ns,
    local,
    ...(prefix ? { prefix } : {}),
    attrs: outAttrs,
    nsDecls: outDecls,
    children: outKids,
  };
}

interface CarriedPart {
  readonly id: string;
  readonly type: string;
  readonly extension: string;
  readonly contentType: string;
  readonly bytes: Uint8Array;
  readonly part: string;
}

/** A validated relationship of a payload. */
type CleanRelationship =
  | ({ readonly kind: 'media' } & CarriedPart)
  | ({
      readonly kind: 'graph';
      /** The name of the new part, without its number and extension. */
      readonly stem: string;
      readonly relationships: readonly CleanRelationship[];
    } & CarriedPart)
  | { readonly kind: 'link'; readonly id: string; readonly type: string; readonly href: string }
  | { readonly kind: 'slide'; readonly id: string; readonly type: string; readonly part: string };

type GraphRelationship = Extract<CleanRelationship, { kind: 'graph' }>;

/**
 * Validate one untrusted relationship. Null refuses the payload; undefined drops the
 * relationship and the attributes that name it. A relationship of a carried part (`nested`)
 * cannot be dropped, because the part names it.
 */
function cleanRelationship(
  r: unknown,
  nested: boolean,
  depth: number,
  budget: { bytes: number; parts: number }
): CleanRelationship | null | undefined {
  if (!isObj(r)) return null;
  const { id, type, part, contentType, base64, target, relationships } = r;
  if (typeof id !== 'string' || typeof type !== 'string' || typeof part !== 'string') return null;
  if (typeof contentType !== 'string' || typeof base64 !== 'string') return null;
  // A nested id goes into a new relationships part as it is: it must be a name.
  if (nested && !NAME.test(id)) return null;
  const skip = nested ? null : undefined;
  if (type === REL.hyperlink) {
    const safe = typeof target === 'string' ? sanitizeHref(target) : null;
    return safe?.ok ? { kind: 'link', id, type, href: safe.href } : skip;
  }
  if (type === REL.slide) {
    const norm = normalizePartName(part);
    return !nested && norm.ok ? { kind: 'slide', id, type, part: norm.partName } : skip;
  }
  const graph = GRAPHS.get(type);
  if (!CARRIED.has(type) && !graph) return skip;
  const valid = graph
    ? graph.types.test(contentType)
    : /^[a-z]+\/[A-Za-z0-9.+-]+$/.test(contentType);
  const extension = /\.([A-Za-z0-9]{1,8})$/.exec(part)?.[1]?.toLowerCase();
  const norm = normalizePartName(part);
  if (!valid || !extension || !norm.ok) return graph ? null : skip;
  const bytes = fromBase64(base64);
  if (!bytes || depth > MAX_DEPTH) return null;
  budget.bytes += bytes.length;
  budget.parts += 1;
  if (budget.bytes > MAX_BYTES || budget.parts > MAX_PARTS) return null;
  const carried = { id, type, extension, contentType, bytes, part: norm.partName };
  if (!graph) return { kind: 'media', ...carried };
  if (relationships !== undefined && !Array.isArray(relationships)) return null;
  const out: CleanRelationship[] = [];
  const seen = new Set<string>();
  for (const c of relationships ?? []) {
    const clean = cleanRelationship(c, true, depth + 1, budget);
    if (!clean || seen.has(clean.id)) return null;
    seen.add(clean.id);
    out.push(clean);
  }
  // An embedded package keeps the stem of its name, such as `Microsoft_Excel_Worksheet`.
  const base = /\/([A-Za-z_][A-Za-z0-9_]{0,63}?)\d*\.[A-Za-z0-9]+$/.exec(norm.partName)?.[1];
  const stem = type === `${R}package` && base ? `/ppt/embeddings/${base}` : graph.stem;
  return { kind: 'graph', ...carried, stem, relationships: out };
}

/** Validate an untrusted payload; null when it is not a shape clipboard this editor accepts. */
export function parsePayload(
  v: unknown
): { payload: ClipboardPayload; shapes: NodeSpec[]; rels: CleanRelationship[] } | null {
  if (!isObj(v) || v.version !== 1) return null;
  if (typeof v.document !== 'string' || typeof v.slidePart !== 'string') return null;
  if (!Array.isArray(v.shapes) || !Array.isArray(v.relationships)) return null;
  if (v.shapes.length === 0 || v.shapes.length > MAX_SHAPES) return null;
  const budget = { n: MAX_NODES };
  const shapes: NodeSpec[] = [];
  for (const s of v.shapes) {
    const spec = sanitizeNode(s, 0, budget);
    const alternate = spec?.ns === NS.mc && spec.local === 'AlternateContent';
    if (!spec || (!alternate && (spec.ns !== NS.p || !SHAPES.has(spec.local)))) return null;
    shapes.push(spec);
  }
  const rels: CleanRelationship[] = [];
  const carried = { bytes: 0, parts: 0 };
  for (const r of v.relationships) {
    const clean = cleanRelationship(r, false, 0, carried);
    if (clean === null) return null;
    if (clean) rels.push(clean);
  }
  return {
    payload: v as unknown as ClipboardPayload,
    shapes,
    rels,
  };
}

const sameBytes = (a: Uint8Array, b: Uint8Array): boolean =>
  a.length === b.length && a.every((x, i) => x === b[i]);

/** The existing part with these bytes, or a new media part. */
function partFor(batch: OpBatch, rel: CarriedPart): string {
  const named = batch.pkg.parts.get(rel.part);
  if (named?.kind === 'binary' && sameBytes(named.bytes, rel.bytes)) return rel.part;
  for (const [name, part] of batch.pkg.parts)
    if (
      part.kind === 'binary' &&
      name.startsWith('/ppt/media/') &&
      part.bytes.length === rel.bytes.length &&
      sameBytes(part.bytes, rel.bytes)
    )
      return name;
  const stem = rel.type.endsWith('/image') ? 'image' : 'media';
  const name = freePartName(batch, `/ppt/media/${stem}`, `.${rel.extension}`);
  batch.push({ op: 'addPart', part: name, base64: toBase64(rel.bytes) });
  ensureContentType(batch, rel.extension, rel.contentType, name);
  return name;
}

/**
 * The bytes of a pasted diagram data part, whose `dsp:dataModelExt/@relId` names the new slide
 * relationship of its saved drawing. A drawing that did not come along leaves the id empty, so
 * the diagram never shows the drawing of another diagram.
 */
function withDrawingId(bytes: Uint8Array, slideIds: ReadonlyMap<string, string>): Uint8Array {
  const text = new TextDecoder('utf-8').decode(bytes);
  const ext = /(<(?:[A-Za-z_][\w.-]*:)?dataModelExt\b[^>]*?\brelId=)("[^"]*"|'[^']*')/.exec(text);
  if (!ext) return bytes;
  const id = slideIds.get(ext[2]!.slice(1, -1)) ?? '';
  const at = ext.index + ext[1]!.length;
  return new TextEncoder().encode(
    `${text.slice(0, at)}"${xmlAttr(id, 'relationship')}"${text.slice(at + ext[2]!.length)}`
  );
}

/**
 * Add a carried part graph under new names, as a duplicated slide copies the parts it owns.
 * Returns the name of the new part.
 */
function addGraph(
  batch: OpBatch,
  rel: GraphRelationship,
  slideIds: ReadonlyMap<string, string>
): string {
  const name = freePartName(batch, rel.stem, `.${rel.extension}`);
  const bytes = rel.type === `${R}diagramData` ? withDrawingId(rel.bytes, slideIds) : rel.bytes;
  batch.push({ op: 'addPart', part: name, base64: toBase64(bytes) });
  addOverride(batch, name, rel.contentType);
  if (rel.relationships.length === 0) return name;
  const items = rel.relationships.map((c) => {
    const external = c.kind === 'link';
    const target =
      c.kind === 'link'
        ? c.href
        : relativeTarget(
            name,
            c.kind === 'media'
              ? partFor(batch, c)
              : c.kind === 'graph'
                ? addGraph(batch, c, slideIds)
                : c.part
          );
    return `<Relationship Id="${xmlAttr(c.id, 'relationship')}" Type="${xmlAttr(c.type, 'relationship')}" Target="${xmlAttr(target, 'relationship')}"${external ? ' TargetMode="External"' : ''}/>`;
  });
  batch.push({
    op: 'addPart',
    part: relsPartName(name),
    xml: `${XML_DECL}<Relationships xmlns="${NS.rel}">${items.join('')}</Relationships>`,
  });
  return name;
}

/**
 * Paste shapes onto a slide's top level. Returns the new shape ids. `offset` is applied to each
 * shape's position (pasting onto the slide the shapes came from). A link to a slide is kept
 * only when the shapes come from the same presentation (`sameDocument`).
 */
export function pasteShapes(
  batch: OpBatch,
  part: string,
  raw: unknown,
  offset: number,
  sameDocument = false
): string[] {
  const parsed = parsePayload(raw);
  if (!parsed) throw new CommandRefusal('errors.clipboard');
  // Rewire: payload rel id → new rel id on the target slide.
  const existing = relationshipsOf(batch.pkg, part);
  const reuse = (type: string, target: string): string =>
    existing.find((r) => !r.external && r.target === target && r.type === type)?.id ??
    addRelationship(batch, part, type, target);
  const slides = new Set(presentationInfo(batch.pkg).slides.map((s) => s.partName));
  const map = new Map<string, string>();
  // A diagram's data part names the slide relationship of its drawing: the data comes last.
  const data = (r: CleanRelationship): number => (r.type === `${R}diagramData` ? 1 : 0);
  for (const rel of [...parsed.rels].sort((a, b) => data(a) - data(b))) {
    if (rel.kind === 'media') map.set(rel.id, reuse(rel.type, partFor(batch, rel)));
    else if (rel.kind === 'graph')
      map.set(rel.id, addRelationship(batch, part, rel.type, addGraph(batch, rel, map)));
    else if (rel.kind === 'link')
      map.set(rel.id, addExternalRelationship(batch, part, REL.hyperlink, rel.href));
    else if (sameDocument && slides.has(rel.part)) map.set(rel.id, reuse(REL.slide, rel.part));
  }
  const rewrite = (spec: NodeSpec): NodeSpec | null => {
    let attrs = spec.attrs ?? [];
    // An empty `r:id` names no relationship: a slide jump (`#next-slide`) keeps it as it is.
    const relAttrs = attrs.filter((a) => a.ns === NS.r && a.value !== '');
    if (relAttrs.some((a) => !map.has(a.value)) && HYPERLINKS.has(spec.local)) return null;
    // A pasted table is a new table: it holds no structure mark (table-edit-marks.ts).
    if (isStructureMarkSpec(spec)) return null;
    // An `a:extLst` that held only marks goes with them.
    if (spec.local === 'extLst' && isOnlyStructureMarks(spec)) return null;
    attrs = attrs
      .filter((a) => a.ns !== NS.r || a.value === '' || map.has(a.value))
      .map((a) => (a.ns === NS.r && a.value !== '' ? { ...a, value: map.get(a.value)! } : a));
    const kids: (NodeSpec | { text: string })[] = [];
    for (const c of spec.children ?? []) {
      if ('text' in c) kids.push(c);
      else {
        const k = rewrite(c);
        if (k) kids.push(k);
      }
    }
    return { ...spec, attrs, children: kids };
  };
  const tree = spTreeOf(batch, part);
  const ext = tree.children.findIndex((c) => c.type === 'element' && c.local === 'extLst');
  let at = ext < 0 ? tree.children.length : ext;
  // New ids for every pasted shape; a connector pasted with its shapes stays glued to them.
  const { specs } = renumberShapes(parsed.shapes, shapeIds(batch, spTreeOf(batch, part)));
  const out: string[] = [];
  for (const shape of specs) {
    const spec = rewrite(shape);
    if (!spec) continue;
    const id = batch.insert({
      op: 'insertNode',
      part,
      parent: spTreeOf(batch, part).id,
      index: at,
      spec: offset ? shifted(spec, offset) : spec,
    });
    at += 1;
    // A pasted `mc:AlternateContent` is selected by the copy that the editor draws.
    const el = batch.element(part, id);
    const branch = isAlternateContent(el)
      ? alternateBranch(el, xmlPartOf(batch.pkg, part)?.root)
      : undefined;
    out.push(elements(branch).find((c) => SHAPES.has(c.local))?.id ?? id);
  }
  return out;
}

/** Move a shape spec's own `off` by `d` on both axes. */
function shifted(spec: NodeSpec, d: number): NodeSpec {
  // Every copy of an `mc:AlternateContent` moves.
  if (spec.ns === NS.mc && spec.local === 'AlternateContent')
    return {
      ...spec,
      children: spec.children?.map((b) =>
        'text' in b
          ? b
          : { ...b, children: b.children?.map((c) => ('text' in c ? c : shifted(c, d))) }
      ),
    };
  const isHost = (n: NodeSpec): boolean =>
    n.ns === NS.p && (n.local === 'spPr' || n.local === 'grpSpPr');
  const moveOff = (xfrm: NodeSpec): NodeSpec => ({
    ...xfrm,
    children: xfrm.children?.map((c) =>
      'text' in c || c.local !== 'off'
        ? c
        : {
            ...c,
            attrs: c.attrs?.map((a) =>
              a.ns === '' && (a.local === 'x' || a.local === 'y')
                ? { ...a, value: coordinate((Number(a.value) || 0) + d, a.local) }
                : a
            ),
          }
    ),
  });
  return {
    ...spec,
    children: spec.children?.map((c) => {
      if ('text' in c) return c;
      if (c.ns === NS.p && c.local === 'xfrm') return moveOff(c);
      if (!isHost(c)) return c;
      return {
        ...c,
        children: c.children?.map((x) =>
          'text' in x || !(x.ns === NS.a && x.local === 'xfrm') ? x : moveOff(x)
        ),
      };
    }),
  };
}
