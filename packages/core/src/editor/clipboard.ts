// The shape clipboard (design D6 of `group-and-multi-edit`). Copy serializes the selected
// shapes and the binary parts their relationships point to; paste validates the payload as
// untrusted input, renumbers shape ids, and rewires relationships into the target slide.
//
// External relationships are never carried (their attributes are dropped), so a pasted shape
// can never make the editor fetch anything. Relationships to XML parts (charts, diagrams, OLE)
// are not carried either; shapes that need one are left out of the copy.

import {
  isOnlyStructureMarks,
  isStructureMarkSpec,
  withoutStructureMarks,
} from '../automation/commands/table-edit-marks.ts';
import { asSavedForCopy } from '../automation/commands/table-repair.ts';
import {
  contentTypeOf,
  NS,
  normalizePartName,
  REL,
  relationshipsOf,
  specOf,
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
  addRelationship,
  coordinate,
  ensureContentType,
  freePartName,
  shapeIds,
  renumberShapes,
  spTreeOf,
} from '../automation/commands/package-edits.ts';

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
/** Relationship types whose binary targets travel with copied shapes. */
const CARRIED = new Set([
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/image',
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/video',
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio',
  'http://schemas.microsoft.com/office/2007/relationships/media',
]);

/** Relationship ids a subtree uses (`r:*` attributes). */
function relIds(el: Element, out: Set<string>, depth = 0): Set<string> {
  if (depth > MAX_DEPTH) return out;
  for (const a of el.attrs) if (a.ns === NS.r) out.add(a.value);
  for (const c of el.children) if (c.type === 'element') relIds(c, out, depth + 1);
  return out;
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
    const shapes = ids.map((id) => batch.element(slidePart, id));
    return copyShapes(pkg, slidePart, shapes, document);
  } catch (e) {
    if (e instanceof CommandRefusal) return null;
    throw e;
  }
}

/**
 * Serialize shapes of one slide. Shapes that need a relationship to an XML part are left out;
 * null when nothing is left.
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
  let bytes = 0;
  for (const shape of shapes) {
    const pending: ClipboardRelationship[] = [];
    let ok = true;
    for (const id of relIds(shape, new Set())) {
      const rel = rels.get(id);
      if (!rel || rel.external || !rel.target || carried.has(id)) continue;
      const part = pkg.parts.get(rel.target);
      // Hyperlinks to other slides are dropped on paste, with their `a:hlinkClick`.
      if (rel.type === REL.hyperlink) continue;
      if (!part || part.kind !== 'binary' || !CARRIED.has(rel.type)) {
        // Slide jumps and other XML targets are dropped on paste; frames that need one are skipped.
        if (shape.local === 'graphicFrame' || shape.local === 'grpSp') ok = false;
        continue;
      }
      bytes += part.bytes.length;
      pending.push({
        id,
        type: rel.type,
        part: rel.target,
        contentType: contentTypeOf(pkg, rel.target) ?? 'application/octet-stream',
        base64: toBase64(part.bytes),
      });
    }
    if (!ok || bytes > MAX_BYTES) continue;
    // A marked table goes as the canvas draws it and a save writes it, without its marks; a
    // table that the save leaves out is not copied.
    const saved = asSavedForCopy(pkg, slidePart, shape);
    if (!saved) continue;
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
  for (const s of payload.shapes) walk(s, []);
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

interface CleanRelationship {
  readonly id: string;
  readonly type: string;
  readonly extension: string;
  readonly contentType: string;
  readonly bytes: Uint8Array;
  readonly part: string;
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
    if (!spec || spec.ns !== NS.p || !SHAPES.has(spec.local)) return null;
    shapes.push(spec);
  }
  const rels: CleanRelationship[] = [];
  let total = 0;
  for (const r of v.relationships) {
    if (!isObj(r)) return null;
    const { id, type, part, contentType, base64 } = r;
    if (typeof id !== 'string' || typeof type !== 'string' || typeof part !== 'string') return null;
    if (typeof contentType !== 'string' || typeof base64 !== 'string') return null;
    if (!CARRIED.has(type) || !/^[a-z]+\/[A-Za-z0-9.+-]+$/.test(contentType)) continue;
    const ext = /\.([A-Za-z0-9]{1,8})$/.exec(part)?.[1]?.toLowerCase();
    const norm = normalizePartName(part);
    if (!ext || !norm.ok) continue;
    const bytes = fromBase64(base64);
    if (!bytes || (total += bytes.length) > MAX_BYTES) return null;
    rels.push({ id, type, extension: ext, contentType, bytes, part: norm.partName });
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
function partFor(batch: OpBatch, rel: CleanRelationship): string {
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
 * Paste shapes onto a slide's top level. Returns the new shape ids. `offset` is applied to each
 * shape's position (pasting onto the slide the shapes came from).
 */
export function pasteShapes(batch: OpBatch, part: string, raw: unknown, offset: number): string[] {
  const parsed = parsePayload(raw);
  if (!parsed) throw new CommandRefusal('errors.clipboard');
  // Rewire: payload rel id → new rel id on the target slide.
  const existing = relationshipsOf(batch.pkg, part);
  const map = new Map<string, string>();
  for (const rel of parsed.rels) {
    const target = partFor(batch, rel);
    const reuse = existing.find((r) => !r.external && r.target === target && r.type === rel.type);
    map.set(rel.id, reuse ? reuse.id : addRelationship(batch, part, rel.type, target));
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
    out.push(id);
  }
  return out;
}

/** Move a shape spec's own `off` by `d` on both axes. */
function shifted(spec: NodeSpec, d: number): NodeSpec {
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
