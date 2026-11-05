// Custom XML parts (ECMA-376 Part 1, 22.5): a data part (`/customXml/itemN.xml`) with a
// properties part (`itemPropsN.xml`, `ds:datastoreItem`) that gives it a GUID `ds:itemID`.
//
// - The presentation holds a part through a `customXml` relationship of the presentation part.
// - A slide, a layout, a master, or a shape holds a part through a customer data entry
//   (`p:custDataLst/p:custData/@r:id`, 19.3.1.17) that names a `customXml` relationship of its
//   part.
//
// A relationship of the right type is a claim of the file, not proof: a part is listed only when
// it is not a relationships part, it relates to a properties part of the properties content type,
// and that part carries an item id. The XML a consumer writes goes through the bounded reader
// (no DTD, no entities, depth and size caps) before it reaches the package.

import { partNameKey } from '@eigenpal/ooxml-core';
import {
  attr,
  child,
  contentTypeOf,
  elements,
  isValidXmlText,
  parseXmlPart,
  presentationInfo,
  relationshipsOf,
  serializeXmlPart,
  xmlPartOf,
  type Element,
  type PptxPackage,
} from '../store/index.ts';
import { toBase64 } from './commands/base64.ts';
import type { OpBatch } from './commands/op-batch.ts';
import {
  addOverride,
  addRelationship,
  ensureContentType,
  removeRelationship,
  xmlAttr,
} from './commands/package-edits.ts';
import { specFromXml } from './commands/shape-commands.ts';
import { removeOrphans } from './commands/slide-structure.ts';
import { rootOf } from './plan-properties.ts';
import type { AutomationCustomXmlPartRead, AutomationRef, AutomationValue } from './protocol.ts';
import {
  AutomationFailure,
  edit,
  ownerPart,
  shapeAt,
  targetOf,
  type PlanContext,
} from './targets.ts';

const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';
const REL_CUSTOM_XML = `${R}customXml`;
const REL_CUSTOM_XML_PROPS = `${R}customXmlProps`;
const CT_PROPS = 'application/vnd.openxmlformats-officedocument.customXmlProperties+xml';
const NS_DS = 'http://schemas.openxmlformats.org/officeDocument/2006/customXml';

/** The longest XML text a write accepts, in UTF-16 code units. */
export const MAX_CUSTOM_XML = 8 * 1024 * 1024;
/** The most custom XML parts one owner lists. */
const MAX_LISTED = 4096;
/**
 * The most relationship ids one listing looks at. A file can name many ids that lead nowhere;
 * the scan stops here even when fewer parts were listed, so a listing stays linear and bounded.
 */
const MAX_SCANNED = 4 * MAX_LISTED;

const GUID =
  /^\{?([0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12})\}?$/;
const RELS_PART = /\/_rels\/[^/]*\.rels$/i;

/** An item id in the canonical form: braced and upper case. Undefined when it is not a GUID. */
export function canonicalItemId(value: unknown): string | undefined {
  if (typeof value !== 'string') return undefined;
  const m = GUID.exec(value.trim());
  return m ? `{${m[1]!.toUpperCase()}}` : undefined;
}

/** One listed part: its data part, its properties part, and what the object model shows. */
interface Listed {
  readonly item: string;
  readonly props: string;
  /** The relationship of the owner that names the data part. */
  readonly rId: string;
  readonly read: AutomationCustomXmlPartRead;
}

/** Where an owner keeps its custom XML parts. */
interface Owner {
  readonly part: string;
  /** For a slide, layout, master, or shape: the element that holds `p:custDataLst`. */
  readonly host?: Element;
  /** Child local names that come after `p:custDataLst` in the host's schema sequence. */
  readonly after: readonly string[];
}

function ownerOf(ctx: PlanContext, ref: AutomationRef): Owner {
  const target = targetOf(ctx, ref, 'presentation', 'slide', 'slideLayout', 'slideMaster', 'shape');
  const pkg = ctx.batch.pkg;
  if (target.kind === 'presentation') {
    const part = presentationInfo(pkg).presentationPart;
    if (!xmlPartOf(pkg, part))
      throw new AutomationFailure('document-unavailable', 'no presentation part');
    return { part, after: [] };
  }
  if (target.kind === 'shape') {
    const at = shapeAt(ctx, ref);
    const nv = elements(at.el).find((c) => c.local.startsWith('nv') && c.local.endsWith('Pr'));
    const host = child(nv, 'p:nvPr');
    if (!host) throw new AutomationFailure('unsupported-content', 'the shape has no p:nvPr');
    // CT_ApplicationNonVisualDrawingProps (19.3.1.31): ph, media, custDataLst, extLst.
    return { part: at.part, host, after: ['extLst'] };
  }
  const part = ownerPart(
    pkg,
    target.kind === 'slide' ? target : { kind: target.kind, part: target.part }
  );
  const host = child(xmlPartOf(pkg, part)?.root, 'p:cSld');
  if (!host) throw new AutomationFailure('unsupported-content', 'the part has no p:cSld');
  // CT_CommonSlideData (19.3.1.6): bg, spTree, custDataLst, controls, extLst.
  return { part, host, after: ['controls', 'extLst'] };
}

/** The relationship ids an owner names for its parts, in order, at most `MAX_SCANNED`. */
function namedIds(pkg: PptxPackage, owner: Owner): string[] {
  const out: string[] = [];
  if (!owner.host) {
    for (const r of relationshipsOf(pkg, owner.part)) {
      if (out.length >= MAX_SCANNED) break;
      if (r.type === REL_CUSTOM_XML && !r.external) out.push(r.id);
    }
    return out;
  }
  for (const e of elements(child(owner.host, 'p:custDataLst'))) {
    if (out.length >= MAX_SCANNED) break;
    const id = e.local === 'custData' ? attr(e, 'r:id') : undefined;
    if (id) out.push(id);
  }
  return out;
}

/** The item id of a properties part, when it is a `ds:datastoreItem` with one. */
function itemIdOf(pkg: PptxPackage, props: string): string | undefined {
  const root = rootOf(pkg, props);
  if (root?.ns !== NS_DS || root.local !== 'datastoreItem') return undefined;
  const id = root.attrs.find((a) => a.ns === NS_DS && a.local === 'itemID')?.value;
  return canonicalItemId(id);
}

/** The custom XML parts of an owner, in the order the owner names them. */
function listed(pkg: PptxPackage, owner: Owner): Listed[] {
  // One map of the owner's custom XML relationships, so each id is one lookup.
  const rels = new Map<string, string | undefined>();
  for (const r of relationshipsOf(pkg, owner.part))
    if (!r.external && r.type === REL_CUSTOM_XML && !rels.has(r.id)) rels.set(r.id, r.target);
  const out: Listed[] = [];
  const seen = new Set<string>();
  for (const rId of namedIds(pkg, owner)) {
    if (out.length >= MAX_LISTED) break;
    const item = rels.get(rId);
    if (!item || RELS_PART.test(item) || !pkg.parts.has(item)) continue;
    const propsRel = relationshipsOf(pkg, item).find(
      (r) => r.type === REL_CUSTOM_XML_PROPS && !r.external
    );
    const props = propsRel?.target;
    if (!props || contentTypeOf(pkg, props) !== CT_PROPS) continue;
    const id = itemIdOf(pkg, props);
    if (!id || seen.has(id)) continue;
    seen.add(id);
    const root = rootOf(pkg, item);
    out.push({ item, props, rId, read: { id, namespaceUri: root?.ns ?? '' } });
  }
  return out;
}

export function listCustomXmlParts(
  ctx: PlanContext,
  ref: AutomationRef,
  namespaceUri?: string
): AutomationValue {
  if (namespaceUri !== undefined && typeof namespaceUri !== 'string')
    throw new AutomationFailure('invalid-argument', 'invalid namespaceUri', 'namespaceUri');
  const parts = listed(ctx.batch.pkg, ownerOf(ctx, ref)).map((p) => p.read);
  return {
    kind: 'customXmlParts',
    parts:
      namespaceUri === undefined ? parts : parts.filter((p) => p.namespaceUri === namespaceUri),
  };
}

/** The listed part with this id, or `item-not-found`. */
function find(ctx: PlanContext, owner: Owner, id: unknown): Listed {
  const key = canonicalItemId(id);
  if (!key) throw new AutomationFailure('invalid-argument', 'invalid id', 'id');
  const hit = listed(ctx.batch.pkg, owner).find((p) => p.read.id === key);
  if (!hit) throw new AutomationFailure('item-not-found', 'no such custom XML part', 'id');
  return hit;
}

function decode(bytes: Uint8Array): string {
  if (bytes[0] === 0xff && bytes[1] === 0xfe) return new TextDecoder('utf-16le').decode(bytes);
  if (bytes[0] === 0xfe && bytes[1] === 0xff) return new TextDecoder('utf-16be').decode(bytes);
  const text = new TextDecoder().decode(bytes);
  return text.charCodeAt(0) === 0xfeff ? text.slice(1) : text;
}

export function readCustomXml(ctx: PlanContext, ref: AutomationRef, id: string): AutomationValue {
  const hit = find(ctx, ownerOf(ctx, ref), id);
  const part = ctx.batch.pkg.parts.get(hit.item)!;
  return {
    kind: 'text',
    text: part.kind === 'xml' ? serializeXmlPart(part.xml) : decode(part.bytes),
  };
}

const DECLARED_ENCODING = /^\s*<\?xml[^>]*\bencoding\s*=\s*["']([^"']*)["']/i;

/**
 * XML text from a consumer, checked by the bounded reader: one root element, no DTD, no
 * entities, inside the depth and size caps. The part stores it as UTF-8, so a declaration of
 * another encoding is refused. Returns the root's namespace.
 */
function checkedXml(xml: unknown): { readonly text: string; readonly namespaceUri: string } {
  const refuse = (): never => {
    throw new AutomationFailure('invalid-argument', 'invalid custom XML', 'custom-xml');
  };
  if (typeof xml !== 'string' || xml.length === 0 || xml.length > MAX_CUSTOM_XML) refuse();
  const text =
    (xml as string).charCodeAt(0) === 0xfeff ? (xml as string).slice(1) : (xml as string);
  const encoding = DECLARED_ENCODING.exec(text)?.[1]?.toLowerCase();
  if (encoding !== undefined && encoding !== 'utf-8' && encoding !== 'utf8') refuse();
  const parsed = parseXmlPart('/customXml/item.xml', text);
  if (!parsed.ok) refuse();
  return { text, namespaceUri: parsed.ok ? parsed.part.root.ns : '' };
}

/** Every item id the package carries, from each `ds:datastoreItem` part. */
function usedItemIds(pkg: PptxPackage): Set<string> {
  const out = new Set<string>();
  for (const name of pkg.parts.keys()) {
    if (contentTypeOf(pkg, name) !== CT_PROPS) continue;
    const id = itemIdOf(pkg, name);
    if (id) out.add(id);
  }
  return out;
}

/** A random item id, for a host that is not given one. */
function randomItemId(): string {
  const b = globalThis.crypto.getRandomValues(new Uint8Array(16));
  b[6] = (b[6]! & 0x0f) | 0x40;
  b[8] = (b[8]! & 0x3f) | 0x80;
  const hex = Array.from(b, (x) => x.toString(16).padStart(2, '0'))
    .join('')
    .toUpperCase();
  return `{${hex.slice(0, 8)}-${hex.slice(8, 12)}-${hex.slice(12, 16)}-${hex.slice(16, 20)}-${hex.slice(20)}}`;
}

/**
 * The first free `/customXml/itemN.xml` whose properties part and relationships part are free
 * too. Names compare the OPC way, and a name the session held before stays taken.
 */
function freeItemNames(batch: OpBatch): { item: string; props: string } {
  const taken = new Set<string>(batch.reservedPartNames);
  for (const name of batch.pkg.parts.keys()) taken.add(partNameKey(name));
  for (let n = 1; n < 100_000; n += 1) {
    const item = `/customXml/item${n}.xml`;
    const props = `/customXml/itemProps${n}.xml`;
    const rels = `/customXml/_rels/item${n}.xml.rels`;
    if (![item, props, rels].some((p) => taken.has(partNameKey(p)))) return { item, props };
  }
  throw new AutomationFailure('invalid-argument', 'too many custom XML parts', 'xml');
}

/** Name a new relationship in the owner's customer data list (`p:custDataLst`). */
function addCustData(batch: OpBatch, owner: Owner, rId: string): void {
  const fresh = batch.element(owner.part, owner.host!.id);
  const entry = `<p:custData r:id="${xmlAttr(rId, 'rId')}"/>`;
  const list = child(fresh, 'p:custDataLst');
  if (list) {
    // CT_CustomerDataList (19.3.1.18): custData*, then at most one tags.
    const tags = list.children.findIndex((c) => c.type === 'element' && c.local === 'tags');
    batch.insert({
      op: 'insertNode',
      part: owner.part,
      parent: list.id,
      index: tags < 0 ? list.children.length : tags,
      spec: specFromXml(entry),
    });
    return;
  }
  let index = fresh.children.length;
  fresh.children.forEach((c, i) => {
    if (c.type === 'element' && owner.after.includes(c.local) && i < index) index = i;
  });
  batch.insert({
    op: 'insertNode',
    part: owner.part,
    parent: fresh.id,
    index,
    spec: specFromXml(`<p:custDataLst>${entry}</p:custDataLst>`),
  });
}

export function addCustomXmlPart(
  ctx: PlanContext,
  ref: AutomationRef,
  xml: string,
  id?: string
): AutomationValue {
  const checked = checkedXml(xml);
  const owner = ownerOf(ctx, ref);
  const batch = ctx.batch;
  const itemId = id === undefined ? randomItemId() : canonicalItemId(id);
  if (!itemId) throw new AutomationFailure('invalid-argument', 'invalid id', 'id');
  if (usedItemIds(batch.pkg).has(itemId))
    throw new AutomationFailure('invalid-argument', 'that id is taken', 'custom-xml-id');
  edit(() => {
    const { item, props } = freeItemNames(batch);
    batch.push({
      op: 'addPart',
      part: item,
      base64: toBase64(new TextEncoder().encode(checked.text)),
    });
    if (contentTypeOf(batch.pkg, item) === undefined)
      ensureContentType(batch, 'xml', 'application/xml', item);
    const schemaRefs = checked.namespaceUri
      ? `<ds:schemaRefs><ds:schemaRef ds:uri="${xmlAttr(checked.namespaceUri, 'namespace')}"/></ds:schemaRefs>`
      : '<ds:schemaRefs/>';
    batch.push({
      op: 'addPart',
      part: props,
      base64: toBase64(
        new TextEncoder().encode(
          `<?xml version="1.0" encoding="UTF-8" standalone="no"?>\r\n<ds:datastoreItem ds:itemID="${itemId}" xmlns:ds="${NS_DS}">${schemaRefs}</ds:datastoreItem>`
        )
      ),
    });
    addOverride(batch, props, CT_PROPS);
    addRelationship(batch, item, REL_CUSTOM_XML_PROPS, props);
    const rId = addRelationship(batch, owner.part, REL_CUSTOM_XML, item);
    if (owner.host) addCustData(batch, owner, rId);
  });
  return {
    kind: 'customXmlParts',
    parts: [{ id: itemId, namespaceUri: checked.namespaceUri }],
  };
}

export function setCustomXml(
  ctx: PlanContext,
  ref: AutomationRef,
  id: string,
  xml: string
): AutomationValue {
  const checked = checkedXml(xml);
  const hit = find(ctx, ownerOf(ctx, ref), id);
  edit(() => {
    // A part another owner also names changes for both: it is one part.
    ctx.batch.push({ op: 'removePart', part: hit.item });
    ctx.batch.push({
      op: 'addPart',
      part: hit.item,
      base64: toBase64(new TextEncoder().encode(checked.text)),
    });
    if (checked.namespaceUri !== hit.read.namespaceUri) refreshSchemaRefs(ctx, hit, checked);
  });
  return { kind: 'applied' };
}

/**
 * Point the properties part's schema references (`ds:schemaRefs`, ECMA-376 Part 1, 22.5) at
 * the new root namespace: the reference to the old namespace goes, the new one comes first, and
 * the other references stay. The item id stays.
 */
function refreshSchemaRefs(
  ctx: PlanContext,
  hit: Listed,
  checked: { readonly namespaceUri: string }
): void {
  const root = rootOf(ctx.batch.pkg, hit.props);
  const schemaRefs = root
    ? elements(root).find((e) => e.ns === NS_DS && e.local === 'schemaRefs')
    : undefined;
  const refs = elements(schemaRefs)
    .filter((e) => e.ns === NS_DS && e.local === 'schemaRef')
    .map((e) => e.attrs.find((a) => a.ns === NS_DS && a.local === 'uri')?.value ?? '')
    .filter((uri) => uri !== '' && uri !== hit.read.namespaceUri && isValidXmlText(uri))
    .slice(0, MAX_LISTED);
  const uris = checked.namespaceUri ? [checked.namespaceUri, ...refs] : refs;
  const unique = [...new Set(uris)];
  const list = unique
    .map((uri) => `<ds:schemaRef ds:uri="${xmlAttr(uri, 'namespace')}"/>`)
    .join('');
  const xml =
    `<?xml version="1.0" encoding="UTF-8" standalone="no"?>\r\n<ds:datastoreItem ds:itemID="${hit.read.id}" xmlns:ds="${NS_DS}">` +
    (list ? `<ds:schemaRefs>${list}</ds:schemaRefs>` : '<ds:schemaRefs/>') +
    '</ds:datastoreItem>';
  ctx.batch.push({ op: 'removePart', part: hit.props });
  ctx.batch.push({
    op: 'addPart',
    part: hit.props,
    base64: toBase64(new TextEncoder().encode(xml)),
  });
}

export function deleteCustomXmlPart(
  ctx: PlanContext,
  ref: AutomationRef,
  id: string
): AutomationValue {
  const owner = ownerOf(ctx, ref);
  const hit = find(ctx, owner, id);
  const batch = ctx.batch;
  edit(() => {
    if (owner.host) {
      const list = child(batch.element(owner.part, owner.host.id), 'p:custDataLst');
      const entries = elements(list).filter((e) => e.local === 'custData');
      const named = entries.filter((e) => attr(e, 'r:id') === hit.rId);
      if (list && named.length === entries.length && elements(list).length === entries.length)
        batch.push({ op: 'removeNode', part: owner.part, node: list.id });
      else for (const e of named) batch.push({ op: 'removeNode', part: owner.part, node: e.id });
    }
    removeRelationship(batch, owner.part, hit.rId);
    // The parts go only when nothing else names them: a duplicated slide can share one.
    removeOrphans(batch, new Set([hit.item, hit.props]));
  });
  return { kind: 'applied' };
}
