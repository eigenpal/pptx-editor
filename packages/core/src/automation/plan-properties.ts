// Document properties, page setup, and tags.
//
// - Core properties: the core properties part (ECMA-376 Part 2, 11, `docProps/core.xml`).
// - Extended properties: `docProps/app.xml` (ECMA-376 Part 1, 22.2), for the company and manager.
// - Tags: a tags part (`p:tagLst`, ECMA-376 Part 1, 19.3.3) that the owner's customer data list
//   (`p:custDataLst/p:tags/@r:id`) names. Keys are stored in upper case.

import {
  attr,
  binaryPartOf,
  child,
  parseXmlPart,
  elements,
  presentationInfo,
  relationshipsOf,
  relsPartName,
  textOf,
  xmlPartOf,
  type Element,
  type NodeSpec,
  isValidXmlText,
  type PptxPackage,
} from '../store/index.ts';
import { partNameKey } from '@eigenpal/ooxml-core';
import type { OpBatch } from './commands/op-batch.ts';
import {
  addOverride,
  addRelationship,
  freePartName,
  xmlAttr,
  xmlText,
} from './commands/package-edits.ts';
import { specFromXml } from './commands/shape-commands.ts';
import type { AutomationPropertiesWrite } from './operations.ts';
import type { AutomationPropertiesRead, AutomationRef, AutomationValue } from './protocol.ts';
import {
  AutomationFailure,
  badArgument,
  edit,
  numberIn,
  shapeAt,
  slidePartOf,
  targetOf,
  toEmu,
  type PlanContext,
} from './targets.ts';

const XML_DECL = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n';

const NS_CP = 'http://schemas.openxmlformats.org/package/2006/metadata/core-properties';
const NS_DC = 'http://purl.org/dc/elements/1.1/';
const NS_DCTERMS = 'http://purl.org/dc/terms/';
const NS_EP = 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties';
const REL_CORE =
  'http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties';
const REL_APP =
  'http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties';
const REL_TAGS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/tags';
const CT_CORE = 'application/vnd.openxmlformats-package.core-properties+xml';
const CT_APP = 'application/vnd.openxmlformats-officedocument.extended-properties+xml';
const CT_TAGS = 'application/vnd.openxmlformats-officedocument.presentationml.tags+xml';

/** Each core property: its element namespace, prefix, and local name. */
const CORE_FIELDS = {
  title: [NS_DC, 'dc', 'title'],
  subject: [NS_DC, 'dc', 'subject'],
  author: [NS_DC, 'dc', 'creator'],
  keywords: [NS_CP, 'cp', 'keywords'],
  comments: [NS_DC, 'dc', 'description'],
  category: [NS_CP, 'cp', 'category'],
  lastAuthor: [NS_CP, 'cp', 'lastModifiedBy'],
  revisionNumber: [NS_CP, 'cp', 'revision'],
} as const;

const APP_FIELDS = { company: 'Company', manager: 'Manager' } as const;

const parsed = new WeakMap<Uint8Array, Element | null>();

/** The text of an XML part kept as bytes, honoring a UTF-8 or UTF-16 byte order mark. */
function decodeXml(bytes: Uint8Array): string {
  if (bytes[0] === 0xff && bytes[1] === 0xfe) return new TextDecoder('utf-16le').decode(bytes);
  if (bytes[0] === 0xfe && bytes[1] === 0xff) return new TextDecoder('utf-16be').decode(bytes);
  const text = new TextDecoder().decode(bytes);
  return text.charCodeAt(0) === 0xfeff ? text.slice(1) : text;
}

/**
 * The root of an XML part. The reader keeps the property and tags parts as bytes (it does not
 * interpret them), so those parse here through the same bounded reader, with no DTD.
 */
export function rootOf(pkg: PptxPackage, part: string | undefined): Element | undefined {
  const xml = xmlPartOf(pkg, part);
  if (xml) return xml.root;
  const bytes = binaryPartOf(pkg, part);
  if (!bytes || !part) return undefined;
  if (!parsed.has(bytes)) {
    const r = parseXmlPart(part, decodeXml(bytes));
    parsed.set(bytes, r.ok ? r.part.root : null);
  }
  return parsed.get(bytes) ?? undefined;
}

/** The root of a part to edit: a part kept as bytes becomes an XML part in this batch. */
export function editableRoot(batch: OpBatch, part: string): Element {
  const xml = xmlPartOf(batch.pkg, part);
  if (xml) return xml.root;
  const bytes = binaryPartOf(batch.pkg, part);
  if (!bytes || !rootOf(batch.pkg, part))
    throw new AutomationFailure('unsupported-content', 'that part is not XML', 'part');
  batch.push({ op: 'removePart', part });
  batch.push({ op: 'addPart', part, xml: decodeXml(bytes) });
  return xmlPartOf(batch.pkg, part)!.root;
}

/** The target of the package relationship of `type`, when its part is XML. */
export function packagePart(pkg: PptxPackage, type: string): string | undefined {
  const rel = relationshipsOf(pkg, '/').find((r) => r.type === type && !r.external);
  return rel?.target && rootOf(pkg, rel.target) ? rel.target : undefined;
}

function fieldOf(root: Element | undefined, ns: string, local: string): Element | undefined {
  return elements(root).find((e) => e.ns === ns && e.local === local);
}

/** The core and extended properties as plain values. */
export function corePropertiesOf(pkg: PptxPackage): AutomationPropertiesRead {
  const core = rootOf(pkg, packagePart(pkg, REL_CORE));
  const app = rootOf(pkg, packagePart(pkg, REL_APP));
  const text = (field: keyof typeof CORE_FIELDS): string => {
    const [ns, , local] = CORE_FIELDS[field];
    return textOf(fieldOf(core, ns, local)).trim();
  };
  const revision = Number.parseInt(text('revisionNumber'), 10);
  const created = textOf(fieldOf(core, NS_DCTERMS, 'created')).trim();
  return {
    title: text('title'),
    subject: text('subject'),
    author: text('author'),
    keywords: text('keywords'),
    comments: text('comments'),
    category: text('category'),
    lastAuthor: text('lastAuthor'),
    revisionNumber: Number.isSafeInteger(revision) && revision >= 0 ? revision : 0,
    creationDate: created || null,
    company: textOf(fieldOf(app, NS_EP, APP_FIELDS.company)).trim(),
    manager: textOf(fieldOf(app, NS_EP, APP_FIELDS.manager)).trim(),
  };
}

export function getProperties(ctx: PlanContext): AutomationValue {
  return { kind: 'properties', properties: corePropertiesOf(ctx.batch.pkg) };
}

/** A package-level part, created with its relationship and content type when missing. */
export function ensurePart(
  batch: OpBatch,
  rel: string,
  name: string,
  contentType: string,
  xml: string
): string {
  const existing = packagePart(batch.pkg, rel);
  if (existing) return existing;
  // The usual name (`/docProps/custom.xml`) when it is free, else the first numbered one.
  const usual = `${name}.xml`;
  const taken = [...batch.pkg.parts.keys(), ...batch.reservedPartNames].some(
    (p) =>
      partNameKey(p) === partNameKey(usual) || partNameKey(p) === partNameKey(relsPartName(usual))
  );
  const part = taken ? freePartName(batch, name, '.xml') : usual;
  batch.push({ op: 'addPart', part, xml: `${XML_DECL}${xml}` });
  addOverride(batch, part, contentType);
  addRelationship(batch, '/', rel, part);
  return part;
}

/** Write one property element's text, inserting the element when missing. */
function writeField(
  batch: OpBatch,
  part: string,
  ns: string,
  prefix: string | undefined,
  local: string,
  value: string
): void {
  const root = editableRoot(batch, part);
  const field = fieldOf(root, ns, local);
  // Validate: the store refuses characters XML 1.0 cannot carry.
  xmlText(value, local);
  if (field) {
    batch.push({ op: 'setText', part, node: field.id, text: value });
    return;
  }
  const spec: NodeSpec = {
    ns,
    local,
    ...(prefix ? { prefix } : {}),
    attrs: [],
    children: value ? [{ text: value }] : [],
  };
  batch.push({ op: 'insertNode', part, parent: root.id, index: root.children.length, spec });
}

const CORE_XML = `<cp:coreProperties xmlns:cp="${NS_CP}" xmlns:dc="${NS_DC}" xmlns:dcterms="${NS_DCTERMS}" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"/>`;
const APP_XML = `<Properties xmlns="${NS_EP}" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"/>`;

/** The longest property text a write accepts. */
const MAX_PROPERTY = 4096;

export function setProperties(ctx: PlanContext, props: AutomationPropertiesWrite): AutomationValue {
  if (typeof props !== 'object' || props === null) badArgument('properties');
  const batch = ctx.batch;
  edit(() => {
    for (const [key, value] of Object.entries(props)) {
      if (value === undefined) continue;
      if (key === 'revisionNumber') {
        const n = numberIn(value, key, 0, 2_147_483_647);
        if (!Number.isInteger(n)) badArgument(key);
        const part = ensurePart(batch, REL_CORE, '/docProps/core', CT_CORE, CORE_XML);
        writeField(batch, part, NS_CP, 'cp', 'revision', String(n));
        continue;
      }
      if (typeof value !== 'string' || value.length > MAX_PROPERTY) badArgument(key);
      if (!isValidXmlText(value)) badArgument(key);
      if (Object.prototype.hasOwnProperty.call(CORE_FIELDS, key) && key !== 'lastAuthor') {
        const [ns, prefix, local] = CORE_FIELDS[key as keyof typeof CORE_FIELDS];
        const part = ensurePart(batch, REL_CORE, '/docProps/core', CT_CORE, CORE_XML);
        writeField(batch, part, ns, prefix, local, value);
      } else if (Object.prototype.hasOwnProperty.call(APP_FIELDS, key)) {
        const part = ensurePart(batch, REL_APP, '/docProps/app', CT_APP, APP_XML);
        writeField(
          batch,
          part,
          NS_EP,
          undefined,
          APP_FIELDS[key as keyof typeof APP_FIELDS],
          value
        );
      } else badArgument(key);
    }
  });
  return { kind: 'applied' };
}

/** `ST_SlideSizeCoordinate`: 1 to 56 inches (ECMA-376 Part 1, 19.7.17). */
const MIN_SLIDE = 914_400;
const MAX_SLIDE = 51_206_400;

export function setPageSetup(ctx: PlanContext, width?: number, height?: number): AutomationValue {
  const info = presentationInfo(ctx.batch.pkg);
  const sz = child(xmlPartOf(ctx.batch.pkg, info.presentationPart)?.root, 'p:sldSz');
  if (!sz) throw new AutomationFailure('unsupported-content', 'the presentation has no p:sldSz');
  for (const [name, value] of [
    ['cx', width],
    ['cy', height],
  ] as const) {
    if (value === undefined) continue;
    const emu = toEmu(numberIn(value, name === 'cx' ? 'slideWidth' : 'slideHeight', 72, 4032));
    if (emu < MIN_SLIDE || emu > MAX_SLIDE)
      badArgument(name === 'cx' ? 'slideWidth' : 'slideHeight');
    ctx.batch.push({
      op: 'setAttribute',
      part: info.presentationPart,
      node: sz.id,
      name: { ns: '', local: name },
      value: String(emu),
    });
    // A custom size: `type` names a preset size that no longer holds.
    if (attr(sz, 'type') !== undefined)
      ctx.batch.push({
        op: 'setAttribute',
        part: info.presentationPart,
        node: sz.id,
        name: { ns: '', local: 'type' },
        value: null,
      });
  }
  return { kind: 'applied' };
}

/** The element that holds an owner's customer data list, and where `p:custDataLst` goes. */
interface TagHost {
  readonly part: string;
  readonly host: Element;
  /** Child local names that come after `p:custDataLst` in the host's schema sequence. */
  readonly after: readonly string[];
}

function tagHost(ctx: PlanContext, ref: AutomationRef): TagHost {
  const target = targetOf(ctx, ref, 'presentation', 'slide', 'shape');
  const pkg = ctx.batch.pkg;
  if (target.kind === 'presentation') {
    const part = presentationInfo(pkg).presentationPart;
    const host = xmlPartOf(pkg, part)?.root;
    if (!host) throw new AutomationFailure('document-unavailable', 'no presentation part');
    // CT_Presentation (ECMA-376 Part 1, 19.2.1.26).
    return { part, host, after: ['kinsoku', 'defaultTextStyle', 'modifyVerifier', 'extLst'] };
  }
  if (target.kind === 'slide') {
    const part = slidePartOf(pkg, target.sldId);
    const host = child(xmlPartOf(pkg, part)?.root, 'p:cSld');
    if (!host) throw new AutomationFailure('unsupported-content', 'the slide has no p:cSld');
    // CT_CommonSlideData (19.3.1.6): bg, spTree, custDataLst, controls, extLst.
    return { part, host, after: ['controls', 'extLst'] };
  }
  const at = shapeAt(ctx, ref);
  const nv = elements(at.el).find((c) => c.local.startsWith('nv') && c.local.endsWith('Pr'));
  const host = child(nv, 'p:nvPr');
  if (!host) throw new AutomationFailure('unsupported-content', 'the shape has no p:nvPr');
  // CT_ApplicationNonVisualDrawingProps (19.3.1.31): ph, media, custDataLst, extLst.
  return { part: at.part, host, after: ['extLst'] };
}

/** The tags part an owner names, if any. */
function tagsPart(pkg: PptxPackage, host: TagHost): string | undefined {
  const rId = attr(child(child(host.host, 'p:custDataLst'), 'p:tags'), 'r:id');
  if (!rId) return undefined;
  const rel = relationshipsOf(pkg, host.part).find((r) => r.id === rId && !r.external);
  return rel?.target && rootOf(pkg, rel.target)?.local === 'tagLst' ? rel.target : undefined;
}

export function readTags(ctx: PlanContext, ref: AutomationRef): AutomationValue {
  const host = tagHost(ctx, ref);
  const part = tagsPart(ctx.batch.pkg, host);
  const root = rootOf(ctx.batch.pkg, part);
  return {
    kind: 'tags',
    tags: elements(root)
      .filter((t) => t.local === 'tag')
      .map((t) => ({ key: attr(t, 'name') ?? '', value: attr(t, 'val') ?? '' })),
  };
}

/** A tag key: the object model stores it in upper case. */
function tagKey(key: unknown): string {
  if (typeof key !== 'string' || key.length === 0 || key.length > 255) badArgument('key');
  if (!isValidXmlText(key)) badArgument('key');
  return key.toUpperCase();
}

/** The tags part of an owner, created with its relationship and list entry when missing. */
function ensureTags(batch: OpBatch, host: TagHost): string {
  const existing = tagsPart(batch.pkg, host);
  if (existing) return existing;
  const part = freePartName(batch, '/ppt/tags/tag', '.xml');
  batch.push({
    op: 'addPart',
    part,
    xml: `${XML_DECL}<p:tagLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>`,
  });
  addOverride(batch, part, CT_TAGS);
  const rId = addRelationship(batch, host.part, REL_TAGS, part);
  const fresh = batch.element(host.part, host.host.id);
  const list = child(fresh, 'p:custDataLst');
  if (list) {
    // CT_CustomerDataList (19.3.1.18): custData*, then at most one tags.
    batch.insert({
      op: 'insertNode',
      part: host.part,
      parent: list.id,
      index: list.children.length,
      spec: specFromXml(`<p:tags r:id="${xmlAttr(rId, 'rId')}"/>`),
    });
    return part;
  }
  let index = fresh.children.length;
  fresh.children.forEach((c, i) => {
    if (c.type === 'element' && host.after.includes(c.local) && i < index) index = i;
  });
  batch.insert({
    op: 'insertNode',
    part: host.part,
    parent: fresh.id,
    index,
    spec: specFromXml(`<p:custDataLst><p:tags r:id="${xmlAttr(rId, 'rId')}"/></p:custDataLst>`),
  });
  return part;
}

export function setTag(
  ctx: PlanContext,
  ref: AutomationRef,
  key: string,
  value: string
): AutomationValue {
  const name = tagKey(key);
  if (typeof value !== 'string' || value.length > 65_536) badArgument('value');
  // Characters XML 1.0 cannot carry are an argument error, not a failed edit.
  if (!isValidXmlText(value)) badArgument('value');
  const host = tagHost(ctx, ref);
  edit(() => {
    const part = ensureTags(ctx.batch, host);
    const root = editableRoot(ctx.batch, part);
    const tag = elements(root).find((t) => t.local === 'tag' && attr(t, 'name') === name);
    if (tag)
      ctx.batch.push({
        op: 'setAttribute',
        part,
        node: tag.id,
        name: { ns: '', local: 'val' },
        value,
      });
    else
      ctx.batch.insert({
        op: 'insertNode',
        part,
        parent: root.id,
        index: root.children.length,
        spec: specFromXml(
          `<p:tag name="${xmlAttr(name, 'key')}" val="${xmlAttr(value, 'value')}"/>`
        ),
      });
  });
  return { kind: 'applied' };
}

export function deleteTag(ctx: PlanContext, ref: AutomationRef, key: string): AutomationValue {
  const name = tagKey(key);
  const host = tagHost(ctx, ref);
  const part = tagsPart(ctx.batch.pkg, host);
  const found = elements(rootOf(ctx.batch.pkg, part)).some(
    (t) => t.local === 'tag' && attr(t, 'name') === name
  );
  if (!part || !found) throw new AutomationFailure('item-not-found', 'no such tag', 'key');
  const tag = elements(editableRoot(ctx.batch, part)).find(
    (t) => t.local === 'tag' && attr(t, 'name') === name
  )!;
  ctx.batch.push({ op: 'removeNode', part, node: tag.id });
  return { kind: 'applied' };
}
