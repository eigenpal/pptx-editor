// Slide import: copy slides of another presentation into this one, with every part they need.
//
// Each slide comes with its notes, its pictures and media, its charts, diagrams, embedded
// objects, tags, and custom XML parts, under fresh part names. Relationship ids stay as they are,
// because the copied XML names them; every relationship target, content type, slide id, master
// id, and layout id is new in this package. Two formats (Office.js `InsertSlideFormatting`):
//
// - Keep source formatting: the slide keeps its layout, master, and theme. They are copied too,
//   once per import, and the copied master lists only the copied layouts (ECMA-376 Part 1,
//   19.3.1.42 `sldLayoutIdLst`).
// - Use destination theme: the slide goes on the layout of this presentation of the same type
//   (`p:sldLayout/@type`; a custom layout matches by name), else on the first layout of the first
//   master. Its placeholders are re-mapped as Change Layout does (`applyLayout`): each placeholder
//   first takes the box it showed in the source as its own, so one with no place keeps it.
//
// Comment lists come along, with their authors merged into this presentation's comment author
// list (`import-comments.ts`). Threaded comments and embedded controls (`p:control`) name
// authors and state of the other package, so a slide with them is refused, naming the slide.
// A link to a slide that is not imported loses its action.

import { resolveSlide, type Xfrm } from '../../cascade/index.ts';
import {
  attr,
  child,
  CONTENT_TYPE,
  CONTENT_TYPES_PART,
  contentTypeOf,
  elements,
  NS,
  nodeIndex,
  path,
  presentationInfo,
  relationshipsOf,
  relsPartName,
  REL,
  resolveInternalTarget,
  serializeXmlPart,
  shapeTreeOf,
  xmlPartOf,
  type Element,
  type PptxOp,
  type PptxPackage,
} from '../../store/index.ts';
import { toBase64 } from './base64.ts';
import { CommandRefusal, type OpBatch } from './op-batch.ts';
import { partNameKey } from '@eigenpal/ooxml-core';
import { addRelationship, relativeTarget, xmlAttr, XML_DECL } from './package-edits.ts';
import { specFromXml } from './shape-commands.ts';
import { registerSlide } from './slide-commands.ts';
import { applyLayout, deckLayouts, keepBox } from './slide-layouts.ts';
import { CT_COMMENTS, mergeImportedComments, REL_COMMENTS } from './import-comments.ts';
import {
  creationIdsOf,
  freshCreationIdAmong,
  hasSections,
  placeInSection,
  treeOf,
} from './slide-structure.ts';

/** Which formatting the imported slides keep. */
export type ImportFormatting = 'keepSource' | 'useDestinationTheme';

/** What to import. */
export interface SlideImport {
  /** The other presentation, read through `readPptx`. */
  readonly source: PptxPackage;
  /** Source slide indexes, in the order they go in. */
  readonly slides: readonly number[];
  /** The slide index the new slides follow; -1 puts them first. */
  readonly after: number;
  readonly formatting: ImportFormatting;
}

const R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/';
const PML = 'application/vnd.openxmlformats-officedocument.presentationml.';
const CT_CUSTOM_XML_PROPS = 'application/vnd.openxmlformats-officedocument.customXmlProperties+xml';

/**
 * Relationships a copy cannot carry: threaded comments name their slide by its creation id and
 * their authors by GUID, and controls carry binary state. Comment lists come along
 * (`import-comments.ts`).
 */
export const REFUSED: ReadonlySet<string> = new Set([
  `${R}commentAuthors`,
  'http://schemas.microsoft.com/office/2018/10/relationships/comments',
  'http://schemas.microsoft.com/office/2018/10/relationships/authors',
  `${R}control`,
]);

const HYPERLINKS: ReadonlySet<string> = new Set(['hlinkClick', 'hlinkHover', 'hlinkMouseOver']);

/** The most parts one import copies. */
const MAX_PARTS = 5_000;

/** `ST_SlideMasterId` and `ST_SlideLayoutId` share one range (19.7.17, 19.7.14). */
const MIN_MASTER_ID = 2_147_483_648;
const MAX_MASTER_ID = 4_294_967_295;

type Role = 'slide' | 'notes' | 'layout' | 'master' | 'notesMaster' | 'other';

/** A relationships part name (ECMA-376 Part 2, 9.3.1). */
const RELS_PART = /\/_rels\/[^/]*\.rels$/i;

/** Parts that a package has once: the main document part, its properties, and the like. */
const PACKAGE_LEVEL: ReadonlySet<string> = new Set([
  `${PML}presProps+xml`,
  `${PML}viewProps+xml`,
  `${PML}handoutMaster+xml`,
  `${PML}commentAuthors+xml`,
  'application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml',
  'application/vnd.openxmlformats-package.core-properties+xml',
  'application/vnd.openxmlformats-officedocument.extended-properties+xml',
  'application/vnd.openxmlformats-officedocument.custom-properties+xml',
  'application/vnd.openxmlformats-package.relationships+xml',
]);

function packageLevel(contentType: string | undefined): boolean {
  return !contentType || contentType.endsWith('.main+xml') || PACKAGE_LEVEL.has(contentType);
}

function roleOf(pkg: PptxPackage, part: string): Role {
  switch (contentTypeOf(pkg, part)) {
    case `${PML}slide+xml`:
      return 'slide';
    case `${PML}notesSlide+xml`:
      return 'notes';
    case `${PML}slideLayout+xml`:
      return 'layout';
    case `${PML}slideMaster+xml`:
      return 'master';
    case `${PML}notesMaster+xml`:
      return 'notesMaster';
    default:
      return 'other';
  }
}

/** `/ppt/charts/chart12.xml` → [`/ppt/charts/chart`, `.xml`]: a free name keeps the pattern. */
function namePattern(partName: string): [string, string] {
  const slash = partName.lastIndexOf('/');
  const dot = partName.lastIndexOf('.');
  const ext = dot > slash ? partName.slice(dot) : '';
  const stem = partName.slice(0, partName.length - ext.length).replace(/\d+$/, '');
  return [stem, ext];
}

function decode(bytes: Uint8Array): string {
  if (bytes[0] === 0xff && bytes[1] === 0xfe) return new TextDecoder('utf-16le').decode(bytes);
  if (bytes[0] === 0xfe && bytes[1] === 0xff) return new TextDecoder('utf-16be').decode(bytes);
  const text = new TextDecoder().decode(bytes);
  return text.charCodeAt(0) === 0xfeff ? text.slice(1) : text;
}

/** A random GUID, braced and upper case, for a copied custom XML part (`ds:itemID`). */
function randomGuid(): string {
  const b = globalThis.crypto.getRandomValues(new Uint8Array(16));
  b[6] = (b[6]! & 0x0f) | 0x40;
  b[8] = (b[8]! & 0x3f) | 0x80;
  const h = Array.from(b, (x) => x.toString(16).padStart(2, '0'))
    .join('')
    .toUpperCase();
  return `{${h.slice(0, 8)}-${h.slice(8, 12)}-${h.slice(12, 16)}-${h.slice(16, 20)}-${h.slice(20)}}`;
}

/** `p:cNvPr` of a shape. */
function cNvPr(el: Element): Element | undefined {
  const nv = elements(el).find((c) => c.local.startsWith('nv') && c.local.endsWith('Pr'));
  return child(nv, 'p:cNvPr');
}

class Importer {
  /** Source part → destination part. */
  readonly #copies = new Map<string, string>();
  /** Imported source slide → its copy. */
  readonly slides = new Map<string, string>();
  /** Ops of the copy, applied to the batch in one pass by `flush`. */
  #ops: PptxOp[] = [];
  /** OPC keys (`partNameKey`) of every part name taken, in the package or by this import. */
  readonly #taken = new Set<string>();
  /** Name pattern → the next number to try, so each new name is one step, not a scan. */
  readonly #next = new Map<string, number>();
  /** The content types this import sees: defaults by extension, overrides by part key. */
  readonly #defaults = new Map<string, string>();
  readonly #overrides = new Map<string, string>();
  readonly #typesRoot: string;
  #typesLength: number;
  /** Copied source parts whose relationships part is still to write. */
  readonly #pending: string[] = [];
  /** Destination part → relationship ids it lost (a link to a slide that was not imported). */
  readonly #dropped = new Map<string, Set<string>>();
  readonly #masters: string[] = [];
  #notesMaster: string | undefined;
  #newNotesMaster = false;

  constructor(
    readonly batch: OpBatch,
    readonly source: PptxPackage,
    /** For the destination theme: imported source slide → its destination layout. */
    readonly layouts: ReadonlyMap<string, string> | null,
    /** For the source formatting: the source layouts the imported slides use. */
    readonly keepLayouts: ReadonlySet<string>
  ) {
    for (const name of batch.pkg.parts.keys()) this.#taken.add(partNameKey(name));
    for (const key of batch.reservedPartNames) this.#taken.add(key);
    const types = xmlPartOf(batch.pkg, CONTENT_TYPES_PART);
    if (!types) throw new CommandRefusal('errors.malformed', 'content types');
    this.#typesRoot = types.root.id;
    this.#typesLength = types.root.children.length;
    for (const e of elements(types.root)) {
      const type = attr(e, 'ContentType');
      if (!type) continue;
      const ext = attr(e, 'Extension');
      const name = attr(e, 'PartName');
      if (e.local === 'Default' && ext && !this.#defaults.has(ext.toLowerCase()))
        this.#defaults.set(ext.toLowerCase(), type);
      if (e.local === 'Override' && name && !this.#overrides.has(partNameKey(name)))
        this.#overrides.set(partNameKey(name), type);
    }
    this.#notesMaster = relationshipsOf(
      batch.pkg,
      presentationInfo(batch.pkg).presentationPart
    ).find(
      (r) => r.type === REL.notesMaster && !r.external && r.target && batch.pkg.parts.has(r.target)
    )?.target;
  }

  /** A part name that no part has and no other copy took: the lowest free number of a pattern. */
  #name(stem: string, ext: string): string {
    const pattern = `${stem}\u0000${ext}`;
    for (let n = this.#next.get(pattern) ?? 1; n < 100_000; n += 1) {
      const name = `${stem}${n}${ext}`;
      const key = partNameKey(name);
      const rels = partNameKey(relsPartName(name));
      if (this.#taken.has(key) || this.#taken.has(rels)) continue;
      this.#taken.add(key);
      this.#taken.add(rels);
      this.#next.set(pattern, n + 1);
      return name;
    }
    throw new CommandRefusal('errors.tooLarge', 'part name');
  }

  /** The copied comment lists. */
  readonly comments: string[] = [];

  /** Queue the copy of an imported slide under a new name. */
  addSlide(src: string): string {
    const xml = xmlPartOf(this.source, src);
    if (!xml) throw new CommandRefusal('errors.malformed', 'import slide');
    const name = this.#name('/ppt/slides/slide', '.xml');
    this.#ops.push({ op: 'addPart', part: name, xml: serializeXmlPart(xml) });
    this.slides.set(src, name);
    this.#copies.set(src, name);
    this.#pending.push(src);
    return name;
  }

  /** Apply every queued op of the copy to the batch, in one pass. */
  flush(): void {
    const ops = this.#ops;
    this.#ops = [];
    this.batch.pushAll(ops);
  }

  /** Copy one part under a free name, now, so the name is taken; its relationships come later. */
  copy(src: string): string {
    const done = this.#copies.get(src);
    if (done) return done;
    if (this.#copies.size >= MAX_PARTS) throw new CommandRefusal('errors.tooLarge', 'import');
    const part = this.source.parts.get(src);
    const contentType = contentTypeOf(this.source, src);
    if (!part || !contentType) throw new CommandRefusal('errors.malformed', `import ${src}`);
    const [stem, ext] = namePattern(src);
    const name = this.#name(stem, ext);
    this.#copies.set(src, name);
    if (contentType === CT_COMMENTS) {
      // A comment list is copied as a tree, so its author ids and indexes can be changed.
      const tree = treeOf(this.source, src);
      if (!tree) throw new CommandRefusal('errors.malformed', `import ${src}`);
      this.#ops.push({ op: 'addPart', part: name, xml: serializeXmlPart(tree) });
      this.comments.push(name);
    } else if (contentType === CT_CUSTOM_XML_PROPS) {
      // A copied data part needs its own item id: two parts with one id are one part to a reader.
      const text = part.kind === 'xml' ? serializeXmlPart(part.xml) : decode(part.bytes);
      const fresh = text.replace(/(\bitemID\s*=\s*["'])[^"']*(["'])/, `$1${randomGuid()}$2`);
      this.#ops.push({
        op: 'addPart',
        part: name,
        base64: toBase64(new TextEncoder().encode(fresh)),
      });
    } else
      this.#ops.push(
        part.kind === 'xml'
          ? { op: 'addPart', part: name, xml: serializeXmlPart(part.xml) }
          : { op: 'addPart', part: name, base64: toBase64(part.bytes) }
      );
    this.#typeAs(name, contentType);
    this.#pending.push(src);
    if (roleOf(this.source, src) === 'master') this.#masters.push(name);
    return name;
  }

  /** Give a copy the content type of its source: by the extension default, else an override. */
  #typeAs(name: string, contentType: string): void {
    const key = partNameKey(name);
    const dot = name.lastIndexOf('.');
    const ext = dot > name.lastIndexOf('/') ? name.slice(dot + 1) : '';
    const now =
      this.#overrides.get(key) ?? (ext ? this.#defaults.get(ext.toLowerCase()) : undefined);
    if (now === contentType) return;
    const entry = (local: string, attrs: readonly [string, string][]): PptxOp => ({
      op: 'insertNode',
      part: CONTENT_TYPES_PART,
      parent: this.#typesRoot,
      index: local === 'Default' ? 0 : this.#typesLength,
      spec: { ns: NS.ct, local, attrs: attrs.map(([n, value]) => ({ ns: '', local: n, value })) },
    });
    if (now === undefined && ext) {
      this.#ops.push(
        entry('Default', [
          ['Extension', ext],
          ['ContentType', contentType],
        ])
      );
      this.#defaults.set(ext.toLowerCase(), contentType);
    } else {
      this.#ops.push(
        entry('Override', [
          ['PartName', name],
          ['ContentType', contentType],
        ])
      );
      this.#overrides.set(key, contentType);
    }
    this.#typesLength += 1;
  }

  /** The destination notes master: this package's, else a copy of the source's. */
  #notesMasterFor(src: string): string {
    if (this.#notesMaster) return this.#notesMaster;
    this.#notesMaster = this.copy(src);
    this.#newNotesMaster = true;
    return this.#notesMaster;
  }

  /** Where a relationship of a copied part points now, or null to drop it. */
  #decide(from: string, role: Role, type: string, target: string | undefined): string | null {
    // A relationships part is never a relationship target (ECMA-376 Part 2, 9.3).
    if (!target || RELS_PART.test(target) || !this.source.parts.has(target)) return null;
    if (REFUSED.has(type))
      throw new CommandRefusal(
        'disabled.slideNotCopyable',
        type.endsWith('control') ? 'import-control' : 'import-comments'
      );
    // A structural relationship is followed only to a part of the role that its type names.
    const is = (want: Role): boolean => roleOf(this.source, target) === want;
    switch (type) {
      case REL.slideLayout:
        if (!is('layout')) return null;
        if (role === 'slide')
          return this.layouts ? (this.layouts.get(from) ?? null) : this.copy(target);
        return role === 'master' && this.keepLayouts.has(target) ? this.copy(target) : null;
      case REL.slideMaster:
        return role === 'layout' && is('master') ? this.copy(target) : null;
      case REL.notesSlide:
        return role === 'slide' && is('notes') ? this.copy(target) : null;
      case REL.notesMaster:
        return role === 'notes' && is('notesMaster') ? this.#notesMasterFor(target) : null;
      case REL.slide:
        return this.slides.get(target) ?? null;
      case REL_COMMENTS:
        return role === 'slide' && contentTypeOf(this.source, target) === CT_COMMENTS
          ? this.copy(target)
          : null;
      case `${R}handoutMaster`:
        return null;
      default:
        // Another type copies only a content part: never a slide, layout, master, or notes part,
        // and never a part that the package has once (the presentation, its properties).
        return is('other') && !packageLevel(contentTypeOf(this.source, target))
          ? this.copy(target)
          : null;
    }
  }

  /** Write the relationships part of each copied part, copying what they name. */
  run(): void {
    for (let src = this.#pending.shift(); src !== undefined; src = this.#pending.shift()) {
      const rels = xmlPartOf(this.source, relsPartName(src));
      if (!rels) continue;
      const dest = this.#copies.get(src)!;
      const role = roleOf(this.source, src);
      const items: string[] = [];
      for (const r of elements(rels.root)) {
        const id = attr(r, 'Id') ?? '';
        const type = attr(r, 'Type') ?? '';
        const raw = attr(r, 'Target') ?? '';
        const external = attr(r, 'TargetMode') === 'External';
        let target = raw;
        if (!external) {
          const resolved = resolveInternalTarget(src, raw);
          const to = this.#decide(src, role, type, resolved.ok ? resolved.partName : undefined);
          if (to === null) {
            const lost = this.#dropped.get(dest) ?? new Set<string>();
            lost.add(id);
            this.#dropped.set(dest, lost);
            continue;
          }
          target = relativeTarget(dest, to);
        }
        items.push(
          `<Relationship Id="${xmlAttr(id, 'relationship')}" Type="${xmlAttr(type, 'relationship')}" Target="${xmlAttr(target, 'relationship')}"${external ? ' TargetMode="External"' : ''}/>`
        );
      }
      this.#ops.push({
        op: 'addPart',
        part: relsPartName(dest),
        xml: `${XML_DECL}<Relationships xmlns="${NS.rel}">${items.join('')}</Relationships>`,
      });
    }
  }

  /** Remove the link actions whose slide was not imported. */
  dropLinks(): void {
    for (const [part, ids] of this.#dropped) {
      const root = xmlPartOf(this.batch.pkg, part)?.root;
      if (!root) continue;
      const users: Element[] = [];
      const walk = (el: Element, depth: number): void => {
        if (depth > 256) return;
        if (
          el.ns === NS.a &&
          HYPERLINKS.has(el.local) &&
          el.attrs.some((a) => a.ns === NS.r && ids.has(a.value))
        )
          users.push(el);
        else for (const c of elements(el)) walk(c, depth + 1);
      };
      walk(root, 0);
      for (const el of users) this.batch.push({ op: 'removeNode', part, node: el.id });
    }
  }

  /** List each copied master in the presentation, with new ids for it and its layouts. */
  registerMasters(): void {
    const presentation = presentationInfo(this.batch.pkg).presentationPart;
    for (const master of this.#masters) {
      const list = path(xmlPartOf(this.batch.pkg, presentation)?.root, 'p:sldMasterIdLst');
      if (!list) throw new CommandRefusal('errors.malformed', 'sldMasterIdLst');
      // Ids are unique among every master and layout of the presentation.
      let next = MIN_MASTER_ID - 1;
      for (const m of elements(list)) {
        next = Math.max(next, Number(attr(m, 'id')) || 0);
        const rel = relationshipsOf(this.batch.pkg, presentation).find(
          (r) => r.id === attr(m, 'r:id')
        );
        for (const l of elements(
          child(xmlPartOf(this.batch.pkg, rel?.target)?.root, 'p:sldLayoutIdLst')
        ))
          next = Math.max(next, Number(attr(l, 'id')) || 0);
      }
      const take = (): string => {
        next += 1;
        if (next > MAX_MASTER_ID) throw new CommandRefusal('errors.tooLarge', 'master id');
        return String(next);
      };
      const own = new Set(relationshipsOf(this.batch.pkg, master).map((r) => r.id));
      const layouts = child(xmlPartOf(this.batch.pkg, master)?.root, 'p:sldLayoutIdLst');
      const masterId = take();
      for (const l of elements(layouts)) {
        if (!own.has(attr(l, 'r:id') ?? '')) {
          this.batch.push({ op: 'removeNode', part: master, node: l.id });
          continue;
        }
        this.batch.push({
          op: 'setAttribute',
          part: master,
          node: l.id,
          name: { ns: '', local: 'id' },
          value: take(),
        });
      }
      const rId = addRelationship(this.batch, presentation, REL.slideMaster, master);
      const fresh = path(xmlPartOf(this.batch.pkg, presentation)?.root, 'p:sldMasterIdLst')!;
      this.batch.insert({
        op: 'insertNode',
        part: presentation,
        parent: fresh.id,
        index: fresh.children.length,
        spec: specFromXml(`<p:sldMasterId id="${masterId}" r:id="${xmlAttr(rId, 'rId')}"/>`),
      });
    }
    if (this.#newNotesMaster && this.#notesMaster) this.#registerNotesMaster(presentation);
  }

  /** `p:notesMasterIdLst` comes right after `p:sldMasterIdLst` (19.2.1.26). */
  #registerNotesMaster(presentation: string): void {
    const rId = addRelationship(this.batch, presentation, REL.notesMaster, this.#notesMaster!);
    const root = xmlPartOf(this.batch.pkg, presentation)!.root;
    const entry = `<p:notesMasterId r:id="${xmlAttr(rId, 'rId')}"/>`;
    const list = child(root, 'p:notesMasterIdLst');
    if (list) {
      this.batch.insert({
        op: 'insertNode',
        part: presentation,
        parent: list.id,
        index: 0,
        spec: specFromXml(entry),
      });
      return;
    }
    const at = root.children.findIndex((c) => c.type === 'element' && c.local === 'sldMasterIdLst');
    this.batch.insert({
      op: 'insertNode',
      part: presentation,
      parent: root.id,
      index: at + 1,
      spec: specFromXml(`<p:notesMasterIdLst>${entry}</p:notesMasterIdLst>`),
    });
  }
}

/**
 * The destination layout for a source layout: the first of the same type (a custom layout: of
 * the same name), else the first layout of the first master.
 */
function destinationLayout(
  sourceLayouts: ReturnType<typeof deckLayouts>,
  all: ReturnType<typeof deckLayouts>,
  layout: string | undefined
): string {
  if (all.length === 0) throw new CommandRefusal('errors.layoutNotFound');
  const from = sourceLayouts.find((l) => l.part === layout);
  const match = from
    ? all.find((l) =>
        from.type === 'cust' ? l.type === 'cust' && l.name === from.name : l.type === from.type
      )
    : undefined;
  return (match ?? all[0]!).part;
}

/** Give each placeholder of a copied slide the box it showed in the source as its own. */
function freezePlaceholders(
  batch: OpBatch,
  part: string,
  source: PptxPackage,
  index: number
): void {
  const ref = presentationInfo(source).slides[index];
  const xml = xmlPartOf(source, ref?.partName);
  if (!ref || !xml) return;
  const nodes = nodeIndex(xml);
  const boxes = new Map<string, Xfrm>();
  for (const e of resolveSlide(source, index)?.elements ?? []) {
    if (e.layer !== 'slide') continue;
    const node = nodes.get(e.nodeId)?.node;
    const id = node?.type === 'element' ? attr(cNvPr(node), 'id') : undefined;
    if (id !== undefined && !boxes.has(id)) boxes.set(id, e.xfrm);
  }
  for (const el of elements(shapeTreeOf(batch.pkg, part))) {
    const nv = elements(el).find((c) => c.local.startsWith('nv') && c.local.endsWith('Pr'));
    if (!path(nv, 'p:nvPr', 'p:ph')) continue;
    const id = attr(cNvPr(el), 'id');
    if (id !== undefined) keepBox(batch, part, el, boxes.get(id));
  }
}

/**
 * List the copied slides in the presentation after slide `after`, in one pass: one relationship,
 * one content type override, and one `p:sldId` each (ECMA-376 Part 1, 19.2.1.33), with ids above
 * every id in use. A list that would need a gap search goes slide by slide.
 */
function registerImported(batch: OpBatch, names: readonly string[], after: number): string[] {
  const info = presentationInfo(batch.pkg);
  const presentation = info.presentationPart;
  const root = xmlPartOf(batch.pkg, presentation)?.root;
  const list = path(root, 'p:sldIdLst');
  const rels = xmlPartOf(batch.pkg, relsPartName(presentation));
  const types = xmlPartOf(batch.pkg, CONTENT_TYPES_PART);
  let max = MIN_SLIDE_ID - 1;
  for (const s of elements(list)) {
    const id = Number(attr(s, 'id'));
    if (Number.isInteger(id) && id > max && id <= MAX_SLIDE_ID) max = id;
  }
  if (!root || !list || !rels || !types || max + names.length > MAX_SLIDE_ID)
    return names.map((name, i) => registerSlide(batch, name, after + 1 + i));
  // Relationship ids: every id of the relationships part and every r:* the part names is taken.
  const taken = new Set<string>();
  let top = 0;
  const take = (value: string | undefined): void => {
    if (!value) return;
    taken.add(value);
    const m = /^rId(\d{1,9})$/.exec(value);
    if (m) top = Math.max(top, Number(m[1]));
  };
  for (const r of elements(rels.root)) take(attr(r, 'Id'));
  const stack: Element[] = [root];
  for (let el = stack.pop(); el; el = stack.pop()) {
    for (const a of el.attrs) if (a.ns === NS.r) take(a.value);
    for (const c of el.children) if (c.type === 'element') stack.push(c);
  }
  const entries = elements(list);
  const before = info.slides[after + 1]?.nodeId;
  const at = before === undefined ? -1 : entries.findIndex((e) => e.id === before);
  const base = at >= 0 ? at : entries.length;
  const ops: PptxOp[] = [];
  const ids: string[] = [];
  let n = top;
  names.forEach((name, i) => {
    do n += 1;
    while (taken.has(`rId${n}`));
    const rId = `rId${n}`;
    const id = String(max + 1 + i);
    ids.push(id);
    ops.push(
      {
        op: 'insertNode',
        part: relsPartName(presentation),
        parent: rels.root.id,
        index: rels.root.children.length + i,
        spec: {
          ns: NS.rel,
          local: 'Relationship',
          attrs: [
            { ns: '', local: 'Id', value: rId },
            { ns: '', local: 'Type', value: REL.slide },
            { ns: '', local: 'Target', value: relativeTarget(presentation, name) },
          ],
        },
      },
      {
        op: 'insertNode',
        part: CONTENT_TYPES_PART,
        parent: types.root.id,
        index: types.root.children.length + i,
        spec: {
          ns: NS.ct,
          local: 'Override',
          attrs: [
            { ns: '', local: 'PartName', value: name },
            { ns: '', local: 'ContentType', value: CONTENT_TYPE.slide },
          ],
        },
      },
      {
        op: 'insertNode',
        part: presentation,
        parent: list.id,
        index: base + i,
        spec: {
          ns: NS.p,
          local: 'sldId',
          prefix: 'p',
          attrs: [
            { ns: '', local: 'id', value: id },
            { ns: NS.r, local: 'id', prefix: 'r', value: rId },
          ],
        },
      }
    );
  });
  batch.pushAll(ops);
  return ids;
}

/**
 * Refuse the import when a slide to import has threaded comments or an embedded control,
 * naming the slide ids, so the caller can leave those slides out with `sourceSlideIds`.
 */
function refuseNotCopyable(
  source: PptxPackage,
  refs: readonly { readonly partName: string; readonly sldId: string }[]
): void {
  const comments: string[] = [];
  const controls: string[] = [];
  for (const ref of refs)
    for (const r of relationshipsOf(source, ref.partName)) {
      if (r.external || !REFUSED.has(r.type)) continue;
      const into = r.type.endsWith('control') ? controls : comments;
      if (into.at(-1) !== ref.sldId && /^\d{1,10}$/.test(ref.sldId)) into.push(ref.sldId);
    }
  const ids = comments.length > 0 ? comments : controls;
  if (ids.length === 0) return;
  throw new CommandRefusal(
    'disabled.slideNotCopyable',
    `${comments.length > 0 ? 'import-comments' : 'import-control'}:${ids.slice(0, MAX_NAMED).join(',')}`
  );
}

/** The most slide ids one refusal names. */
const MAX_NAMED = 20;

/** The value for a key that the import put there, or a refusal: never an unchecked lookup. */
function found(map: ReadonlyMap<string, string>, key: string): string {
  const value = map.get(key);
  if (value === undefined) throw new CommandRefusal('errors.malformed', `import ${key}`);
  return value;
}

/** `ST_SlideId` (ECMA-376 Part 1, 19.7.13). */
const MIN_SLIDE_ID = 256;
const MAX_SLIDE_ID = 2_147_483_647;

/**
 * Import slides of another presentation after slide `after`, in one batch. Returns the new
 * slide ids, in order. Refuses a slide with threaded comments or controls (`disabled.slideNotCopyable`).
 */
export function importSlides(batch: OpBatch, request: SlideImport): string[] {
  const { source } = request;
  const sourceSlides = presentationInfo(source).slides;
  const refs = request.slides.map((i) => {
    const ref = sourceSlides[i];
    if (!ref || !xmlPartOf(source, ref.partName))
      throw new CommandRefusal('errors.malformed', 'import slide');
    return ref;
  });
  if (refs.length === 0) return [];
  refuseNotCopyable(source, refs);
  // Two list entries for one slide part would make one copy with two ids.
  if (new Set(refs.map((r) => partNameKey(r.partName))).size !== refs.length)
    throw new CommandRefusal('errors.malformed', 'import slide');
  const keep = request.formatting === 'keepSource';
  const layouts = keep ? null : new Map<string, string>();
  const keepLayouts = new Set<string>();
  const sourceLayouts = keep ? [] : deckLayouts(source);
  const destLayouts = keep ? [] : deckLayouts(batch.pkg);
  for (const ref of refs) {
    if (ref.layoutPart) keepLayouts.add(ref.layoutPart);
    layouts?.set(ref.partName, destinationLayout(sourceLayouts, destLayouts, ref.layoutPart));
  }
  if (keep && refs.some((r) => !r.layoutPart || roleOf(source, r.layoutPart) !== 'layout'))
    throw new CommandRefusal('errors.malformed', 'import layout');
  const importer = new Importer(batch, source, layouts, keepLayouts);
  for (const ref of refs) importer.addSlide(ref.partName);
  const copies = importer.slides;
  importer.run();
  importer.flush();
  mergeImportedComments(batch, source, importer.comments);
  importer.registerMasters();
  importer.dropLinks();

  const info = presentationInfo(batch.pkg);
  const names = refs.map((ref) => found(copies, ref.partName));
  const used = creationIdsOf(
    batch,
    info.slides.map((s) => s.partName)
  );
  const fresh: PptxOp[] = [];
  for (const name of names) freshCreationIdAmong(batch, name, used, fresh);
  batch.pushAll(fresh);
  const ids = registerImported(batch, names, request.after);
  const firstSlide = info.slides[0]?.sldId;
  let previous = request.after >= 0 ? info.slides[request.after]?.sldId : undefined;
  if (hasSections(batch, info.presentationPart))
    for (const id of ids) {
      if (previous) placeInSection(batch, info.presentationPart, id, previous, 'after');
      else if (firstSlide) placeInSection(batch, info.presentationPart, id, firstSlide, 'before');
      previous = id;
    }
  if (layouts)
    refs.forEach((ref, i) => {
      const at = request.after + 1 + i;
      freezePlaceholders(batch, found(copies, ref.partName), source, request.slides[i] ?? -1);
      applyLayout(batch, at, found(layouts, ref.partName));
    });
  return ids;
}
