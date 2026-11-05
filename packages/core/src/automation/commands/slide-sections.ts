// Sections: named runs of slides that the slide pane groups (Home > Section). The
// list is `p14:sectionLst` (the `p14` extension schema, `CT_SectionList`) in the presentation's
// `p:extLst`, under the extension URI below; each `p14:section` names its slides by `p:sldId/@id`
// in slide order. Add, rename, remove, and move a section; each command is one transaction.
// New, duplicated, moved, and deleted slides keep the list in order (`slide-structure.ts`).

import {
  attr,
  canonicalSlideId,
  child,
  elements,
  NS,
  presentationInfo,
  repairedPresentationRoot,
  xmlPartOf,
  type Element,
  type PptxPackage,
} from '../../store/index.ts';
import type { SlideSection } from '../../contracts/neutral.ts';
import type { NodeSpec } from '../../store/index.ts';
import { CommandRefusal, type OpBatch } from './op-batch.ts';
import { xmlAttr } from './package-edits.ts';
import { specFromXml } from './shape-commands.ts';
import { deleteSlides, reorderSlides } from './slide-multi.ts';
import { cleanText } from './text-units.ts';

/** The `p:ext/@uri` of the section list (the `p14` extension schema, `CT_SectionList`). */
export const SECTION_LIST_URI = '{521415D9-36F7-43E2-AB2F-B90AF26B5E84}';

/**
 * One section of the deck: its element, its section list, and the slide ids it names. The ids
 * are canonical (`canonicalSlideId`): "0258" in a section and "258" in `p:sldIdLst` name one
 * slide, as in the load and save repairs. Compare them with `sldIdOf`.
 */
interface DeckSection {
  readonly el: Element;
  readonly list: Element;
  readonly ext: Element;
  readonly id: string;
  readonly name: string;
  readonly slideIds: readonly string[];
}

/** The `p:ext` that holds the section list, and the list. */
function sectionList(pkg: PptxPackage): { ext: Element; list: Element } | undefined {
  const root = xmlPartOf(pkg, presentationInfo(pkg).presentationPart)?.root;
  for (const ext of elements(child(root, 'p:extLst'))) {
    if (attr(ext, 'uri') !== SECTION_LIST_URI) continue;
    const list = elements(ext).find((e) => e.ns === NS.p14 && e.local === 'sectionLst');
    if (list) return { ext, list };
  }
  return undefined;
}

/**
 * Every section of the deck, in order. Two collaborators who add a section at the same time can
 * leave two section lists in the part (the writer joins them, `repairPresentationRoot`); they
 * read as one list here, a section id once.
 */
function sectionsOf(
  pkg: PptxPackage,
  root = xmlPartOf(pkg, presentationInfo(pkg).presentationPart)?.root
): DeckSection[] {
  const out: DeckSection[] = [];
  const ids = new Set<string>();
  for (const extLst of elements(root).filter((e) => e.ns === NS.p && e.local === 'extLst'))
    for (const ext of elements(extLst)) {
      if (attr(ext, 'uri') !== SECTION_LIST_URI) continue;
      const list = elements(ext).find((e) => e.ns === NS.p14 && e.local === 'sectionLst');
      for (const el of elements(list)) {
        const id = attr(el, 'id') ?? '';
        if (el.ns !== NS.p14 || el.local !== 'section' || ids.has(id)) continue;
        ids.add(id);
        out.push({
          el,
          list: list!,
          ext,
          id,
          name: attr(el, 'name') ?? '',
          slideIds: elements(elements(el).find((e) => e.local === 'sldIdLst')).map((s) =>
            canonicalSlideId(attr(s, 'id') ?? '')
          ),
        });
      }
    }
  return out;
}

/** The canonical slide id of a slide, to compare with `DeckSection.slideIds`. */
const sldIdOf = (slide: { readonly sldId: string }): string => canonicalSlideId(slide.sldId);

const summaryCache = new WeakMap<PptxPackage, readonly SlideSection[]>();

/**
 * The deck's sections for the snapshot: name, first slide, and slide count, in order. A slide
 * id that no slide has is left out. Memoized per package identity.
 */
export function deckSections(pkg: PptxPackage): readonly SlideSection[] {
  const hit = summaryCache.get(pkg);
  if (hit) return hit;
  const index = new Map(presentationInfo(pkg).slides.map((s, i) => [sldIdOf(s), i]));
  // A slide counts once, in the first section that names it: a file can repeat an id.
  const seen = new Set<number>();
  let next = 0;
  // The sections as the writer saves them, which slide and section commands act on.
  const out = sectionsOf(pkg, repairedPresentationRoot(pkg)).map((s): SlideSection => {
    let first = Infinity;
    let last = -1;
    let count = 0;
    for (const id of s.slideIds) {
      const i = index.get(id);
      if (i === undefined || seen.has(i)) continue;
      seen.add(i);
      count += 1;
      if (i < first) first = i;
      if (i > last) last = i;
    }
    const firstSlide = count ? first : next;
    next = count ? last + 1 : next;
    return { id: s.id, name: s.name, firstSlide, slideCount: count };
  });
  summaryCache.set(pkg, out);
  return out;
}

/**
 * Bring the presentation part in `batch` to what the writer saves (`repairPresentationRoot`):
 * one slide entry for each slide, one section list, and its sections in slide order. A
 * concurrent edit can leave the part otherwise. Slide and section commands run
 * this first, so they act on the sections that the rail shows. The existing nodes move in
 * place and only the extras go, so an edit that another replica makes at the same time on a
 * section or an extension still lands.
 */
export function normalizePresentation(batch: OpBatch): void {
  const part = presentationInfo(batch.pkg).presentationPart;
  const root = xmlPartOf(batch.pkg, part)?.root;
  const fixed = repairedPresentationRoot(batch.pkg);
  if (!root || !fixed || fixed === root) return;
  const is = (e: Element, local: string): boolean => e.local === local;
  /** Move node `id` under `parent`, right after `prev` (first when null). */
  const place = (id: string, parent: string, prev: string | null): void => {
    const now = batch.element(part, parent).children;
    const others = now.filter((c) => c.id !== id);
    const index = prev === null ? 0 : others.findIndex((c) => c.id === prev) + 1;
    if (now.findIndex((c) => c.id === id) !== index)
      batch.push({ op: 'moveNode', part, node: id, parent, index });
  };
  // Each slide entry that stays, with its slide id: a repeated or invalid id is renewed.
  const fixedList = elements(fixed).find((e) => e.ns === NS.p && is(e, 'sldIdLst'));
  const keep = new Map(elements(fixedList).map((e) => [e.id, attr(e, 'id')]));
  const slideLists = elements(root).filter((e) => e.ns === NS.p && is(e, 'sldIdLst'));
  for (const entry of slideLists.flatMap((list) => elements(list))) {
    const id = keep.get(entry.id);
    if (!keep.has(entry.id)) batch.push({ op: 'removeNode', part, node: entry.id });
    else if (id !== undefined && id !== attr(entry, 'id'))
      batch.push({
        op: 'setAttribute',
        part,
        node: entry.id,
        name: { ns: '', local: 'id' },
        value: id,
      });
  }
  // Two replicas that each add the first slide each create a `p:sldIdLst`; `CT_Presentation`
  // allows one (ECMA-376 Part 1, 19.2.1.26). The entries of the others move into the first, in
  // document order, and the other lists go.
  if (slideLists.length > 1 && fixedList) {
    let prev: string | null = null;
    for (const entry of elements(fixedList)) {
      place(entry.id, fixedList.id, prev);
      prev = entry.id;
    }
    for (const other of slideLists.slice(1)) batch.push({ op: 'removeNode', part, node: other.id });
  }
  const lists = elements(root).filter((e) => e.ns === NS.p && is(e, 'extLst'));
  const target = elements(fixed).find((e) => e.ns === NS.p && is(e, 'extLst'));
  // A part without extension lists needed only the slide entries.
  if (!lists.length || (lists.length === 1 && lists[0] === target)) return;
  const exts = elements(target);
  const sectionExt = exts.find((e) => attr(e, 'uri') === SECTION_LIST_URI);
  const list = elements(sectionExt).find((e) => e.ns === NS.p14 && is(e, 'sectionLst'));
  // The sections, into the kept list in slide order.
  let prev: string | null = null;
  for (const section of elements(list)) {
    place(section.id, list!.id, prev);
    prev = section.id;
  }
  // Then their slides: an entry that a section already holds moves where it belongs, also from
  // another section, so two replicas that repair at the same time keep one entry per slide.
  if (list) {
    const pool = new Map<string, string>();
    const extra: string[] = [];
    for (const section of elements(batch.element(part, list.id)))
      for (const e of elements(elements(section).find((c) => is(c, 'sldIdLst')))) {
        // Slide ids compare as numbers ("0258" is 258), as the writer's repair does.
        const id = canonicalSlideId(attr(e, 'id') ?? '');
        if (pool.has(id)) extra.push(e.id);
        else pool.set(id, e.id);
      }
    for (const section of elements(list)) fillSection(batch, part, section, pool, place);
    for (const node of [...extra, ...pool.values()]) batch.push({ op: 'removeNode', part, node });
  }
  // The extensions, into the first extension list.
  const first = lists[0]!;
  prev = null;
  for (const ext of exts) {
    place(ext.id, first.id, prev);
    prev = ext.id;
  }
  // What the repair left out goes: the other lists, extensions, and sections.
  const wanted = new Set(exts.map((e) => e.id));
  for (const other of lists.slice(1)) batch.push({ op: 'removeNode', part, node: other.id });
  for (const ext of elements(batch.element(part, first.id)))
    if (!wanted.has(ext.id)) batch.push({ op: 'removeNode', part, node: ext.id });
  if (!target) batch.push({ op: 'removeNode', part, node: first.id });
  if (!list) return;
  const sections = new Set(elements(list).map((e) => e.id));
  for (const section of elements(batch.element(part, list.id)))
    if (!sections.has(section.id)) batch.push({ op: 'removeNode', part, node: section.id });
}

/**
 * Give section `want.id` the slide entries of `want`, in order: an entry from `pool` (slide id to
 * node, taken when used) moves into place, and only a slide with no entry gets a new one.
 */
function fillSection(
  batch: OpBatch,
  part: string,
  want: Element,
  pool: Map<string, string>,
  place: (id: string, parent: string, prev: string | null) => void
): void {
  const ids = elements(elements(want).find((e) => e.local === 'sldIdLst')).map(
    (e) => attr(e, 'id') ?? ''
  );
  const section = batch.element(part, want.id);
  const entry = (id: string) => ({
    ns: NS.p14,
    local: 'sldId',
    prefix: section.prefix ?? 'p14',
    attrs: [{ ns: '', local: 'id', value: id }],
  });
  let own = elements(section).find((e) => e.local === 'sldIdLst')?.id;
  // `p14:sldIdLst` is the first child of a section (the `p14` schema, `CT_Section`).
  own ??= batch.insert({
    op: 'insertNode',
    part,
    parent: section.id,
    index: 0,
    spec: { ...entry(''), local: 'sldIdLst', attrs: [] },
  });
  let prev: string | null = null;
  for (const id of ids) {
    const node = pool.get(canonicalSlideId(id));
    if (node) {
      pool.delete(canonicalSlideId(id));
      place(node, own, prev);
      prev = node;
    } else {
      const now = batch.element(part, own).children;
      const index = prev === null ? 0 : now.findIndex((c) => c.id === prev) + 1;
      prev = batch.insert({ op: 'insertNode', part, parent: own, index, spec: entry(id) });
    }
  }
}

/**
 * Section summaries for the snapshot, reference-stable while the sections stay the same.
 */
export function createSectionSummaries(): (
  pkg: PptxPackage | undefined
) => readonly SlideSection[] {
  let last: readonly SlideSection[] = [];
  return (pkg) => {
    const next = pkg ? deckSections(pkg) : [];
    const same =
      next.length === last.length &&
      next.every(
        (s, i) =>
          s.id === last[i].id &&
          s.name === last[i].name &&
          s.firstSlide === last[i].firstSlide &&
          s.slideCount === last[i].slideCount
      );
    if (!same) last = next;
    return last;
  };
}

/** A section name as the editor keeps it: one line, no control characters, 1 to 255 long. */
function sectionName(raw: string): string {
  const name = cleanText(raw)
    .replace(/[\n\u000B]/g, ' ')
    .trim();
  if (!name || name.length > 255) throw new CommandRefusal('errors.sectionName');
  return name;
}

/**
 * A random version 4 GUID in upper case. `crypto.getRandomValues` also works in an insecure
 * context (a page served over plain HTTP), where `crypto.randomUUID` does not exist.
 */
function randomGuid(): string {
  const b = globalThis.crypto.getRandomValues(new Uint8Array(16));
  b[6] = (b[6]! & 0x0f) | 0x40;
  b[8] = (b[8]! & 0x3f) | 0x80;
  let hex = '';
  for (const x of b) hex += x.toString(16).padStart(2, '0');
  hex = hex.toUpperCase();
  return `${hex.slice(0, 8)}-${hex.slice(8, 12)}-${hex.slice(12, 16)}-${hex.slice(16, 20)}-${hex.slice(20)}`;
}

/** A new section id: a GUID in braces (`ST_Guid`). */
function freshSectionId(pkg: PptxPackage): string {
  const used = new Set(sectionsOf(pkg).map((s) => s.id.toUpperCase()));
  for (;;) {
    const id = `{${randomGuid()}}`;
    if (!used.has(id)) return id;
  }
}

/** A `p14:section`; `declare` false leaves out `xmlns:p14` for a list that declares it. */
function sectionXml(id: string, name: string, slideIds: readonly string[], declare = true): string {
  const ids = slideIds.map((s) => `<p14:sldId id="${xmlAttr(s, 'slide id')}"/>`).join('');
  const ns = declare ? ` xmlns:p14="${NS.p14}"` : '';
  return (
    `<p14:section${ns} name="${xmlAttr(name, 'section name')}" id="${id}">` +
    `<p14:sldIdLst>${ids}</p14:sldIdLst></p14:section>`
  );
}

/**
 * The spec of a new section for section list `list`: without its own `xmlns:p14` when the list
 * declares `p14` already, so the saved part declares the namespace once.
 */
function sectionSpec(list: Element, xml: string): NodeSpec {
  return withoutDeclaredP14(list, specFromXml(xml));
}

function withoutDeclaredP14<T extends NodeSpec | { readonly text: string }>(
  list: Element,
  spec: T
): T {
  if ('text' in spec) return spec;
  const declared =
    list.prefix === 'p14' && list.nsDecls.some((d) => d.prefix === 'p14' && d.uri === NS.p14);
  if (!declared || !spec.nsDecls?.length) return spec;
  return {
    ...spec,
    nsDecls: spec.nsDecls.filter((d) => !(d.prefix === 'p14' && d.uri === NS.p14)),
  };
}

function findSection(pkg: PptxPackage, sectionId: string): { all: DeckSection[]; at: number } {
  const all = sectionsOf(pkg);
  const at = all.findIndex((s) => s.id === sectionId);
  if (at < 0) throw new CommandRefusal('errors.sectionNotFound');
  return { all, at };
}

/**
 * Add a section that starts at slide `index` and runs to the end of the section it splits.
 * In a deck without sections, the slides before `index` go into
 * a first section named `defaultName` ("Default Section" in English).
 */
export function addSection(
  batch: OpBatch,
  index: number,
  name: string,
  defaultName: string
): string {
  const info = presentationInfo(batch.pkg);
  const slide = info.slides[index];
  if (!slide) throw new CommandRefusal('errors.slideNotFound');
  const id = freshSectionId(batch.pkg);
  const label = sectionName(name);
  const found = sectionList(batch.pkg);
  const all = sectionsOf(batch.pkg);
  if (!all.length) {
    const before = info.slides.slice(0, index).map((s) => s.sldId);
    const after = info.slides.slice(index).map((s) => s.sldId);
    const firstId = before.length ? freshSectionId(batch.pkg) : '';
    // A new section list declares `p14` once; an existing one may declare it elsewhere.
    const sections = (declare: boolean): string =>
      (before.length ? sectionXml(firstId, sectionName(defaultName), before, declare) : '') +
      sectionXml(id, label, after, declare);
    if (found) {
      // An empty section list: the sections go into it.
      const specs = specFromXml(
        `<p14:sectionLst xmlns:p14="${NS.p14}">${sections(true)}</p14:sectionLst>`
      );
      (specs.children ?? []).forEach((spec, i) => {
        if ('text' in spec) return;
        batch.push({
          op: 'insertNode',
          part: info.presentationPart,
          parent: found.list.id,
          index: found.list.children.length + i,
          spec: withoutDeclaredP14(found.list, spec),
        });
      });
      return id;
    }
    const ext =
      `<p:ext uri="${SECTION_LIST_URI}"><p14:sectionLst xmlns:p14="${NS.p14}">` +
      `${sections(false)}</p14:sectionLst></p:ext>`;
    const root = xmlPartOf(batch.pkg, info.presentationPart)!.root;
    const extLst = child(root, 'p:extLst');
    if (extLst)
      batch.push({
        op: 'insertNode',
        part: info.presentationPart,
        parent: extLst.id,
        index: 0,
        spec: specFromXml(ext),
      });
    else
      batch.push({
        op: 'insertNode',
        part: info.presentationPart,
        parent: root.id,
        // `p:extLst` is the last child of `p:presentation` (`CT_Presentation`).
        index: root.children.length,
        spec: specFromXml(`<p:extLst>${ext}</p:extLst>`),
      });
    return id;
  }
  const owner = all.find((s) => s.slideIds.includes(sldIdOf(slide)));
  if (!owner) {
    // A list that misses the slide: the new section takes it and the unlisted slides after it,
    // and goes after the section of the listed slide before it (first when there is none). It
    // writes the slides' raw `p:sldId/@id` values and compares them canonically.
    const listed = new Set(all.flatMap((s) => s.slideIds));
    const run: string[] = [];
    for (let i = index; i < info.slides.length && !listed.has(sldIdOf(info.slides[i]!)); i += 1)
      run.push(info.slides[i]!.sldId);
    const before = info.slides
      .slice(0, index)
      .reverse()
      .find((s) => listed.has(sldIdOf(s)));
    const prev = before ? all.find((s) => s.slideIds.includes(sldIdOf(before))) : undefined;
    const anchor = prev ?? all[0]!;
    batch.push({
      op: 'insertNode',
      part: info.presentationPart,
      parent: anchor.list.id,
      index: anchor.list.children.indexOf(anchor.el) + (prev ? 1 : 0),
      spec: sectionSpec(anchor.list, sectionXml(id, label, run)),
    });
    return id;
  }
  // The new section takes the owner's slides from `index` on, in slide order, whatever order
  // the list names them in. It writes each slide's raw `p:sldId/@id`, so "0258" stays "0258".
  const owned = new Set(owner.slideIds);
  const movingSlides = info.slides.slice(index).filter((s) => owned.has(sldIdOf(s)));
  const moving = movingSlides.map((s) => s.sldId);
  const leaving = new Set(movingSlides.map(sldIdOf));
  const list = elements(owner.el).find((e) => e.local === 'sldIdLst');
  for (const el of elements(list))
    if (leaving.has(canonicalSlideId(attr(el, 'id') ?? '')))
      batch.push({ op: 'removeNode', part: info.presentationPart, node: el.id });
  batch.push({
    op: 'insertNode',
    part: info.presentationPart,
    parent: owner.list.id,
    index: owner.list.children.indexOf(owner.el) + 1,
    spec: sectionSpec(owner.list, sectionXml(id, label, moving)),
  });
  return id;
}

/** Rename a section. */
export function renameSection(batch: OpBatch, sectionId: string, name: string): void {
  const { all, at } = findSection(batch.pkg, sectionId);
  batch.push({
    op: 'setAttribute',
    part: presentationInfo(batch.pkg).presentationPart,
    node: all[at].el.id,
    name: { ns: '', local: 'name' },
    value: sectionName(name),
  });
}

/**
 * Remove a section. Its slides join the section before it (the one after it, for the first
 * section); with `withSlides` they are deleted
 * (Remove Section & Slides). Removing the last section removes the section list.
 */
export function removeSection(
  batch: OpBatch,
  sectionId: string,
  withSlides: boolean
): number | undefined {
  const presentation = presentationInfo(batch.pkg).presentationPart;
  const { all, at } = findSection(batch.pkg, sectionId);
  const section = all[at];
  let current: number | undefined;
  if (withSlides && section.slideIds.length) {
    const ids = new Set(section.slideIds);
    const indices = presentationInfo(batch.pkg)
      .slides.map((s, i) => (ids.has(sldIdOf(s)) ? i : -1))
      .filter((i) => i >= 0);
    current = deleteSlides(batch, indices);
  }
  const fresh = sectionsOf(batch.pkg);
  const self = fresh.find((s) => s.id === sectionId)!;
  if (fresh.length === 1) {
    // The last section: its extension goes, and the extension list when nothing else is in it.
    const extLst = xmlPartOf(batch.pkg, presentation)!
      .root.children.filter((c): c is Element => c.type === 'element' && c.local === 'extLst')
      .find((l) => l.children.includes(self.ext));
    const node = extLst && elements(extLst).length === 1 ? extLst : self.ext;
    batch.push({ op: 'removeNode', part: presentation, node: node.id });
    return current;
  }
  const k = fresh.indexOf(self);
  const target = k > 0 ? fresh[k - 1] : fresh[k + 1];
  const list = elements(target.el).find((e) => e.local === 'sldIdLst');
  const own = elements(elements(self.el).find((e) => e.local === 'sldIdLst'));
  if (list)
    own.forEach((el, i) =>
      batch.push({
        op: 'moveNode',
        part: presentation,
        node: el.id,
        parent: list.id,
        index: k > 0 ? list.children.length + i : i,
      })
    );
  batch.push({ op: 'removeNode', part: presentation, node: self.el.id });
  return current;
}

/**
 * Move a section and its slides one section up or down (Move Section Up and
 * Move Section Down). Returns the new index of the first slide that moved, if any.
 */
export function moveSection(
  batch: OpBatch,
  sectionId: string,
  direction: -1 | 1
): number | undefined {
  if (direction !== -1 && direction !== 1) throw new CommandRefusal('errors.unsupported');
  const presentation = presentationInfo(batch.pkg).presentationPart;
  const { all, at } = findSection(batch.pkg, sectionId);
  const other = all[at + direction];
  if (!other) throw new CommandRefusal('disabled.sectionEdge');
  const [first, second] = direction < 0 ? [all[at], other] : [other, all[at]];
  // Slides: the block of `first` goes before the block of `second`.
  const slides = presentationInfo(batch.pkg).slides;
  const nodeOf = new Map(slides.map((s) => [sldIdOf(s), s.nodeId]));
  const order = slides.map(sldIdOf);
  const inSecond = new Set(second.slideIds);
  const inFirst = new Set(first.slideIds);
  const a = order.filter((id) => inSecond.has(id));
  const b = order.filter((id) => inFirst.has(id) && !inSecond.has(id));
  const moving = new Set([...a, ...b]);
  const start = order.findIndex((id) => moving.has(id));
  const rest = order.filter((id) => !moving.has(id));
  const head = start < 0 ? rest : rest.filter((id) => order.indexOf(id) < start);
  const tail = start < 0 ? [] : rest.filter((id) => order.indexOf(id) >= start);
  reorderSlides(
    batch,
    [...head, ...b, ...a, ...tail].map((id) => nodeOf.get(id)!)
  );
  // The section elements: `first` goes before `second`, in `second`'s list.
  const list = batch.element(presentation, second.list.id);
  batch.push({
    op: 'moveNode',
    part: presentation,
    node: first.el.id,
    parent: list.id,
    index: list.children.findIndex((c) => c.id === second.el.id),
  });
  const moved = all[at].slideIds[0];
  const now = presentationInfo(batch.pkg).slides.findIndex((s) => sldIdOf(s) === moved);
  return now >= 0 ? now : undefined;
}

/** Remove every section (Remove All Sections): the slides stay. */
export function removeAllSections(batch: OpBatch): void {
  const presentation = presentationInfo(batch.pkg).presentationPart;
  const all = sectionsOf(batch.pkg);
  if (!all.length) throw new CommandRefusal('disabled.noSections');
  const root = xmlPartOf(batch.pkg, presentation)!.root;
  for (const ext of new Set(all.map((s) => s.ext))) {
    const extLst = root.children.find(
      (c): c is Element => c.type === 'element' && c.local === 'extLst' && c.children.includes(ext)
    );
    const now = extLst ? batch.element(presentation, extLst.id) : undefined;
    const node = now && elements(now).length === 1 ? now : ext;
    batch.push({ op: 'removeNode', part: presentation, node: node.id });
  }
}

/** The id of the section that holds slide `index`, or a refusal when no section does. */
export function sectionAt(pkg: PptxPackage, index: number): string {
  const s = deckSections(pkg).find(
    (x) => index >= x.firstSlide && index < x.firstSlide + x.slideCount
  );
  if (!s)
    throw new CommandRefusal(
      deckSections(pkg).length ? 'errors.sectionNotFound' : 'disabled.noSections'
    );
  return s.id;
}
