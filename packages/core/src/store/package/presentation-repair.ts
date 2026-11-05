// The presentation part as the writer emits it: one `p:extLst` (ECMA-376 Part 1, §19.2.1.26
// `CT_Presentation` allows one), one section-list extension, one section list (`p14:sectionLst`
// `CT_SectionList`) whose sections name every slide of `p:sldIdLst` once, in slide order, and
// only those, and one `p:sldId` for each slide, with a relationship to a part. Concurrent edits
// on two replicas can break this: two collaborators who add a section to a deck without
// sections each create a `p:extLst`; a slide deleted on one replica can stay in a section that the other
// replica added, or come back in `p:sldIdLst` when the other replica moved it; and two
// concurrent moves can leave the sections out of slide order. A file can also arrive that way.
// The writer repairs it, so a presentation app opens the file without a repair prompt. A part that is
// already valid is written as it is, and an extension that is not the section list is kept.

import { NS } from './namespaces.ts';
import { attr, elements } from './query.ts';
import type { Element, Node } from './tree.ts';

/** `p:ext/@uri` of the section list. */
const SECTION_LIST_URI = '{521415D9-36F7-43E2-AB2F-B90AF26B5E84}';

const isExtLst = (n: Node): n is Element =>
  n.type === 'element' && n.ns === NS.p && n.local === 'extLst';
const isSlideList = (n: Node): n is Element =>
  n.type === 'element' && n.ns === NS.p && n.local === 'sldIdLst';
const sectionListOf = (ext: Element): Element | undefined =>
  elements(ext).find((e) => e.ns === NS.p14 && e.local === 'sectionLst');
const slideIdsOf = (section: Element): Element[] =>
  elements(elements(section).find((e) => e.local === 'sldIdLst'));

/**
 * The `p:sldId` elements of every `p:sldIdLst`, in document order. `CT_Presentation` allows one
 * list (ECMA-376 Part 1, 19.2.1.26), but two replicas that each add the first slide to a deck
 * without a list each create one. The repair joins them into the first.
 */
const slideEntries = (root: Element): Element[] =>
  root.children.filter(isSlideList).flatMap((list) => elements(list));

/**
 * The relationships of the presentation part that name a part: their `r:id` values, or a map
 * from each `r:id` to the part it names. With the map, two entries whose `r:id` values name
 * one part are one slide.
 */
export type LiveRelationships = ReadonlySet<string> | ReadonlyMap<string, string>;

/**
 * Whether a `p:sldId` stays: when `live` is given, its own `r:id` must be in it. The slide id
 * alone says nothing: a replica can reuse the slide id of an entry that is dead.
 */
const isLive = (entry: Element, live: LiveRelationships | undefined): boolean =>
  !live || live.has(attr(entry, 'r:id') ?? '');

/** ECMA-376 Part 1, 19.2.1.33 `sldId`: `ST_SlideId` is 256 to 2147483647. */
const MIN_SLIDE_ID = 256;
const MAX_SLIDE_ID = 2_147_483_647;

/**
 * A slide id as a number names it: `ST_SlideId` is an `xsd:unsignedInt`, so "0258" and "258" are
 * one id. Digits lose their leading zeros; any other value stays as it is. @internal
 */
export const canonicalSlideId = (id: string): string =>
  /^\d+$/.test(id) ? id.replace(/^0+(?=\d)/, '') : id;

/** Whether `id` is an `ST_SlideId`: an integer from 256 to 2147483647. */
const isSlideId = (id: string): boolean =>
  /^\d+$/.test(id) &&
  canonicalSlideId(id).length <= 10 &&
  Number(id) >= MIN_SLIDE_ID &&
  Number(id) <= MAX_SLIDE_ID;

/** Every slide id that the part names, canonical: in `p:sldIdLst` and in each section list. */
function takenSlideIds(root: Element): Set<string> {
  const taken = new Set(slideEntries(root).map((s) => canonicalSlideId(attr(s, 'id') ?? '')));
  for (const list of root.children.filter(isExtLst))
    for (const ext of elements(list))
      for (const section of elements(sectionListOf(ext)))
        for (const s of slideIdsOf(section)) taken.add(canonicalSlideId(attr(s, 'id') ?? ''));
  return taken;
}

/** A slide id that `taken` does not hold: one past the largest, or the first gap from 256. */
function freshSlideId(taken: Set<string>): string {
  let max = MIN_SLIDE_ID - 1;
  for (const id of taken) {
    // Only an `ST_SlideId` counts: "1e3" or "-1" sets no maximum.
    if (!isSlideId(id)) continue;
    const n = Number(id);
    if (n > max) max = n;
  }
  let next = max + 1;
  if (next > MAX_SLIDE_ID) for (next = MIN_SLIDE_ID; taken.has(String(next));) next += 1;
  const id = String(next);
  taken.add(id);
  return id;
}

/** The entries that stay, to the entry as it is saved, and the slide ids that were renewed. */
interface KeptEntries {
  readonly out: Map<Element, Element>;
  /**
   * Canonical old id to fresh id, for an entry that got a fresh id because its own id is not an
   * `ST_SlideId` and no other entry has it. A section that names the old id names the slide.
   */
  readonly renewed: Map<string, string>;
}

/**
 * The `p:sldId` entries that stay, in slide order: live ones, and of entries that name one
 * slide part (by `r:id`, or by the part that `live` maps it to) only the first. An entry that
 * repeats a slide id (compared as a number, `canonicalSlideId`) but names another part is
 * another slide: two replicas that each add a slide can give both one id. It
 * stays, with a fresh slide id, so the slide is not lost; no section names the new id, so the
 * slide joins the section of the slide before it. An entry whose id is missing, not an integer,
 * or outside `ST_SlideId` also gets a fresh id; when no other entry has its id, a section that
 * names the old id keeps the slide (`renewed`).
 */
function keptEntries(root: Element, live: LiveRelationships | undefined): KeptEntries {
  const parts = new Set<string>();
  const candidates: Element[] = [];
  for (const s of slideEntries(root)) {
    if (!isLive(s, live)) continue;
    const rId = attr(s, 'r:id') ?? '';
    const part = (live instanceof Map ? live.get(rId) : undefined) ?? rId;
    if (part && parts.has(part)) continue;
    if (part) parts.add(part);
    candidates.push(s);
  }
  const count = new Map<string, number>();
  for (const s of candidates) {
    const key = canonicalSlideId(attr(s, 'id') ?? '');
    count.set(key, (count.get(key) ?? 0) + 1);
  }
  const ids = new Set<string>();
  let taken: Set<string> | undefined;
  const out = new Map<Element, Element>();
  const renewed = new Map<string, string>();
  for (const s of candidates) {
    const id = attr(s, 'id') ?? '';
    const key = canonicalSlideId(id);
    if (isSlideId(id) && !ids.has(key)) {
      ids.add(key);
      out.set(s, s);
      continue;
    }
    taken ??= takenSlideIds(root);
    const fresh = freshSlideId(taken);
    ids.add(fresh);
    // Only an invalid id that one entry has: a repeated id belongs to the slide that kept it.
    if (id && !isSlideId(id) && count.get(key) === 1) renewed.set(key, fresh);
    const has = s.attrs.some((a) => a.ns === '' && a.local === 'id');
    out.set(s, {
      ...s,
      attrs: has
        ? s.attrs.map((a) => (a.ns === '' && a.local === 'id' ? { ...a, value: fresh } : a))
        : [{ ns: '', local: 'id', value: fresh }, ...s.attrs],
    });
  }
  return { out, renewed };
}

/**
 * The slide that a section's `p14:sldId` value names, as a canonical id of the saved deck:
 * renewed when its entry got a fresh id.
 */
const sectionKey = (kept: KeptEntries, id: string): string => {
  const key = canonicalSlideId(id);
  return kept.renewed.get(key) ?? key;
};

/** Whether the root needs the repair. */
function needsRepair(root: Element, live: LiveRelationships | undefined): boolean {
  const kept = keptEntries(root, live);
  if (root.children.filter(isSlideList).length > 1) return true;
  if (kept.out.size !== slideEntries(root).length || [...kept.out].some(([a, b]) => a !== b))
    return true;
  const entries = [...kept.out.values()];
  const lists = root.children.filter(isExtLst);
  if (lists.length > 1) return true;
  let sectionExts = 0;
  const sectionIds = new Set<string>();
  const named: string[] = [];
  const seen = new Set<string>();
  for (const ext of elements(lists[0])) {
    if (attr(ext, 'uri') !== SECTION_LIST_URI) continue;
    sectionExts += 1;
    const list = sectionListOf(ext);
    if (!list || sectionExts > 1) return true;
    for (const section of elements(list)) {
      const id = attr(section, 'id') ?? '';
      if (sectionIds.has(id)) return true;
      // `p14:sldIdLst` is required (the `p14` schema, `CT_Section`).
      if (!elements(section).some((e) => e.local === 'sldIdLst')) return true;
      sectionIds.add(id);
      for (const s of slideIdsOf(section)) {
        const value = sectionKey(kept, attr(s, 'id') ?? '');
        if (seen.has(value)) return true;
        seen.add(value);
        named.push(value);
      }
    }
  }
  if (!sectionIds.size) return false;
  // The sections name every slide of the deck once, in slide order, and no other slide.
  const slides = entries.map((s) => canonicalSlideId(attr(s, 'id') ?? ''));
  return slides.length !== named.length || slides.some((id, i) => named[i] !== id);
}

/** A `p14:sldId` element for slide id `id`, shaped like `like`. */
function slideIdElement(like: Element | undefined, id: string, key: string): Element {
  return {
    type: 'element',
    id: `repair-${key}`,
    kind: like?.kind ?? 'generic',
    ns: NS.p14,
    local: 'sldId',
    prefix: like?.prefix ?? 'p14',
    nsDecls: [],
    attrs: [{ ns: '', local: 'id', value: id }],
    children: [],
  } as Element;
}

/** A `p14:sldIdLst` for a section that has none. */
function slideListElement(section: Element, entries: Element[], k: number): Element {
  return {
    type: 'element',
    id: `repair-list-${k}`,
    kind: 'generic',
    ns: NS.p14,
    local: 'sldIdLst',
    prefix: section.prefix ?? 'p14',
    nsDecls: [],
    attrs: [],
    children: entries,
  } as Element;
}

/**
 * The presentation root with no `p:sldId` outside `live` (the `r:id` values whose relationship
 * names a part; every id when omitted), one entry for each slide part, a fresh slide id for an
 * entry that repeats the slide id of another part (`keptEntries`), one `p:sldIdLst` at the place
 * of the first that holds the entries of all lists in document order, one `p:extLst` at the place
 * of the first, one section-list extension (the first that holds a list), and every section
 * once by id, in the
 * order of its first slide. Each slide of the deck is in one section, in slide order: the first
 * section that names it, unless a slide before it is in a later section; a slide that no
 * section names goes in the section of the slide before it. Other extensions stay as they are.
 * The root itself when it needs none of this.
 */
export function repairPresentationRoot(root: Element, live?: LiveRelationships): Element {
  if (!needsRepair(root, live)) return root;
  const chosen = keptEntries(root, live);
  const saved = chosen.out;
  const entries = [...saved.values()];
  // The saved slide ids, and the canonical form that sections are matched by.
  const deck = entries.map((s) => attr(s, 'id') ?? '');
  const keys = deck.map(canonicalSlideId);
  const lists = root.children.filter(isExtLst);
  const all = lists.flatMap((l) => elements(l));
  const sectionExt = all.find((e) => attr(e, 'uri') === SECTION_LIST_URI && sectionListOf(e));
  const sections: Element[] = [];
  const exts: Element[] = [];
  let placed = false;
  for (const ext of all) {
    if (attr(ext, 'uri') !== SECTION_LIST_URI) {
      exts.push(ext);
      continue;
    }
    for (const section of elements(sectionListOf(ext))) sections.push(section);
    // The section list goes where the first section-list extension was; an extension without
    // a list goes.
    if (!placed && sectionExt) exts.push(sectionExt);
    placed = true;
  }
  // Each section once; the first section that names a slide of the deck owns it.
  const at = new Map<string, number>();
  keys.forEach((key, i) => at.has(key) || at.set(key, i));
  const sectionIds = new Set<string>();
  const kept: { el: Element; like: Element | undefined; first: number }[] = [];
  const owner = new Map<string, (typeof kept)[number]>();
  for (const section of sections) {
    const id = attr(section, 'id') ?? '';
    if (sectionIds.has(id)) continue;
    sectionIds.add(id);
    const named = slideIdsOf(section);
    const entry = { el: section, like: named[0], first: Infinity };
    for (const s of named) {
      const value = sectionKey(chosen, attr(s, 'id') ?? '');
      const i = at.get(value);
      if (i === undefined || owner.has(value)) continue;
      owner.set(value, entry);
      entry.first = Math.min(entry.first, i);
    }
    kept.push(entry);
  }
  // The sections follow slide order: each goes where its first slide is, and a section with no
  // slide stays after the section listed before it.
  let last = -1;
  const order = kept
    .map((entry) => {
      if (entry.first !== Infinity) last = entry.first;
      return { entry, key: entry.first === Infinity ? last : entry.first };
    })
    .sort((a, b) => a.key - b.key)
    .map((x) => x.entry);
  const index = new Map(order.map((entry, k) => [entry, k]));
  // In slide order, the section only moves forward: a slide goes to its owner, or to the
  // section of the slide before it when that one is later or the slide has no owner. So each
  // section holds one run of slides.
  const ids: string[][] = order.map(() => []);
  const listed = new Set<string>();
  let cursor = 0;
  for (const [i, id] of deck.entries()) {
    const key = keys[i]!;
    if (listed.has(key)) continue;
    listed.add(key);
    const own = owner.get(key);
    cursor = Math.max(cursor, own ? index.get(own)! : cursor);
    ids[cursor]?.push(id);
  }
  const rebuilt = order.map(({ el, like }, k): Element => {
    const entries = ids[k]!.map((id, i) => slideIdElement(like, id, `${k}-${i}`));
    const has = el.children.some((c) => c.type === 'element' && c.local === 'sldIdLst');
    // `p14:sldIdLst` is the first child of a section (the `p14` schema, `CT_Section`).
    return {
      ...el,
      children: has
        ? el.children.map((c) =>
            c.type === 'element' && c.local === 'sldIdLst' ? { ...c, children: entries } : c
          )
        : [slideListElement(el, entries, k), ...el.children],
    };
  });
  const merged = exts.map((ext): Element => {
    const list = sectionListOf(ext);
    if (attr(ext, 'uri') !== SECTION_LIST_URI || !list) return ext;
    return {
      ...ext,
      children: ext.children.map((c) => (c === list ? { ...list, children: rebuilt } : c)),
    };
  });
  const first = lists[0];
  const slideLists = root.children.filter(isSlideList);
  // The first slide list keeps its other nodes and takes the entries of every list.
  const slideList = slideLists[0];
  const savedEntries = slideLists.flatMap((list) =>
    list.children.flatMap((s): Node[] =>
      s.type !== 'element' ? (list === slideList ? [s] : []) : saved.has(s) ? [saved.get(s)!] : []
    )
  );
  return {
    ...root,
    children: root.children
      .filter((c) => (!isExtLst(c) || c === first) && (!isSlideList(c) || c === slideList))
      .map((c) =>
        c === first
          ? { ...first, children: merged }
          : c === slideList
            ? { ...c, children: savedEntries }
            : c
      ),
  };
}
