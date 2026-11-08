/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// What an undo must not take from a peer. A `Y.UndoManager` undo deletes this author's
// writes. Two of them can leave a part that a peer's later content breaks:
//
// - An attribute both peers wrote at once: Yjs kept one write and dropped the other, so the
//   undo of the winning write leaves no value, and a required attribute (`a:off/@x`) is gone.
//   `attributeWrites` records each attribute's value before a step; after the undo, a value
//   that is missing is written back.
// - A relationship and its part that this author added (an inserted picture), which a peer's
//   copy still points at. `lostReferences` finds each relationship id that a part still uses
//   but lost in the undo, and gives back the relationship, its target part, and the content
//   type entry the part needs, from the package before the undo.
// - A slide entry whose slide a peer deleted while this replica moved it: the delete wins, and
//   `withSlideListsRepaired` removes the entry and its section entry.
// - A relationship to a slide that a peer deleted (a copy of a slide jump): the delete wins, and
//   the hyperlinks that use it go, so no slide part comes back without its content type.

import * as Y from 'yjs';
import {
  CONTENT_TYPES_PART,
  asciiFold,
  NS,
  attr,
  children,
  materialize,
  relationshipsOf,
  relsPartName,
  specOf,
  type Element,
  type Node,
  type PackagePart,
  type PptxPackage,
  type XmlPart,
} from '@pptx-editor.dev/core/store';

/** The value of an attribute that a step added: there was none before it. */
export const ABSENT: unique symbol = Symbol('absent');

/**
 * One attribute a step wrote: its attribute map, its key, the value before the step (`ABSENT`
 * when it added the attribute), and the value the step left.
 */
export type AttributeWrite = [map: Y.Map<unknown>, key: string, before: unknown, after: unknown];

/**
 * The attributes a local transaction wrote, with their values before it and after it. Only maps
 * directly inside a node record count (`F.attrs`). The first change per key gives the value
 * before the step (an `add` gives `ABSENT`), and the last gives the value after it.
 */
export function attributeWrites(
  events: Iterable<Y.YEvent<Y.AbstractType<unknown>>>,
  isAttrs: (map: Y.Map<unknown>) => boolean,
  into: AttributeWrite[]
): void {
  const index = indexOf(into);
  for (const event of events) {
    const map = event.target;
    if (!(map instanceof Y.Map) || !isAttrs(map)) continue;
    let byKey = index.get(map as Y.Map<unknown>);
    if (!byKey) index.set(map as Y.Map<unknown>, (byKey = new Map()));
    for (const [key, change] of event.changes.keys) {
      const after = (map as Y.Map<unknown>).get(key);
      const known = byKey.get(key);
      if (known) known[3] = after;
      else {
        const write: AttributeWrite = [
          map as Y.Map<unknown>,
          key,
          change.action === 'add' ? ABSENT : change.oldValue,
          after,
        ];
        into.push(write);
        byKey.set(key, write);
      }
    }
  }
}

/** Each record's writes by map and key, so a large step records in linear time. */
const indexes = new WeakMap<AttributeWrite[], Map<Y.Map<unknown>, Map<string, AttributeWrite>>>();

function indexOf(into: AttributeWrite[]): Map<Y.Map<unknown>, Map<string, AttributeWrite>> {
  let index = indexes.get(into);
  if (index) return index;
  index = new Map();
  for (const write of into) {
    let byKey = index.get(write[0]);
    if (!byKey) index.set(write[0], (byKey = new Map()));
    if (!byKey.has(write[1])) byKey.set(write[1], write);
  }
  indexes.set(into, index);
  return index;
}

const live = (map: Y.Map<unknown>): boolean => map._item?.deleted !== true;

/**
 * After an undo: each attribute the step had overwrote or deleted that has no value now (a
 * peer wrote it at once, and Yjs dropped the earlier write), with the value to write back.
 */
export function missingAttributes(
  writes: readonly AttributeWrite[]
): [Y.Map<unknown>, string, unknown][] {
  return writes
    .filter(([map, key, before]) => before !== ABSENT && live(map) && !map.has(key))
    .map(([map, key, before]) => [map, key, before]);
}

/**
 * After a redo: each attribute the step wrote whose value is still the one from before the
 * step (a repair wrote it after the undo, so Yjs cannot redo over it), with the step's value.
 */
export function unredoneAttributes(
  writes: readonly AttributeWrite[]
): [Y.Map<unknown>, string, unknown][] {
  const same = (a: unknown, b: unknown) => JSON.stringify(a) === JSON.stringify(b);
  return writes
    .filter(
      ([map, key, before, after]) =>
        live(map) &&
        after !== undefined &&
        before !== ABSENT &&
        same(map.get(key), before) &&
        !same(before, after)
    )
    .map(([map, key, , after]) => [map, key, after]);
}

/**
 * The relationship ids a part's elements use (`r:embed`, `r:link`, `r:id`, and the rest). A slide
 * entry (`p:sldId`) does not count: a deleted slide does not come back for it (see
 * `withSlideListsRepaired`).
 */
function usedIds(root: Element): Set<string> {
  const out = new Set<string>();
  const stack: Node[] = [root];
  while (stack.length > 0) {
    const n = stack.pop()!;
    if (n.type !== 'element' || isSlideEntry(n)) continue;
    for (const a of n.attrs) if (a.ns === NS.r) out.add(a.value);
    for (const c of n.children) stack.push(c);
  }
  return out;
}

/** `part` with `added` appended to its root, with ids minted after its own. */
function appended(part: XmlPart, added: readonly Element[]): XmlPart {
  const counter = { next: part.nextId };
  const kids = added.map((el) => materialize(specOf(el), part.root.kind, counter));
  return {
    partName: part.partName,
    root: { ...part.root, children: [...part.root.children, ...kids] },
    nextId: counter.next,
  };
}

const isSlideEntry = (el: Element): boolean => el.ns === NS.p && el.local === 'sldId';

/**
 * The presentation part with its slide lists made consistent, or undefined when they are.
 *
 * - A slide entry (`p:sldId`) whose `r:id` names no relationship goes, with its section entries.
 * A slide that one peer deletes while another moves it comes back in `p:sldIdLst` and in its
 * section: the move inserts both entries again. The delete wins, as a deletion wins over a
 * concurrent move in DOCX Editor
 * (`packages/pro/src/collaboration/document/paragraph-text-deletions.ts`). DOCX Editor records the
 * deletion; here it is inferred from the missing relationship. A peer update that Yjs keeps pending
 * is not in the package yet, so it cannot make a live entry look dead (the test of a move that
 * arrives before its slide shows this).
 * - A section entry (`p14:sldId`) that names no slide of the list goes, and a slide that two
 *   sections name stays in one: walking the slides in list order, the first section at or after
 *   the section of the slide before. Each section then lists its slides in list order. An undo
 *   of a delete puts the slide back in its old section while a peer's move put it in another;
 *   a presentation app repairs a file that names a slide twice. DOCX Editor commit adf0e178.
 *
 * Every replica that sees the lists computes the same result, so two that repair at once remove
 * the same records. A file that arrives with such lists is made valid the same way.
 */
function withSlideListsRepaired(after: PptxPackage, live: Set<string>): XmlPart | undefined {
  for (const part of after.parts.values()) {
    if (part.kind !== 'xml') continue;
    const root = part.xml.root;
    if (root.ns !== NS.p || root.local !== 'presentation') continue;
    const list = root.children.find(
      (c): c is Element => c.type === 'element' && c.ns === NS.p && c.local === 'sldIdLst'
    );
    if (!list) return undefined;
    const order: string[] = [];
    // Dead entries by node: core can give a new slide the id of a dead entry.
    const dead = new Set<Element>();
    for (const e of list.children) {
      if (e.type !== 'element' || !isSlideEntry(e)) continue;
      const rid = e.attrs.find((a) => a.ns === NS.r && a.local === 'id')?.value;
      if (rid === undefined || !live.has(rid)) dead.add(e);
      else order.push(attr(e, 'id') ?? '');
    }
    // The sections, in file order, and the slide ids each one names.
    const sections: Element[] = [];
    const walk = (el: Element): void => {
      for (const c of el.children)
        if (c.type === 'element') {
          if (c.ns === NS.p14 && c.local === 'section') sections.push(c);
          else walk(c);
        }
    };
    walk(root);
    // The sections that name each slide id, in file order. Maps keep the work linear in the
    // number of entries, since this runs on every peer change and undo.
    const sectionIndex = new Map<Element, number>(sections.map((sec, i) => [sec, i]));
    const namedIn = new Map<string, number[]>();
    sections.forEach((sec, i) => {
      const lst = sec.children.find(
        (c): c is Element => c.type === 'element' && c.ns === NS.p14 && c.local === 'sldIdLst'
      );
      if (!lst) return;
      for (const e of children(lst, { ns: NS.p14, local: 'sldId' })) {
        const id = attr(e, 'id') ?? '';
        const at = namedIn.get(id);
        if (!at) namedIn.set(id, [i]);
        else if (at[at.length - 1] !== i) at.push(i);
      }
    });
    // Where each slide stays: the first section at or after the section of the slide before.
    const home = new Map<string, number>();
    let at = 0;
    for (const id of order) {
      const where = namedIn.get(id);
      if (!where) continue;
      const found = where.find((i) => i >= at) ?? where[0]!;
      home.set(id, found);
      at = Math.max(at, found);
    }
    const rank = new Map(order.map((id, i) => [id, i]));
    let changed = dead.size > 0;
    const fixSection = (sec: Element, index: number): Element => {
      const kids = sec.children.map((c) => {
        if (c.type !== 'element' || c.ns !== NS.p14 || c.local !== 'sldIdLst') return c;
        const seen = new Set<string>();
        const entries = c.children.filter((e) => {
          if (e.type !== 'element' || e.ns !== NS.p14 || e.local !== 'sldId') return true;
          const id = attr(e, 'id') ?? '';
          if (home.get(id) !== index || seen.has(id)) return false;
          seen.add(id);
          return true;
        });
        const sorted = [...entries].sort((x, y) => {
          const rx = x.type === 'element' ? (rank.get(attr(x, 'id') ?? '') ?? 0) : 0;
          const ry = y.type === 'element' ? (rank.get(attr(y, 'id') ?? '') ?? 0) : 0;
          return rx - ry;
        });
        if (sorted.length === c.children.length && sorted.every((k, i) => k === c.children[i]))
          return c;
        changed = true;
        return { ...c, children: sorted };
      });
      return kids.every((k, i) => k === sec.children[i]) ? sec : { ...sec, children: kids };
    };
    const rebuild = (el: Element): Element => {
      const kids = el.children
        .filter((n) => n.type !== 'element' || !dead.has(n))
        .map((c) => {
          if (c.type !== 'element') return c;
          const index = sectionIndex.get(c) ?? -1;
          return index >= 0 ? fixSection(c, index) : rebuild(c);
        });
      const same = kids.length === el.children.length && kids.every((k, i) => k === el.children[i]);
      return same ? el : { ...el, children: kids };
    };
    const next = rebuild(root);
    return changed ? { ...part.xml, root: next } : undefined;
  }
  return undefined;
}

const HYPERLINKS = new Set(['hlinkClick', 'hlinkHover', 'hlinkMouseOver']);

/**
 * Whether a relationship names a slide, by its type (transitional or strict). The type decides,
 * not the target part: a replica that never had the slide (it got the copy but not the delete)
 * must drop the link too.
 */
const isSlideRelationship = (type: string): boolean => type.endsWith('/relationships/slide');

/**
 * `root` without the hyperlinks that use relationship `id`, or undefined when another kind of
 * element uses it too. A slide jump to a deleted slide loses its action, as `deleteSlide` in
 * core does for a local delete.
 */
function withoutLinks(root: Element, id: string): Element | undefined {
  let other = false;
  const strip = (el: Element): Element => {
    let changed = false;
    const kids: Node[] = [];
    for (const c of el.children) {
      if (c.type !== 'element') {
        kids.push(c);
        continue;
      }
      const uses = c.attrs.some((a) => a.ns === NS.r && a.value === id);
      if (uses && c.ns === NS.a && HYPERLINKS.has(c.local)) {
        changed = true;
        continue;
      }
      if (uses) other = true;
      const next = strip(c);
      if (next !== c) changed = true;
      kids.push(next);
    }
    return changed ? { ...el, children: kids } : el;
  };
  const next = strip(root);
  return other ? undefined : next;
}

/** `rels` without the relationship `id`. */
function withoutRelationship(rels: XmlPart, id: string): XmlPart {
  return {
    ...rels,
    root: {
      ...rels.root,
      children: rels.root.children.filter(
        (c) => c.type !== 'element' || c.local !== 'Relationship' || attr(c, 'Id') !== id
      ),
    },
  };
}

const xmlOf = (pkg: PptxPackage, name: string): XmlPart | undefined => {
  const p = pkg.parts.get(name);
  return p?.kind === 'xml' ? p.xml : undefined;
};

/**
 * The parts that give back what an undo, or a peer's change, took from `after` while a part
 * still uses it, or an empty list. Only parts whose tree or relationships the change touched are
 * read. The presentation part also loses the slide entries whose slide is gone
 * (`withSlideListsRepaired`).
 */
export function lostReferences(before: PptxPackage, after: PptxPackage): PackagePart[] {
  const out = new Map<string, XmlPart>();
  // Part names compare without ASCII case (ECMA-376 Part 2): a target `Slide3.xml` names the
  // part `slide3.xml`.
  const folded = new Set<string>();
  for (const name of after.parts.keys()) folded.add(asciiFold(name));
  const exists = (name: string): boolean => after.parts.has(name) || folded.has(asciiFold(name));
  const restored: PackagePart[] = [];
  const restoredNames = new Set<string>();
  const restore = (part: PackagePart | undefined): void => {
    if (!part || restoredNames.has(part.partName)) return;
    restoredNames.add(part.partName);
    restored.push(part);
  };
  // When the undo removed a part, any part may still point at it (a peer's copy on another
  // slide shares the media part), so every part is read. Otherwise only the changed ones.
  let removedPart = false;
  for (const name of before.parts.keys()) if (!after.parts.has(name)) removedPart = true;
  for (const [name, part] of after.parts) {
    if (part.kind !== 'xml') continue;
    const rels = relsPartName(name);
    if (
      !removedPart &&
      before.parts.get(name) === part &&
      before.parts.get(rels) === after.parts.get(rels)
    )
      continue;
    // A relationship to a slide that is gone: the delete wins. The hyperlinks that
    // use it go with it, so no slide part comes back without its relationships and content
    // type. A relationship that no element uses goes too, as `deleteSlide` in core does. One
    // that another kind of element uses stays. Returns whether the relationship went.
    const dropLink = (id: string): boolean => {
      const root = withoutLinks((out.get(name) ?? part.xml).root, id);
      if (!root) return false;
      out.set(name, { ...(out.get(name) ?? part.xml), root });
      const r = out.get(rels) ?? xmlOf(after, rels);
      if (r) out.set(rels, withoutRelationship(r, id));
      return true;
    };
    for (const rel of relationshipsOf(after, name))
      if (!rel.external && rel.target && !exists(rel.target) && isSlideRelationship(rel.type))
        dropLink(rel.id);
    const used = usedIds((out.get(name) ?? part.xml).root);
    if (used.size === 0) continue;
    // A relationship that stayed, but whose target part the undo removed (a peer's picture
    // that shares the part name of this author's).
    for (const rel of relationshipsOf(after, name)) {
      if (!used.has(rel.id) || rel.external || !rel.target || exists(rel.target)) continue;
      if (!isSlideRelationship(rel.type)) restore(before.parts.get(rel.target));
    }
    const have = new Set(relationshipsOf(after, name).map((r) => r.id));
    const old = xmlOf(before, rels);
    if (!old) continue;
    const missing = children(old.root, { ns: NS.rel, local: 'Relationship' }).filter((el) => {
      const id = attr(el, 'Id');
      return id !== undefined && used.has(id) && !have.has(id);
    });
    // A lost relationship to a slide that is gone does not come back either, when its links
    // can go. When another kind of element uses it, it comes back, so no `r:id` dangles.
    for (const rel of relationshipsOf(before, name)) {
      const at = missing.findIndex((el) => attr(el, 'Id') === rel.id);
      if (at < 0 || rel.external || !rel.target || exists(rel.target)) continue;
      if (!isSlideRelationship(rel.type)) continue;
      const root = withoutLinks((out.get(name) ?? part.xml).root, rel.id);
      if (!root) continue;
      out.set(name, { ...(out.get(name) ?? part.xml), root });
      missing.splice(at, 1);
    }
    if (missing.length === 0) continue;
    const base = out.get(rels) ??
      xmlOf(after, rels) ?? {
        partName: rels,
        root: { ...old.root, children: [] },
        nextId: old.nextId,
      };
    out.set(rels, appended(base, missing));
    const ids = new Set(missing.map((el) => attr(el, 'Id')));
    for (const rel of relationshipsOf(before, name)) {
      if (!ids.has(rel.id) || rel.external || !rel.target || exists(rel.target)) continue;
      restore(before.parts.get(rel.target));
    }
  }
  {
    const pres = [...after.parts.values()].find(
      (p) => p.kind === 'xml' && p.xml.root.ns === NS.p && p.xml.root.local === 'presentation'
    );
    // The relationships after this repair: also the ones it gives back.
    const rels =
      pres && (out.get(relsPartName(pres.partName)) ?? xmlOf(after, relsPartName(pres.partName)));
    // Only when the presentation part or its relationships changed, and never without its
    // relationships: then every entry would look dead.
    const presRels = pres && relsPartName(pres.partName);
    const touched =
      pres &&
      (before.parts.get(pres.partName) !== pres ||
        before.parts.get(presRels!) !== after.parts.get(presRels!) ||
        out.has(presRels!));
    if (rels && touched && presRels && after.parts.has(presRels)) {
      const live = new Set(
        children(rels.root, { ns: NS.rel, local: 'Relationship' }).map((r) => attr(r, 'Id') ?? '')
      );
      const presentation = withSlideListsRepaired(after, live);
      if (presentation) out.set(presentation.partName, presentation);
    }
  }
  // Every part needs its content type: its `Override`, or the `Default` of its extension. An undo
  // can remove an entry that a peer's part still uses (the first png's `Default`), so every part
  // of the result is checked, and a missing entry comes back from `before`.
  const types = xmlOf(after, CONTENT_TYPES_PART);
  const oldTypes = xmlOf(before, CONTENT_TYPES_PART);
  if (types && oldTypes) {
    // Extensions and part names compare without ASCII case (ECMA-376 Part 2).
    const keys = (root: Element, local: string, key: string) => {
      const m = new Map<string, Element>();
      for (const el of children(root, { ns: NS.ct, local })) {
        const v = attr(el, key);
        if (v !== undefined && !m.has(asciiFold(v))) m.set(asciiFold(v), el);
      }
      return m;
    };
    const overrides = keys(types.root, 'Override', 'PartName');
    const defaults = keys(types.root, 'Default', 'Extension');
    const oldOverrides = keys(oldTypes.root, 'Override', 'PartName');
    const oldDefaults = keys(oldTypes.root, 'Default', 'Extension');
    const needed = new Set<Element>();
    const names = new Set<string>(after.parts.keys());
    for (const p of restored) names.add(p.partName);
    for (const name of names) {
      if (name === CONTENT_TYPES_PART) continue;
      const ext = asciiFold(name.slice(name.lastIndexOf('.') + 1));
      if (overrides.has(asciiFold(name)) || defaults.has(ext)) continue;
      const el = oldOverrides.get(asciiFold(name)) ?? oldDefaults.get(ext);
      if (el) needed.add(el);
    }
    if (needed.size > 0) out.set(CONTENT_TYPES_PART, appended(types, [...needed]));
  }
  return [
    ...[...out.values()].map((xml): PackagePart => ({ kind: 'xml', partName: xml.partName, xml })),
    ...restored,
  ];
}
