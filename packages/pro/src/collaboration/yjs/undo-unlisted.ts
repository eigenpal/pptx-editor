/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// What a peer wrote into a node that an undo takes out of the document, after DOCX Editor's
// `packages/pro/src/collaboration/document/undo-unlisted-text.ts` (commit 5993bd22). DOCX Editor
// keeps a paragraph's text in one shared text, so it records where an unlisted paragraph's text
// joins; here every node has its own record, so the rule records where each node was listed.
//
// An undo of a step that created a node removes the key that lists the node under its parent,
// and keeps the record (`undo-delete-filter.ts`). A peer's later write into the node (a line
// width in a new `a:ln`, a character typed into a new run, a bullet in a new `a:pPr`) stays in
// that record. So before the undo runs, the undoing replica records, for each key the step listed,
// its parent, the key before it, the undoing client, and its clock (`UNDONE_KEY`, outside the
// undo scope). A read shows such a node again at that place while it holds anything another
// client wrote, or the undoing client wrote after the undo, with only that content: the undo
// took the undoing author's own attributes, children, and characters. A node that someone
// removed on purpose (`rw`) stays removed. Every replica reads the same room, so all show the
// same nodes, also when the peer's write arrives after the undo.
//
// Everything read here comes from the room and is untrusted: an entry of a wrong shape is
// ignored, and each read looks at no more than `MAX_UNDONE` entries and `MAX_DEPTH` levels.

import * as Y from 'yjs';
import type { ReplicaElementSpec, ReplicaTextSpec } from '@pptx-editor.dev/core/store';
import { F, type Room } from './schema.ts';
import { IDENTITY } from './text-identity.ts';
import { isRemoved } from './text-lineage.ts';
import { forEachItem, holderItem, mapItems, sequenceItems, type StepRanges } from './yjs-items.ts';

/** The room map of undone listings: key → `{ p, u, a, c, q }` (see `UndoneEntry`). */
export const UNDONE_KEY = 'pptx-undone-v1';

/** The deepest chain of undone nodes a read follows. */
const MAX_DEPTH = 64;
/** The most nodes one parent shows again in a read; only a hostile room unlists more. */
const MAX_PLACED = 1_024;
/** The records one search for an emptied node's lists looks at. */
const MAX_SCAN = 1 << 18;

/** Where an undo took a node from. */
interface UndoneEntry {
  /** The node's key. The stored value names its part (`p`) and the fields below. */
  readonly key: string;
  /** Its parent's key. */
  readonly parent: string;
  /** The key listed right before it, or '' when it was first. */
  readonly after: string;
  /** The key listed right after it, or '' when it was last. */
  readonly before: string;
  /** The client that undid the step. */
  readonly client: number;
  /** That client's clock before the undo: what it wrote in the node from then on counts too. */
  readonly clock: number;
  /**
   * For a listing a write deleted because the read showed the node elsewhere (`s`): the key of
   * the parent it showed in. The node shows here again only while that parent is listed.
   */
  readonly shownIn: string | null;
}

/**
 * Whether an item is content the undo did not take: another client wrote it, or the undoing
 * client wrote it after the undo (the undo itself writes nothing into a node it unlists).
 */
const isForeign = (item: Y.Item, e: UndoneEntry): boolean =>
  item.id.client !== e.client || item.id.clock >= e.clock;

function entryOf(key: string, value: unknown, partName?: string): UndoneEntry | null {
  if (typeof value !== 'object' || value === null) return null;
  const v = value as Record<string, unknown>;
  if (typeof v.p !== 'string' || (partName !== undefined && v.p !== partName)) return null;
  if (typeof v.u !== 'string' || typeof v.a !== 'string') return null;
  if (!Number.isSafeInteger(v.c) || (v.c as number) < 0) return null;
  // No clock (or a wrong one) counts nothing of the undoing client's as written after the undo.
  const clock = Number.isSafeInteger(v.q) && (v.q as number) >= 0 ? (v.q as number) : Infinity;
  const before = typeof v.n === 'string' ? v.n : '';
  const shownIn = typeof v.s === 'string' ? v.s : null;
  return { key, parent: v.u, after: v.a, before, client: v.c as number, clock, shownIn };
}

/**
 * Before an undo of `step` runs: record, for each key the step listed in a node's children,
 * where it was listed. `transact` runs the write under an origin no history tracks.
 */
export function noteUndoneListings(
  room: Room,
  step: StepRanges,
  transact: (write: () => void) => void
): void {
  const nodes: unknown = room.nodes;
  const listed = new Map<Y.Array<unknown>, Set<string>>();
  // A part the step created goes with its undo: none of its nodes can show again.
  const created = new Set<string>();
  forEachItem(room.doc, step.insertions, (item) => {
    if (item.parent === (room.parts as unknown) && item.parentSub !== null && !item.deleted)
      created.add(item.parentSub);
  });
  forEachItem(room.doc, step.insertions, (item) => {
    const list = item.parent;
    if (item.deleted || !(list instanceof Y.Array)) return;
    const record = list.parent;
    if (!(record instanceof Y.Map) || holderItem(record)?.parent !== nodes) return;
    if (record.get(F.children) !== list) return;
    if (item.parentSub !== null || !(item.content instanceof Y.ContentAny)) return;
    let keys = listed.get(list);
    if (!keys) listed.set(list, (keys = new Set()));
    for (const k of item.content.getContent()) if (typeof k === 'string') keys.add(k);
  });
  if (listed.size === 0) return;
  const undone = room.doc.getMap(UNDONE_KEY);
  const client = room.doc.clientID;
  const clock = Y.getState(room.doc.store, client);
  type Value = { p: string; u: string; a: string; n: string; c: number; q: number };
  const writes: [string, Value][] = [];
  for (const [list, keys] of listed) {
    const record = list.parent;
    if (!(record instanceof Y.Map) || record.get(F.children) !== list) continue;
    const parent = record.get(F.key);
    const part = record.get(F.part);
    if (typeof parent !== 'string' || typeof part !== 'string' || created.has(part)) continue;
    const siblings = list.toArray().filter((k): k is string => typeof k === 'string');
    siblings.forEach((k, i) => {
      if (!keys.has(k)) return;
      const n = siblings[i + 1] ?? '';
      const value = { p: part, u: parent, a: siblings[i - 1] ?? '', n, c: client, q: clock };
      if (JSON.stringify(undone.get(k)) !== JSON.stringify(value)) writes.push([k, value]);
    });
  }
  if (writes.length > 0)
    transact(() => {
      for (const [k, v] of writes) undone.set(k, v);
    });
}

/**
 * Before a write deletes listings that no read shows (`dropHiddenListings` in `mirror.ts`):
 * record each such listing of `list` (the children of `parent` in `part`) at `indexes`, as an
 * undo does, with the parent that the read shows the node in (`shownIn`). The shown listing can
 * be a peer's undo of a removal, and that peer's redo deletes it again; the node then shows here
 * once more, as the room listed it, with all its content (clock 0: everything counts). It does
 * not when the parent it showed in went too (a peer removed that place on purpose), or while a
 * list holds the node. Call inside the write's transaction.
 */
export function noteTidiedListings(
  room: Room,
  part: string,
  parent: string,
  list: readonly unknown[],
  indexes: readonly number[],
  shownIn: (key: string) => string | undefined
): void {
  const undone = room.doc.getMap(UNDONE_KEY);
  const client = room.doc.clientID;
  const noted = new Set<string>();
  for (const i of indexes) {
    const key = list[i];
    const shown = typeof key === 'string' ? shownIn(key) : undefined;
    if (typeof key !== 'string' || shown === undefined || noted.has(key)) continue;
    noted.add(key);
    // The nearest other keys around it: a repeat of the same key is no anchor.
    let a = '';
    for (let j = i - 1; j >= 0 && a === ''; j -= 1)
      if (list[j] !== key && typeof list[j] === 'string') a = list[j] as string;
    let n = '';
    for (let j = i + 1; j < list.length && n === ''; j += 1)
      if (list[j] !== key && typeof list[j] === 'string') n = list[j] as string;
    const value = { p: part, u: parent, a, n, c: client, q: 0, s: shown };
    if (JSON.stringify(undone.get(key)) !== JSON.stringify(value)) undone.set(key, value);
  }
}

/**
 * After a redo of `step`: forget the undone listings of the keys it listed again. They show
 * where they are listed now, and a later undo records them again.
 */
export function forgetUndoneListings(
  room: Room,
  step: StepRanges,
  transact: (write: () => void) => void
): void {
  const undone = room.doc.getMap(UNDONE_KEY);
  if (undone.size === 0) return;
  const keys = new Set<string>();
  forEachItem(room.doc, step.deletions, (item) => {
    const list = item.parent;
    if (!(list instanceof Y.Array) || !(item.content instanceof Y.ContentAny)) return;
    for (const k of item.content.getContent())
      if (typeof k === 'string' && undone.has(k)) keys.add(k);
  });
  if (keys.size > 0)
    transact(() => {
      for (const k of keys) undone.delete(k);
    });
}

/**
 * After an undo of `step`: a node the step created that a peer listed again meanwhile (a
 * repair that moved it) stays listed, by the peer's entry, though the undo took all of its
 * content. Such a node that holds nothing now leaves every list of its part, as the undo meant.
 * Only an element record with no attribute and no child that the read still shows (`shows`)
 * counts, so the search runs only then. Call inside a transaction; returns whether it wrote
 * anything.
 */
export function unlistEmptied(
  room: Room,
  step: StepRanges,
  shows: (part: string, key: string) => boolean
): boolean {
  const emptied = new Map<string, Set<string>>();
  forEachItem(room.doc, step.insertions, (item) => {
    if ((item.parent as unknown) !== room.nodes || item.parentSub === null) return;
    if (!(item.content instanceof Y.ContentType) || !(item.content.type instanceof Y.Map)) return;
    const rec = item.content.type as Y.Map<unknown>;
    if (room.nodes.get(item.parentSub) !== rec) return;
    const part = rec.get(F.part);
    const attrs = rec.get(F.attrs);
    const children = rec.get(F.children);
    if (typeof part !== 'string' || !(attrs instanceof Y.Map) || !(children instanceof Y.Array))
      return;
    if (attrs.size > 0 || children.length > 0 || !shows(part, item.parentSub)) return;
    let keys = emptied.get(part);
    if (!keys) emptied.set(part, (keys = new Set()));
    keys.add(item.parentSub);
  });
  if (emptied.size === 0) return false;
  let wrote = false;
  let looked = 0;
  for (const rec of room.nodes.values()) {
    if ((looked += 1) > MAX_SCAN) break;
    if (!(rec instanceof Y.Map)) continue;
    const part = rec.get(F.part);
    const keys = typeof part === 'string' ? emptied.get(part) : undefined;
    const list = rec.get(F.children);
    if (!keys || !(list instanceof Y.Array)) continue;
    const entries: unknown[] = list.toArray();
    for (let i = entries.length - 1; i >= 0; i -= 1) {
      const entry = entries[i];
      if (typeof entry === 'string' && keys.has(entry)) {
        list.delete(i, 1);
        wrote = true;
      }
    }
  }
  return wrote;
}

/**
 * The undone listings by part, then parent, then key. Built in one walk the first time, then
 * kept current from the keys each change of the map names, so a long session costs each undo
 * only its own entries. No count cap: an honest session keeps every listing it needs.
 */
interface UndoneIndex {
  readonly byPart: Map<string, Map<string, Map<string, UndoneEntry>>>;
  readonly placeOf: Map<string, { part: string; parent: string }>;
}
const indexes = new WeakMap<Y.Map<unknown>, UndoneIndex>();

function indexEntry(index: UndoneIndex, key: string, value: unknown): void {
  const was = index.placeOf.get(key);
  if (was) {
    const byParent = index.byPart.get(was.part);
    const list = byParent?.get(was.parent);
    list?.delete(key);
    if (list?.size === 0) byParent!.delete(was.parent);
    index.placeOf.delete(key);
  }
  const e = value === undefined ? null : entryOf(key, value);
  const part = e ? (value as { p: string }).p : null;
  if (!e || part === null) return;
  let byParent = index.byPart.get(part);
  if (!byParent) index.byPart.set(part, (byParent = new Map()));
  let list = byParent.get(e.parent);
  if (!list) byParent.set(e.parent, (list = new Map()));
  list.set(key, e);
  index.placeOf.set(key, { part, parent: e.parent });
}

function undoneOfPart(room: Room, partName: string): Map<string, Map<string, UndoneEntry>> {
  const map = room.doc.getMap(UNDONE_KEY);
  let index = indexes.get(map);
  if (!index) {
    const built: UndoneIndex = { byPart: new Map(), placeOf: new Map() };
    index = built;
    for (const [key, value] of map.entries()) indexEntry(built, key, value);
    indexes.set(map, built);
    map.observe((event) => {
      for (const key of event.keysChanged) indexEntry(built, key, map.get(key));
    });
  }
  return index.byPart.get(partName) ?? new Map();
}

type Spec = ReplicaElementSpec | ReplicaTextSpec;

const NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
/**
 * Elements whose attributes and children are all optional in the schema, so a part of their
 * content is a valid element: the paragraph (CT_TextParagraph) and the property elements a
 * format or a command creates (CT_TextParagraphProperties, CT_TextCharacterProperties,
 * CT_LineProperties, CT_TextBodyProperties, CT_TableCellProperties). ECMA-376 Part 1, 21.1.2
 * and 20.1.2.2.24 ln.
 */
const PARTIAL_OK: ReadonlySet<string> = new Set([
  'p',
  'pPr',
  'rPr',
  'endParaRPr',
  'ln',
  'bodyPr',
  'tcPr',
]);

/**
 * Whether a node an undo unlisted stands alone with only the content the read shows of it. A
 * node outside this rule (a shape, a table row, a text body) needs content the undo took, so it
 * does not show again: a part of it would not be valid (ECMA-376 Part 1, 21.1.3.18 tr requires
 * `h`; 21.1.2.1.1 txBody requires `a:bodyPr`).
 */
function standsAlone(node: Spec): boolean {
  if ('text' in node) return node.text.length > 0;
  if (node.ns !== NS_A) return false;
  const kids = node.children ?? [];
  // A run needs its text (CT_RegularTextRun: `a:t` is required), and `a:t` holds the text.
  if (node.local === 'r')
    return kids.some((c) => !('text' in c) && c.ns === NS_A && c.local === 't' && standsAlone(c));
  if (node.local === 't') return kids.some((c) => 'text' in c && c.text.length > 0);
  return PARTIAL_OK.has(node.local) && (kids.length > 0 || (node.attrs ?? []).length > 0);
}

/**
 * One part read's view of the undone listings: which nodes show again, under which parent, and
 * after which key. Built per part read; null when the part has no undone listing.
 */
export class UndoneListings {
  readonly #room: Room;
  readonly #byParent: Map<string, Map<string, UndoneEntry>>;
  /** The keys the room lists under the part's root: those show where they are listed. */
  readonly #reached: ReadonlySet<string>;
  readonly #foreign = new Map<string, boolean>();
  /** The keys this read placed: no list holds them, and a write does not list them. */
  readonly shown = new Set<string>();

  private constructor(
    room: Room,
    byParent: Map<string, Map<string, UndoneEntry>>,
    reached: Set<string>
  ) {
    this.#room = room;
    this.#byParent = byParent;
    this.#reached = reached;
  }

  /** The undone listings of a part, or null when it has none. */
  static of(room: Room, partName: string, rootKey: string): UndoneListings | null {
    if (room.doc.getMap(UNDONE_KEY).size === 0) return null;
    const byParent = undoneOfPart(room, partName);
    if (byParent.size === 0) return null;
    // The keys the room lists under the part's root.
    const reached = new Set<string>();
    const stack = [rootKey];
    while (stack.length > 0) {
      const key = stack.pop()!;
      if (reached.has(key)) continue;
      reached.add(key);
      const rec = room.nodes.get(key);
      if (!(rec instanceof Y.Map) || rec.get(F.part) !== partName) continue;
      const list: unknown = rec.get(F.children);
      if (!(list instanceof Y.Array)) continue;
      for (const c of list.toArray()) if (typeof c === 'string' && !reached.has(c)) stack.push(c);
    }
    return new UndoneListings(room, byParent, reached);
  }

  /** Whether a node holds content that the undo of `e` did not take, here or below. */
  #holdsForeign(key: string, e: UndoneEntry, depth: number): boolean {
    const memoKey = `${key}\u0000${e.client}\u0000${e.clock}`;
    const known = this.#foreign.get(memoKey);
    if (known !== undefined) return known;
    // A cycle or a chain deeper than any writer makes reads as holding nothing: set first.
    this.#foreign.set(memoKey, false);
    const result = depth <= MAX_DEPTH && this.#holds(key, e, depth);
    this.#foreign.set(memoKey, result);
    return result;
  }

  #holds(key: string, e: UndoneEntry, depth: number): boolean {
    const rec = this.#room.nodes.get(key);
    if (!(rec instanceof Y.Map)) return false;
    const foreign = (item: Y.Item) => isForeign(item, e);
    const text = rec.get(F.text);
    if (text instanceof Y.Text)
      for (const item of sequenceItems(text)) if (item.countable && foreign(item)) return true;
    const attrs = rec.get(F.attrs);
    if (attrs instanceof Y.Map) for (const item of mapItems(attrs)) if (foreign(item)) return true;
    const children = rec.get(F.children);
    if (children instanceof Y.Array)
      for (const item of sequenceItems(children)) if (foreign(item)) return true;
    // A run a write removed from the node while a peer typed into it shows again with that
    // text (`withRemovedRuns` in `text-identity.ts`), so it counts as content of the node.
    let runs = 0;
    for (const field of rec.keys()) {
      if (!field.startsWith(IDENTITY.removedRun)) continue;
      if ((runs += 1) > MAX_PLACED) break;
      if (this.#holdsForeign(field.slice(IDENTITY.removedRun.length), e, depth + 1)) return true;
    }
    return this.#shown(key).some((child) => this.#holdsForeign(child.key, child, depth + 1));
  }

  /** The undone listings under `parent` that may show: no list holds them, none was removed. */
  #shown(parent: string): UndoneEntry[] {
    return [...(this.#byParent.get(parent)?.values() ?? [])].filter((e) => {
      if (this.#reached.has(e.key)) return false;
      if (e.shownIn !== null && !this.#reached.has(e.shownIn)) return false;
      const rec = this.#room.nodes.get(e.key);
      return rec instanceof Y.Map && !isRemoved(rec as Y.Map<unknown>);
    });
  }

  /**
   * `children` of `parent` with the undone nodes that hold another client's content, each after
   * the key it followed (or the nearest earlier one that shows), else first or last. `read`
   * reads an element record; a key already in `seen` is not placed again.
   */
  withUndone(
    parent: string,
    partName: string,
    children: readonly Spec[],
    seen: Set<string>,
    read: (key: string) => ReplicaElementSpec | null
  ): readonly Spec[] {
    // Key order, so every replica places nodes that follow one key in the same order, whatever
    // order its map holds the entries in.
    const entries = this.#shown(parent)
      .filter((e) => !seen.has(e.key) && this.#holdsForeign(e.key, e, 0))
      .sort((x, y) => (x.key < y.key ? -1 : x.key > y.key ? 1 : 0))
      .slice(0, MAX_PLACED);
    if (entries.length === 0) return children;
    const out = [...children];
    const present = new Set(out.map((c) => c.key));
    // Where each undone key of this parent followed, also of those that do not show, so a node
    // whose anchor stays hidden goes after the nearest earlier key that shows.
    const after = new Map(
      [...(this.#byParent.get(parent)?.values() ?? [])].map((e) => [e.key, e.after])
    );
    const pending = new Set(entries.map((e) => e.key));
    const END = '\u0000';
    // Each pass places the nodes whose anchor shows or is gone; a node whose anchor still waits
    // goes in a later pass, and a cycle of waiting anchors (a hostile room) goes last.
    for (let force = false; pending.size > 0;) {
      let progress = false;
      for (const e of entries) {
        if (!pending.has(e.key)) continue;
        let anchor = e.after;
        // Its previous sibling is gone (a peer deleted it): it goes before its next one.
        const next =
          anchor !== '' && !present.has(anchor) && !pending.has(anchor) && present.has(e.before)
            ? e.before
            : null;
        if (next !== null) anchor = END;
        const visited = new Set([e.key]);
        let waits = false;
        while (anchor !== '' && anchor !== END && !present.has(anchor)) {
          if (pending.has(anchor) && !force) {
            waits = true;
            break;
          }
          if (visited.has(anchor)) anchor = END;
          else {
            visited.add(anchor);
            anchor = after.get(anchor) ?? END;
          }
        }
        if (waits) continue;
        pending.delete(e.key);
        progress = true;
        const node = this.#read(e.key, partName, seen, read);
        if (!node || !standsAlone(node)) continue;
        const index =
          next !== null
            ? out.findIndex((c) => c.key === next)
            : anchor === ''
              ? 0
              : anchor === END
                ? out.length
                : out.findIndex((c) => c.key === anchor) + 1;
        out.splice(index, 0, node);
        present.add(e.key);
        this.shown.add(e.key);
      }
      if (!progress) force = true;
    }
    return out;
  }

  #read(
    key: string,
    partName: string,
    seen: Set<string>,
    read: (key: string) => ReplicaElementSpec | null
  ): Spec | null {
    const rec = this.#room.nodes.get(key);
    if (!(rec instanceof Y.Map) || rec.get(F.part) !== partName) return null;
    const text = rec.get(F.text);
    if (text instanceof Y.Text) {
      seen.add(key);
      return { key, text: text.toString() };
    }
    // Only a node that can stand alone is read, so a node that cannot show (a shape, a row)
    // neither marks its subtree as seen nor refuses the part.
    const shell: unknown = rec.get(F.shell);
    const { ns, local } = (typeof shell === 'object' && shell !== null ? shell : {}) as {
      ns?: unknown;
      local?: unknown;
    };
    if (ns !== NS_A || typeof local !== 'string') return null;
    if (!PARTIAL_OK.has(local) && local !== 'r' && local !== 't') return null;
    return read(key);
  }
}
