/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// The text a peer's split moved while this replica edited it (openspec change
// `collaboration-text-lineage`, "Character identity"). The rules follow DOCX Editor's
// paragraph text design ("Moved text keeps an identity", "Deleted text stays deleted"), but
// PPTX Editor keeps one `Y.Text` per run, so a rule writes the room instead of changing a read:
// the replica that holds the knowledge writes the repair, once, when a peer's copy arrives
// (`text-identity.ts` marks copies). Every other replica reads the repaired room.
//
// - Follow: text this replica typed right after a character that the peer's split moved goes
//   after the shown copy of that character that "Once" keeps. Only the typist writes it (a move
//   is not idempotent). Only text the peer had not seen: text it had seen and did not move was
//   deleted in its view, and that delete is on its way.
// - Delete: a copy of a character this replica deleted is deleted too, and so is text that a
//   peer's "Follow" moved out of that character (`repairFollows`): the peer moved it before the
//   delete reached it. A delete is idempotent, so every replica that deleted the character may
//   write it.
// - Once: when two copies of one character show and an Enter made one of them (or a copy it
//   copies), the copy whose split starts later keeps it, then the one later in the document; the
//   other copy loses the character. Only the characters the new copy holds count. Two copies
//   made by formats alone are the split lineage's (`dropLosingSplits` in `text-lineage.ts`).
//   Any replica may write it (a delete).
// - Order: a paragraph that this replica's Enter made goes after the paragraph of a peer's Enter
//   that cut the same text earlier. Only the paragraph's author moves it.
// - Follow an Enter: a peer's copy in the paragraph that this replica's Enter moved its own copy
//   of the same characters out of goes along. Only the author of the Enter moves it.
//
// Every value read from the room is untrusted: live copies, list reads, and the steps of the
// paragraph rules are capped, ids are checked against the replica's own store before use, and a
// repair never writes a key it does not find in a list.

import * as Y from 'yjs';
import type { Element, Node, NodeId, PptxPackage } from '@pptx-editor.dev/core/store';
import { F, type Room } from './schema.ts';
import { LINEAGE } from './text-lineage.ts';
import {
  COPIES_KEY,
  MAX_CHAIN,
  copyOf,
  followsOf,
  idKey,
  isGone,
  itemOf,
  markFollow,
  markRestored,
  recordOfItem,
  restoredOf,
  rootOf,
  seenBy,
  type Copy,
} from './text-identity.ts';
import { deleteChars, liveIndex } from './yjs-items.ts';

/** The live copies one repair looks at; over the cap (only a hostile room), it repairs nothing. */
const MAX_COPIES = 4_096;
/** The copied characters one repair looks at, over all live copies. */
const MAX_CHARS = 1 << 20;
/** The list entries one search for the copies' runs reads. */
const MAX_LOOKS = 65_536;
/** The steps the paragraph rules (order, follow an Enter) take in one repair. */
const MAX_STEPS = 1 << 20;

/** What a repair reads of this replica: its keys and the package the store shows. */
export interface RepairView {
  readonly keyOf: (partName: string, id: NodeId) => string | undefined;
  readonly package: PptxPackage;
}

/** Text this replica typed that a repair moved: the first old id, the first new id, a length. */
export type MovedText = (from: Y.ID, to: Y.ID, length: number) => void;

/**
 * A peer's run that a repair moved after this replica's Enter (`followEnter`): the first
 * character of the Enter's copy, the list entry the move deleted, and the one it inserted. The
 * move belongs to the Enter's step, so an undo of the Enter takes the run back.
 */
export type MovedRun = (anchor: Y.ID, deleted: Y.ID, inserted: Y.ID, run: string) => void;

interface CopyEntry {
  readonly key: string;
  readonly copy: Copy;
  readonly text: Y.Text;
}

/** A shown, live copy character, and the copy that holds it. */
interface Holder {
  readonly copy: CopyEntry;
  readonly id: Y.ID;
}

/** Where the store shows a text record: its paragraph (with the paragraph's parent), its run. */
interface Place {
  readonly order: number;
  readonly p?: { readonly key: string; readonly parent: string };
  readonly run?: string;
}

/** What the paragraph rules share within one repair. */
interface Repair {
  readonly room: Room;
  readonly me: number;
  readonly incoming: readonly CopyEntry[];
  readonly all: readonly CopyEntry[];
  readonly placed: ReadonlyMap<string, Place>;
  /** Whether two copies hold a copy of one root character; it counts its steps. */
  readonly share: (a: CopyEntry, b: CopyEntry) => boolean;
  readonly laterStart: (a: CopyEntry, b: CopyEntry) => number;
  /** False once the rules took `MAX_STEPS` steps. */
  readonly step: (n?: number) => boolean;
  readonly movedRun?: MovedRun | undefined;
}

const NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const isElement = (n: Node): n is Element => n.type === 'element';

/** Where a root character stands in the text it was typed in, to order two splits' starts. */
function placeOf(doc: Y.Doc, id: Y.ID): { text: unknown; at: number } | null {
  const item = itemOf(doc, id);
  if (!item) return null;
  let at = 0;
  for (let it = item.parent instanceof Y.Text ? item.parent._start : null; it; it = it.right) {
    if (it === item) return { text: item.parent, at: at + (id.clock - item.id.clock) };
    at += it.length;
  }
  return null;
}

/** The place in the store's view of each text record of a part, in document order. */
function placesOf(view: RepairView, partName: string): Map<string, Place> {
  const placed = new Map<string, Place>();
  const part = view.package.parts.get(partName);
  if (part?.kind !== 'xml') return placed;
  type At = Omit<Place, 'order'>;
  const walk = (el: Element, at: At, parent: string | undefined): void => {
    const key = view.keyOf(partName, el.id);
    const here: At =
      el.ns === NS_A && el.local === 'p' && key && parent
        ? { p: { key, parent } }
        : el.ns === NS_A && el.local === 'r' && key
          ? { ...at, run: key }
          : at;
    for (const c of el.children) {
      if (isElement(c)) walk(c, here, key);
      else {
        const k = view.keyOf(partName, c.id);
        if (k !== undefined) placed.set(k, { order: placed.size, ...here });
      }
    }
  };
  walk(part.xml.root, {}, undefined);
  return placed;
}

/**
 * Repair the copies of one part after a peer's change. `fresh` are the copies that change
 * brought; `deleted` the root characters this replica deleted. Call inside a transaction.
 */
export function repairCopies(
  room: Room,
  partName: string,
  fresh: readonly string[],
  deleted: ReadonlySet<string>,
  view: RepairView,
  moved: MovedText,
  movedRun?: MovedRun
): void {
  const doc = room.doc;
  const me = doc.clientID;
  const entryOf = (key: string): CopyEntry | null => {
    const rec = room.nodes.get(key);
    const copy = copyOf(rec);
    const text = rec instanceof Y.Map ? rec.get(F.text) : undefined;
    if (!copy || !(text instanceof Y.Text) || (rec as Y.Map<unknown>).get(F.part) !== partName)
      return null;
    return { key, copy, text };
  };
  const arrived = new Set(
    fresh.filter((k) => {
      const c = entryOf(k);
      return c !== null && c.copy.client !== me;
    })
  );
  if (arrived.size === 0) return;
  const placed = placesOf(view, partName);
  // The live copies: a copy whose characters all moved on or went, and that the store does not
  // show, takes part in no rule (`rootOf` reads a copy of a copy from its record).
  const all: CopyEntry[] = [];
  let chars = 0;
  for (const [key, part] of room.doc.getMap(COPIES_KEY).entries()) {
    if (part !== partName) continue;
    const c = entryOf(key);
    if (!c || (c.text.length === 0 && !placed.has(key))) continue;
    if (all.length >= MAX_COPIES || (chars += c.copy.length) > MAX_CHARS) return;
    all.push(c);
  }
  const incoming = all.filter((c) => arrived.has(c.key));
  if (incoming.length === 0) return;

  const memo = new Map<string, Y.ID>();
  const root = (id: Y.ID) => rootOf(doc, id, memo);
  const copyMemo = new Map<string, Y.ID>();
  const startOf = new Map<string, { text: unknown; at: number } | null>();
  const start = (c: CopyEntry) => {
    if (!startOf.has(c.key)) startOf.set(c.key, placeOf(doc, root(c.copy.source(0))));
    return startOf.get(c.key)!;
  };
  /** > 0 when copy `a`'s split starts later than `b`'s; 0 when they cannot be told apart. */
  const laterStart = (a: CopyEntry, b: CopyEntry): number => {
    const x = start(a);
    const y = start(b);
    return x && y && x.text === y.text ? x.at - y.at : 0;
  };
  const orderOf = (c: CopyEntry) => placed.get(c.key)?.order ?? -1;
  /** The copy that keeps a character two copies hold: later start, then later in the document. */
  const better = (a: CopyEntry, b: CopyEntry): CopyEntry =>
    (laterStart(a, b) || orderOf(a) - orderOf(b)) >= 0 ? a : b;
  const ownId = (c: CopyEntry, j: number) => Y.createID(c.copy.client, c.copy.clock + j);
  /** Whether an Enter made the copy, or a copy it copies (a format of the moved text). */
  const fromEnter = (c: CopyEntry): boolean => {
    if (c.copy.newParagraph) return true;
    let at = c.copy.source(0);
    for (let depth = 0; depth < MAX_CHAIN; depth += 1) {
      const item = itemOf(doc, at);
      const copy = item ? copyOf(recordOfItem(item)) : null;
      const j = copy ? at.clock - copy.clock : -1;
      if (!copy || at.client !== copy.client || j < 0 || j >= copy.length) return false;
      if (copy.newParagraph) return true;
      at = copy.source(j);
    }
    return false;
  };

  // Live, shown copy characters by root.
  const holders = new Map<string, Holder[]>();
  for (const c of all) {
    if (!placed.has(c.key)) continue;
    for (let j = 0; j < c.copy.length; j += 1) {
      const id = ownId(c, j);
      if (isGone(doc, id)) continue;
      const r = root(c.copy.source(j));
      const k = idKey(r.client, r.clock);
      const list = holders.get(k);
      if (list) list.push({ copy: c, id });
      else holders.set(k, [{ copy: c, id }]);
    }
  }
  /** The shown copy of a root character that keeps it, or undefined. */
  const keeper = (r: Y.ID): Holder | undefined =>
    holders
      .get(idKey(r.client, r.clock))
      ?.reduce((x, y) => (better(x.copy, y.copy) === x.copy ? x : y));

  // Follow: this replica's text typed after a moved character goes after its copy.
  for (const c of incoming)
    followMoves(room, c, seenBy(room.nodes.get(c.key)), (id) => keeper(root(id)), moved);

  const doomed = new Map<Y.Text, Y.ID[]>();
  const doom = (text: Y.Text, id: Y.ID) => {
    const list = doomed.get(text);
    if (list) list.push(id);
    else doomed.set(text, [id]);
  };
  // Delete: copies of characters this replica deleted.
  if (deleted.size > 0)
    for (const c of incoming)
      for (let j = 0; j < c.copy.length; j += 1) {
        // Through copies only: what an undo put back after the delete stays (`rootOf`).
        const r = rootOf(doc, c.copy.source(j), copyMemo, false);
        if (deleted.has(idKey(r.client, r.clock)) && !isGone(doc, ownId(c, j)))
          doom(c.text, ownId(c, j));
      }
  // Once: of two shown copies of one character, where an Enter made one, one keeps it. Only the
  // characters the new copies hold: a duplicate that stood before them is not this repair's.
  const enterMemo = new Map<string, boolean>();
  const isEnter = (c: CopyEntry): boolean => {
    let out = enterMemo.get(c.key);
    if (out === undefined) enterMemo.set(c.key, (out = fromEnter(c)));
    return out;
  };
  for (const list of holders.values()) {
    if (list.length < 2 || !list.some((h) => arrived.has(h.copy.key))) continue;
    if (!list.some((h) => isEnter(h.copy))) continue;
    if (new Set(list.map((h) => h.copy.key)).size < 2) continue;
    const keep = list.reduce((x, y) => (better(x.copy, y.copy) === x.copy ? x : y));
    for (const h of list) if (h !== keep) doom(h.copy.text, h.id);
  }
  for (const [text, ids] of doomed)
    deleteChars(text, new Set(ids.map((id) => idKey(id.client, id.clock))));

  // The paragraph rules compare the roots of copies: each set is read once, every step counts.
  const roots = new Map<string, Set<string>>();
  const rootsOf = (c: CopyEntry): Set<string> => {
    let out = roots.get(c.key);
    if (!out) {
      out = new Set();
      for (let j = 0; j < c.copy.length; j += 1) {
        const r = root(c.copy.source(j));
        out.add(idKey(r.client, r.clock));
      }
      roots.set(c.key, out);
    }
    return out;
  };
  let steps = MAX_STEPS;
  const step = (n = 1) => (steps -= n) >= 0;
  const share = (a: CopyEntry, b: CopyEntry): boolean => {
    const [x, y] = [rootsOf(a), rootsOf(b)].sort((s, t) => s.size - t.size) as [
      Set<string>,
      Set<string>,
    ];
    if (!step(x.size)) return false;
    for (const r of x) if (y.has(r)) return true;
    return false;
  };
  const ctx: Repair = { room, me, incoming, all, placed, share, laterStart, step, movedRun };
  // Order: this replica's new paragraph goes after a peer's that cut the same text earlier.
  reorderParagraphs(ctx);
  followEnter(ctx);
}

/** A record's child list in the room, or undefined. */
function childrenOf(room: Room, key: string): Y.Array<string> | undefined {
  const rec = room.nodes.get(key);
  const list = rec instanceof Y.Map ? rec.get(F.children) : undefined;
  return list instanceof Y.Array ? (list as Y.Array<string>) : undefined;
}

/**
 * Follow an Enter: a peer's copy that shows in the paragraph this replica's Enter moved its own
 * copy of the same characters out of goes along: right before that copy, or right after it when
 * the peer's split starts later (a format of text after the cut). Two peers format one word
 * while one of them presses Enter before it: the word goes with the Enter, whichever format
 * keeps it. A peer formats text after the cut of this replica's Enter at the same time: the
 * formatted runs go into the new paragraph, so the text stays in order.
 */
function followEnter(ctx: Repair): void {
  const { room, me, incoming, all, placed, share, laterStart, step, movedRun } = ctx;
  // Where each text record is in the room's lists under a paragraph parent: its paragraph and
  // its run. A copy that lost to a peer's format shows nowhere, so the store cannot tell.
  const under = new Map<string, Map<string, { p: string; run: string }>>();
  const locate = (parent: string, textKey: string): { p: string; run: string } | undefined => {
    let found = under.get(parent);
    if (!found) {
      under.set(parent, (found = new Map()));
      let looks = MAX_LOOKS;
      const list = (key: string): string[] => {
        const keys = childrenOf(room, key)?.toArray() ?? [];
        looks -= keys.length;
        return looks < 0 ? [] : keys;
      };
      for (const p of list(parent))
        for (const run of list(p))
          for (const t of list(run)) for (const x of list(t)) found.set(x, { p, run });
    }
    return found.get(textKey);
  };
  // The peer's copies in the order their splits start, so the runs that go after one copy of
  // this replica keep their order: each goes after the last one placed there.
  const lastAfter = new Map<string, string>();
  for (const p of [...incoming].sort(laterStart)) {
    const pp = placed.get(p.key);
    if (!pp?.p || !pp.run) continue;
    for (const q of all) {
      if (!step()) return;
      if (q.copy.client !== me || q.key === p.key || !share(p, q)) continue;
      const at = locate(pp.p.parent, q.key);
      const qRec = at ? room.nodes.get(at.run) : undefined;
      if (!at || at.p === pp.p.key || !(qRec instanceof Y.Map)) continue;
      // This replica's Enter moved the run out of the peer's paragraph: the peer's copy goes
      // right before it. Or the Enter cut a run there and made this copy in the new paragraph,
      // and the peer's split starts after the cut: the peer's copy goes right after it.
      const before = qRec.get(LINEAGE.movedFrom) === pp.p.key;
      const after =
        !before &&
        q.copy.newParagraph &&
        laterStart(p, q) > 0 &&
        sourceParagraph(room, q, locate, pp.p.parent) === pp.p.key;
      if (!before && !after) continue;
      const from = childrenOf(room, pp.p.key);
      const to = childrenOf(room, at.p);
      const fromAt = from?.toArray().indexOf(pp.run) ?? -1;
      if (!from || !to || fromAt < 0) continue;
      const deleted = listEntryId(from, fromAt);
      from.delete(fromAt, 1);
      if (!after) {
        to.insert(to.toArray().indexOf(at.run), [pp.run]);
        break;
      }
      const anchor = lastAfter.get(at.run) ?? at.run;
      const clock = Y.getState(room.doc.store, me);
      to.insert(to.toArray().indexOf(anchor) + 1, [pp.run]);
      lastAfter.set(at.run, pp.run);
      if (deleted)
        movedRun?.(Y.createID(q.copy.client, q.copy.clock), deleted, Y.createID(me, clock), pp.run);
      break;
    }
  }
}

/**
 * The id of the live entry at `index` of a list, or undefined. Its item and the items that hold
 * it keep their content after a delete, as `Y.UndoManager` keeps what it may restore: an undo
 * of the Enter that a repair ties the entry's move to puts it back (`MovedRun`).
 */
function listEntryId(list: Y.Array<string>, index: number): Y.ID | undefined {
  let at = 0;
  for (let item = list._start; item !== null; item = item.right) {
    if (item.deleted || !item.countable) continue;
    if (index < at + item.length) {
      for (let kept: Y.Item | null = item; kept !== null && !kept.keep;) {
        kept.keep = true;
        kept = (kept.parent as Y.AbstractType<unknown>)._item;
      }
      return Y.createID(item.id.client, item.id.clock + index - at);
    }
    at += item.length;
  }
  return undefined;
}

/** The paragraph under `parent` whose run holds the first character a copy copies. */
function sourceParagraph(
  room: Room,
  c: CopyEntry,
  locate: (parent: string, textKey: string) => { p: string; run: string } | undefined,
  parent: string
): string | undefined {
  const item = itemOf(room.doc, c.copy.source(0));
  const key = item ? recordOfItem(item)?.get(F.key) : undefined;
  return typeof key === 'string' ? locate(parent, key)?.p : undefined;
}

/**
 * Follow: this replica's live characters whose left neighbor when typed was a character that
 * the copy moved (or such a character of its own) go right after the shown copy of that
 * character that keeps it. A character the copy's writer had seen (`seen`) stays: it was live
 * here and not moved there, so the writer saw it deleted. Without
 * the mark, every such character follows. A character this replica's undo or redo put back
 * keeps its identity: it goes with its restore mark, and when a shown copy holds its root, it
 * goes and is not written again ("Once"): an undo puts a character back right after the one
 * before it, so a peer's Enter that copied both leaves the copy beside it.
 */
function followMoves(
  room: Room,
  incoming: CopyEntry,
  seen: ((id: Y.ID) => boolean) | null,
  keeper: (moved: Y.ID) => Holder | undefined,
  moved: MovedText
): void {
  const doc = room.doc;
  const me = doc.clientID;
  const first = itemOf(doc, incoming.copy.source(0));
  const source = first?.parent;
  if (!(source instanceof Y.Text)) return;
  const sources = new Set<string>();
  for (let j = 0; j < incoming.copy.length; j += 1) {
    const s = incoming.copy.source(j);
    sources.add(idKey(s.client, s.clock));
  }
  const anchorOf = new Map<string, Y.ID>();
  /** A group: the characters it takes, the text it writes, and its restore marks by offset. */
  type Group = {
    anchor: Y.ID;
    ids: Y.ID[];
    /** The characters of `ids` that `text` writes again, in order. */
    written: Y.ID[];
    text: string;
    marks: [number, Y.ID][];
  };
  const groups: Group[] = [];
  let open: Group | null = null;
  for (let item = source._start; item !== null; item = item.right) {
    if (item.deleted || !(item.content instanceof Y.ContentString)) continue;
    const str = item.content.str;
    for (let j = 0; j < str.length; j += 1) {
      const id = Y.createID(item.id.client, item.id.clock + j);
      const origin = j === 0 ? item.origin : Y.createID(item.id.client, item.id.clock + j - 1);
      const ok = origin ? idKey(origin.client, origin.clock) : '';
      const anchor =
        item.id.client !== me || !origin ? undefined : sources.has(ok) ? origin : anchorOf.get(ok);
      if (!anchor) {
        open = null;
        continue;
      }
      anchorOf.set(idKey(id.client, id.clock), anchor);
      // Text typed after it still follows, in the same group.
      if (seen?.(id)) continue;
      if (!open || !Y.compareIDs(open.anchor, anchor)) {
        open = { anchor, ids: [], written: [], text: '', marks: [] };
        groups.push(open);
      }
      open.ids.push(id);
      const stands = restoredOf(doc, id);
      if (stands && keeper(id)) continue;
      if (stands) open.marks.push([open.text.length, stands]);
      open.written.push(id);
      open.text += str[j]!;
    }
  }
  for (const group of groups) {
    const best = keeper(group.anchor);
    const at = best ? liveIndex(best.copy.text, best.id) : -1;
    if (!best || at < 0) continue;
    deleteChars(source, new Set(group.ids.map((id) => idKey(id.client, id.clock))));
    if (group.text.length === 0) continue;
    const clock = Y.getState(doc.store, me);
    best.copy.text.insert(at + 1, group.text);
    // A peer that deleted the old characters and has not seen this move deletes the new ones.
    const rec = room.nodes.get(best.copy.key);
    if (rec instanceof Y.Map) markFollow(rec as Y.Map<unknown>, me, clock, group.written);
    markRestored(
      doc,
      group.marks.flatMap(([k, stands]) => [clock + k, stands.client, stands.clock, 1])
    );
    moved(group.ids[0]!, Y.createID(me, clock), group.text.length);
  }
}

/**
 * Delete, for a peer's "Follow" moves: in the text records `keys`, the text a peer moved out of
 * a character this replica deleted goes too. The peer moved it before this replica's delete
 * reached it. `deleted` holds root characters. Call inside a transaction.
 */
export function repairFollows(
  room: Room,
  keys: Iterable<string>,
  deleted: ReadonlySet<string>
): void {
  if (deleted.size === 0) return;
  const doc = room.doc;
  const memo = new Map<string, Y.ID>();
  let chars = 0;
  for (const key of keys) {
    const rec = room.nodes.get(key);
    const text = rec instanceof Y.Map ? rec.get(F.text) : undefined;
    if (!(text instanceof Y.Text)) continue;
    const doomed: Y.ID[] = [];
    for (const mark of followsOf(rec)) {
      if (mark.client === doc.clientID || (chars += mark.length) > MAX_CHARS) continue;
      for (let j = 0; j < mark.length; j += 1) {
        const own = Y.createID(mark.client, mark.clock + j);
        // Through copies only, as for a copy: what an undo put back after the delete stays.
        const r = rootOf(doc, mark.source(j), memo, false);
        if (deleted.has(idKey(r.client, r.clock)) && !isGone(doc, own)) doomed.push(own);
      }
    }
    if (doomed.length > 0)
      deleteChars(text, new Set(doomed.map((id) => idKey(id.client, id.clock))));
  }
}

/** Order: this replica's new paragraph goes after a peer's that cut the same text earlier. */
function reorderParagraphs({
  room,
  me,
  incoming,
  all,
  placed,
  share,
  laterStart,
  step,
}: Repair): void {
  for (const q of all) {
    if (q.copy.client !== me || !q.copy.newParagraph) continue;
    const qp = placed.get(q.key)?.p;
    const list = qp ? childrenOf(room, qp.parent) : undefined;
    if (!qp || !list) continue;
    const keys = list.toArray();
    if (!step(keys.length)) return;
    const index = new Map(keys.map((k, i) => [k, i]));
    const qAt = index.get(qp.key) ?? -1;
    if (qAt < 0) continue;
    let target: { key: string; at: number } | undefined;
    for (const p of incoming) {
      if (!step()) return;
      if (p.key === q.key || laterStart(q, p) <= 0) continue;
      const pp = placed.get(p.key)?.p;
      if (!pp || pp.parent !== qp.parent || pp.key === qp.key) continue;
      const pAt = index.get(pp.key) ?? -1;
      if (pAt <= qAt || (target && pAt <= target.at) || !share(p, q)) continue;
      target = { key: pp.key, at: pAt };
    }
    if (!target) continue;
    list.delete(qAt, 1);
    list.insert(list.toArray().indexOf(target.key) + 1, [qp.key]);
  }
}
