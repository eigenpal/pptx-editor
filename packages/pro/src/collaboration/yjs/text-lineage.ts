/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0 — see packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Text that concurrent structure edits would duplicate or lose (openspec change
// `collaboration-text-lineage`). Adapted from DOCX Editor
// `packages/pro/src/collaboration/document/split-dedup.ts` and `materialize-orphans.ts` at
// 5ffe9674 (its #581): PPTX Editor's `splitRunAt` keeps the origin run, so the winner rule and
// the text comparison differ. The names differ too, because these functions work on replica
// specs and one module holds both rules: `markLineage` and `dropLosingSplits` stand for
// `SplitDedupIndex.record` and `loserRuns` with `splitWinnerOrder`, and `rescueMoved` and
// `placeEditedPieces` stand for `reportOrphans` and `adoptLooseMembers`. DOCX Editor takes the
// first replica id in string order as the winner; this module takes the earliest first cut,
// then the smaller numeric client id.
//
// Split lineage. Formatting part of a run splits it: the run keeps the text before the cut, and
// new runs take the rest. Two peers that split one run at once each mint their own new runs,
// and the merge keeps both, so the text shows twice ("worldworld") or scrambled ("Hellllo").
// A run that a split mints records the run it came from (`sf`), that run's text before the
// split (`sb`), where in it the new run starts (`sa`), how long its text was (`sl`), and its
// generation (`sg`). Each split also enters its first cut in the room's split index, which no
// undo reverts, so the read knows every peer that split one run at once, also after a peer
// deleted its pieces. When more than one peer split one run at once, a paragraph read keeps
// the pieces of the peer whose first cut is the earliest (then the smaller client id), because
// only its pieces hold all the text the merged deletes took from the run. A piece the user
// edited is never dropped: duplicated text is better than lost text.
//
// Rescue. Moving a node records where it came from (`mf`) and the chain of parents it went to
// (`mt`); removing a node records the keys its subtree held at that moment (`rw`). A node that
// a peer moved into a subtree that another peer removed without having seen it would be lost
// (a cell merge against a row delete). A read puts it back where it came from (`mi`, its index
// there), or, when that parent is gone too, at the end of the latest earlier parent that is
// still there (`mo`). The rescue itself is in `move-rescue.ts`.
//
// Everything read here comes from the room and is untrusted: wrong types are ignored, a
// rescue only moves a node the part already holds, once, and every read is linear in the room
// content it looks at, with caps on what a peer can make it look at (`MAX_*`).

import * as Y from 'yjs';
import type {
  Element,
  Node,
  NodeId,
  ReplicaElementSpec,
  ReplicaTextSpec,
} from '@pptx-editor.dev/core/store';
import { F, type Room } from './schema.ts';
import { clientOf } from './singleton-children.ts';

/** The record fields this module writes. Short names, as the other record fields. */
export const LINEAGE = Object.freeze({
  splitFrom: 'sf',
  splitBaseText: 'sb',
  splitAt: 'sa',
  splitLength: 'sl',
  splitGeneration: 'sg',
  splitEdited: 'se',
  splitBreak: 'sk',
  movedFrom: 'mf',
  movedIndex: 'mi',
  movedClock: 'mq',
  movedTo: 'mt',
  movedEarlier: 'mo',
  removedWith: 'rw',
});

/** The room map of split pieces: key → part name, so a read finds a piece no list holds. */
export const PIECES_KEY = 'pptx-pieces-v1';

/** The room map of moved nodes: key → part name, so a read finds them without a scan. */
export const MOVES_KEY = 'pptx-moves-v1';
/**
 * The room map of splits: `origin \0 base hash \0 generation \0 client` → that client's first
 * cut. It is outside the undo scope, so a group keeps every peer that split one run at once. A
 * split takes the generation after the highest one its writer has seen for the run, so splits
 * that saw each other (one after an undo of the other) never share a group.
 */
export const SPLITS_KEY = 'pptx-splits-v1';

/** A removal snapshot larger than this is not written; such a removal never rescues. */
export const MAX_SNAPSHOT = 20_000;
/**
 * Each removal writes its snapshot under its writer's own field (`rw.<client>`), so two peers
 * that remove one node at once both count. A record with more such fields than this (only a
 * hostile room) counts as seen by everything.
 */
export const MAX_REMOVERS = 64;
const REMOVED_PREFIX = `${LINEAGE.removedWith}.`;
/** The highest split generation the index reads. */
const MAX_GENERATION = 2 ** 31;
/** The longest chain of parents a move records (the read depth bounds it too). */
export const MAX_PATH = 300;
/** The most earlier parents a moved node keeps (`mo`), the latest first. */
export const MAX_EARLIER = 8;
/**
 * The most split pieces one paragraph read, or one part read, looks at. Only a hostile room
 * holds more (each piece is one format of one run); then the read keeps every piece as listed.
 */
const MAX_PIECES = 4_096;

/** A cut, a length, or an index: a safe integer of 0 or more. */
export const isCount = (v: unknown): v is number => Number.isSafeInteger(v) && (v as number) >= 0;
/** A split generation a writer can reach. */
const isGeneration = (v: unknown): v is number =>
  Number.isSafeInteger(v) && (v as number) >= 1 && (v as number) < MAX_GENERATION;

const NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const NONE: ReadonlySet<string> = new Set();

const isElement = (n: Node | undefined): n is Element => n?.type === 'element';

export const isRemovalField = (field: string): boolean =>
  field === LINEAGE.removedWith || field.startsWith(REMOVED_PREFIX);

/** Whether a removal took this record: it has a removal snapshot of any writer. */
export function isRemoved(rec: Y.Map<unknown>): boolean {
  for (const field of rec.keys()) if (isRemovalField(field)) return true;
  return false;
}

/** Records this replica's removal snapshot in its own field, beside any peer's. */
function setRemoved(rec: Y.Map<unknown>, client: string, snapshot: readonly string[]): void {
  rec.set(`${REMOVED_PREFIX}${client}`, snapshot);
}
const kids = (el: Element): Element[] => el.children.filter(isElement);

/** The text of a run (`a:r`): its `a:t`, or '' when it has none. */
function runText(run: Element): string {
  const t = kids(run).find((c) => c.local === 't');
  const text = t?.children.find((c) => c.type === 'text');
  return text?.type === 'text' ? text.value : '';
}

export type Spec = ReplicaElementSpec | ReplicaTextSpec;

/** The text of a run spec (`a:r`): its `a:t`'s text, or '' when it has none. */
function specText(run: Spec | undefined): string {
  if (!run || 'text' in run) return '';
  const t = run.children?.find((c) => !('text' in c) && c.local === 't');
  const text = t && !('text' in t) ? t.children?.find((c) => 'text' in c) : undefined;
  return text && 'text' in text ? text.text : '';
}

/** FNV-1a of a run's base text, so a split index key stays short. */
function baseHash(text: string): string {
  let h = 0x811c9dc5;
  for (let i = 0; i < text.length; i += 1) {
    h ^= text.charCodeAt(i);
    h = Math.imul(h, 0x01000193) >>> 0;
  }
  return h.toString(36);
}

/**
 * A split index key: `origin \0 base hash \0 generation \0 client`. With no client, the key of
 * the split group, which every client's entry starts with.
 */
export function splitIndexKey(
  origin: string,
  base: string,
  generation: number,
  client?: string
): string {
  const group = `${origin}\u0000${baseHash(base)}\u0000${generation}`;
  return client === undefined ? group : `${group}\u0000${client}`;
}

/** A client id as a number (keys carry it in base 36), for the "smaller client id" tie. */
const clientNumber = (client: string): number => {
  const n = Number.parseInt(client, 36);
  return Number.isFinite(n) ? n : Number.MAX_SAFE_INTEGER;
};

/**
 * After a local write: record split lineage on the runs a split minted, the origin and the
 * destination of moved nodes, and what a removed subtree held. `before` is the part's nodes
 * before the change, `keyOf` the replica's keys for both versions. Only changed elements are
 * visited: an element the change did not touch is the same object as before.
 */
export function markLineage(
  room: Room,
  partName: string,
  root: Element,
  before: ReadonlyMap<NodeId, Node>,
  keyOf: ReadonlyMap<NodeId, string>
): void {
  if (before.size === 0) return;
  const record = (id: NodeId): Y.Map<unknown> | undefined => {
    const key = keyOf.get(id);
    const rec = key === undefined ? undefined : room.nodes.get(key);
    return rec instanceof Y.Map ? (rec as Y.Map<unknown>) : undefined;
  };
  const client = room.doc.clientID.toString(36);
  // Children that left a changed element, and children that came to one from elsewhere.
  const left = new Map<NodeId, Node>();
  const arrived: { node: Node; path: readonly string[] }[] = [];
  const present = new Set<NodeId>();

  // Runs whose change this write explains as a split; any other change of a piece's text is
  // an edit by its writer, marked for good (`se`), so no reader drops it after a take-back.
  const splitOrigins = new Set<NodeId>();
  const markEdits = (p: Element): void => {
    for (const run of kids(p)) {
      if (run.local !== 'r' || splitOrigins.has(run.id)) continue;
      const old = before.get(run.id);
      if (!isElement(old) || old === run || runText(old) === runText(run)) continue;
      const rec = record(run.id);
      if (
        rec &&
        typeof rec.get(LINEAGE.splitFrom) === 'string' &&
        rec.get(LINEAGE.splitEdited) !== true
      )
        rec.set(LINEAGE.splitEdited, true);
    }
  };
  const markSplits = (p: Element): void => {
    const list = kids(p);
    list.forEach((run, i) => {
      if (run.local !== 'r' || !before.has(run.id)) return;
      const old = before.get(run.id);
      if (!isElement(old) || old === run) return;
      const base = runText(old);
      const kept = runText(run);
      if (base.length <= kept.length || !base.startsWith(kept)) return;
      const rest = base.slice(kept.length);
      const pieces: { el: Element; at: number; length: number }[] = [];
      let got = '';
      const breaks: { el: Element; at: number }[] = [];
      for (let j = i + 1; j < list.length && got.length < rest.length; j += 1) {
        const next = list[j]!;
        // A line break typed at the cut stays between the pieces (Shift+Enter in a run).
        if (next.local === 'br' && !before.has(next.id)) {
          breaks.push({ el: next, at: kept.length + got.length });
          continue;
        }
        if (next.local !== 'r' || before.has(next.id)) break;
        const text = runText(next);
        pieces.push({ el: next, at: kept.length + got.length, length: text.length });
        got += text;
      }
      const origin = keyOf.get(run.id);
      if (got !== rest || origin === undefined || pieces.length === 0) return;
      splitOrigins.add(run.id);
      const runKey = `${origin}\u0000${baseHash(base)}`;
      const generation = Math.min(
        (splitIndex(room).latest.get(runKey) ?? 0) + 1,
        MAX_GENERATION - 1
      );
      for (const { el, at, length } of pieces) {
        const rec = record(el.id);
        if (!rec || rec.has(LINEAGE.splitFrom)) continue;
        rec.set(LINEAGE.splitFrom, origin);
        rec.set(LINEAGE.splitBaseText, base);
        rec.set(LINEAGE.splitAt, at);
        rec.set(LINEAGE.splitLength, length);
        rec.set(LINEAGE.splitGeneration, generation);
        const pieceKey = keyOf.get(el.id);
        if (pieceKey !== undefined) room.doc.getMap(PIECES_KEY).set(pieceKey, partName);
      }
      // A break typed at the cut records the split and its place in the run's text, so a read
      // that drops this writer's pieces keeps the break at that place in the winner's text.
      for (const { el, at } of breaks) {
        const rec = record(el.id);
        if (!rec || rec.has(LINEAGE.splitBreak)) continue;
        rec.set(LINEAGE.splitBreak, splitIndexKey(origin, base, generation));
        rec.set(LINEAGE.splitAt, at);
      }
      room.doc.getMap(SPLITS_KEY).set(splitIndexKey(origin, base, generation, client), kept.length);
    });
  };

  const walk = (el: Element, chain: readonly string[]): void => {
    const old = before.get(el.id);
    if (old === el) return;
    const key = keyOf.get(el.id);
    const path = key === undefined ? chain : [key, ...chain].slice(0, MAX_PATH);
    const oldKids = isElement(old) ? new Set(old.children.map((c) => c.id)) : new Set<NodeId>();
    for (const c of el.children) {
      present.add(c.id);
      if (!oldKids.has(c.id) && before.has(c.id)) arrived.push({ node: c, path });
    }
    if (isElement(old)) {
      const kept = new Set(el.children.map((c) => c.id));
      for (const c of old.children) if (!kept.has(c.id)) left.set(c.id, c);
    }
    if (el.ns === NS_A && el.local === 'p' && isElement(old)) {
      markSplits(el);
      markEdits(el);
    }
    for (const c of el.children) if (isElement(c)) walk(c, path);
  };

  walk(root, []);

  // A node this replica moved: where it came from, and the parents it went to.
  if (arrived.length > 0) {
    const parentOf = new Map<NodeId, { parent: NodeId; index: number }>();
    for (const n of before.values())
      if (isElement(n))
        n.children.forEach((c, index) => parentOf.set(c.id, { parent: n.id, index }));
    const moves = room.doc.getMap(MOVES_KEY);
    // This replica's Yjs clock orders its moves, so nodes that left one index come back in the
    // order they left.
    const clock = Y.getState(room.doc.store, room.doc.clientID);
    for (const { node, path } of arrived) {
      const was = parentOf.get(node.id);
      const from = was === undefined ? undefined : keyOf.get(was.parent);
      const key = keyOf.get(node.id);
      const rec = record(node.id);
      if (!rec || !was || from === undefined || key === undefined) continue;
      // A node moved again keeps the parents it came from before, for the rescue.
      const prev: unknown = rec.get(LINEAGE.movedFrom);
      if (typeof prev === 'string' && prev !== from) {
        const old: unknown = rec.get(LINEAGE.movedEarlier);
        const list = Array.isArray(old)
          ? old.filter((k) => typeof k === 'string' && k !== prev)
          : [];
        rec.set(LINEAGE.movedEarlier, [prev, ...list].slice(0, MAX_EARLIER));
      }
      rec.set(LINEAGE.movedFrom, from);
      rec.set(LINEAGE.movedIndex, was.index);
      rec.set(LINEAGE.movedClock, clock);
      rec.set(LINEAGE.movedTo, path);
      moves.set(key, partName);
    }
  }

  // A subtree this replica removed: the keys it held, so a read knows what the removal saw.
  // The room's lists count too: they keep a split piece that this replica's read dropped. A
  // node the same write moved out of the subtree is not removed, and neither is what it holds.
  const pieceIndex = room.doc.getMap(PIECES_KEY);
  // Built on first use: the part's keys now, and the keys this replica showed before the write.
  let keysNow: Set<string> | undefined;
  let shownBefore: Set<string> | undefined;
  const stays = (k: string): boolean => (keysNow ??= keysOf(root, keyOf)).has(k);
  for (const [id, n] of left) {
    if (present.has(id)) continue;
    const held = new Set<string>();
    const stack: Node[] = [n];
    while (stack.length > 0 && held.size <= MAX_SNAPSHOT) {
      const x = stack.pop()!;
      const k = keyOf.get(x.id);
      if (k !== undefined) held.add(k);
      if (isElement(x)) for (const c of x.children) if (!present.has(c.id)) stack.push(c);
    }
    const key = keyOf.get(id);
    if (key !== undefined) roomHeld(room, key, held, stays);
    const rec = record(id);
    if (!rec || held.size > MAX_SNAPSHOT) continue;
    const snapshot = [...held];
    setRemoved(rec, client, snapshot);
    // A split piece the removal took is removed for good: no read places it again. A piece
    // that this replica's read dropped goes only with the winner's pieces of its split: while
    // one of them stays (a paragraph join moves them), the piece follows it.
    for (const k of held) {
      if (k === key || !pieceIndex.has(k) || (keysNow ??= keysOf(root, keyOf)).has(k)) continue;
      const piece = room.nodes.get(k);
      if (!(piece instanceof Y.Map) || isRemoved(piece as Y.Map<unknown>)) continue;
      shownBefore ??= new Set([...before.keys()].flatMap((id) => keyOf.get(id) ?? []));
      if (!shownBefore.has(k) && winnerStays(room, partName, piece, keysNow)) continue;
      // The piece's own key marks it removed; the subtree's snapshot is on the removed node.
      setRemoved(piece as Y.Map<unknown>, client, [k]);
    }
  }
}

/** The keys of every node of `root`. */
function keysOf(root: Element, keyOf: ReadonlyMap<NodeId, string>): Set<string> {
  const out = new Set<string>();
  const stack: Node[] = [root];
  while (stack.length > 0) {
    const n = stack.pop()!;
    const k = keyOf.get(n.id);
    if (k !== undefined) out.add(k);
    if (isElement(n)) for (const c of n.children) stack.push(c);
  }
  return out;
}

/** Whether a piece of the winner of `piece`'s split group is still in the part. */
function winnerStays(
  room: Room,
  partName: string,
  piece: Y.Map<unknown>,
  keysNow: ReadonlySet<string>
): boolean {
  const p = lineageOf(piece);
  const winner = p ? winners(room)(p.group) : undefined;
  if (!p || winner === undefined) return false;
  let looked = 0;
  for (const [key, part] of room.doc.getMap(PIECES_KEY).entries()) {
    if (part !== partName || clientOf(key) !== winner || !keysNow.has(key)) continue;
    // More pieces than any writer makes: count the winner as staying, so no piece is removed.
    if ((looked += 1) > MAX_PIECES) return true;
    if (lineageOf(room.nodes.get(key))?.group === p.group) return true;
  }
  return false;
}

/**
 * Adds the keys the room lists under `key` to `held`, but not those that `stays`, until it holds
 * more than the cap.
 */
function roomHeld(
  room: Room,
  key: string,
  held: Set<string>,
  stays: (key: string) => boolean
): void {
  const stack = [key];
  const visited = new Set<string>();
  while (stack.length > 0 && held.size <= MAX_SNAPSHOT) {
    const k = stack.pop()!;
    if (visited.has(k)) continue;
    visited.add(k);
    held.add(k);
    const rec = room.nodes.get(k);
    const list = rec instanceof Y.Map ? rec.get(F.children) : undefined;
    if (!(list instanceof Y.Array)) continue;
    for (const c of list.toArray())
      if (typeof c === 'string' && !visited.has(c) && !stays(c)) stack.push(c);
  }
}

/** The split index read by group: clients' first cuts, and the latest generation of each run. */
interface SplitIndex {
  /** `origin \0 base hash \0 generation` → client → first cut. */
  readonly groups: Map<string, Map<string, number>>;
  /** `origin \0 base hash` → the highest generation in the index. */
  readonly latest: Map<string, number>;
}

const indexCache = new WeakMap<Y.Map<unknown>, SplitIndex>();
const observed = new WeakSet<Y.Map<unknown>>();

/** The split index by group, built once per change of the index. */
function splitIndex(room: Room): SplitIndex {
  const map = room.doc.getMap(SPLITS_KEY);
  const cached = indexCache.get(map);
  if (cached) return cached;
  const index: SplitIndex = { groups: new Map(), latest: new Map() };
  for (const [entry, at] of map.entries()) {
    const parts = entry.split('\u0000');
    // A cut that no split writes (a hostile entry) would win every group: skip it.
    if (!isCount(at) || parts.length !== 4) continue;
    const [origin, hash, gen, client] = parts as [string, string, string, string];
    const generation = Number(gen);
    // A generation no writer reaches (a hostile entry) would stop the grouping: skip it.
    if (!Number.isSafeInteger(generation) || generation < 1 || generation >= MAX_GENERATION)
      continue;
    const run = `${origin}\u0000${hash}`;
    const group = `${run}\u0000${generation}`;
    let clients = index.groups.get(group);
    if (!clients) index.groups.set(group, (clients = new Map()));
    clients.set(client, at);
    index.latest.set(run, Math.max(index.latest.get(run) ?? 0, generation));
  }
  indexCache.set(map, index);
  if (!observed.has(map)) {
    observed.add(map);
    map.observe(() => indexCache.delete(map));
  }
  return index;
}

/** A split piece's lineage, read from its record; null when a field is missing or wrong. */
export interface Lineage {
  readonly from: string;
  readonly group: string;
  readonly original: string;
  /** Where in the run of origin the piece starts. */
  readonly at: number;
  readonly edited: boolean;
}

export function lineageOf(rec: unknown): Lineage | null {
  if (!(rec instanceof Y.Map)) return null;
  const from = rec.get(LINEAGE.splitFrom);
  const base = rec.get(LINEAGE.splitBaseText);
  const at = rec.get(LINEAGE.splitAt);
  const length = rec.get(LINEAGE.splitLength);
  const generation = rec.get(LINEAGE.splitGeneration);
  if (typeof from !== 'string' || typeof base !== 'string') return null;
  if (!isCount(at) || !isCount(length) || !isGeneration(generation)) return null;
  return {
    from,
    group: splitIndexKey(from, base, generation),
    original: base.slice(at, at + length),
    at,
    edited: rec.get(LINEAGE.splitEdited) === true,
  };
}

/** The winner of each split group that more than one peer split at once, computed once. */
export function winners(room: Room): (group: string) => string | undefined {
  const { groups } = splitIndex(room);
  const memo = new Map<string, string | undefined>();
  return (group) => {
    if (memo.has(group)) return memo.get(group);
    const clients = groups.get(group);
    let winner: string | undefined;
    if (clients && clients.size >= 2) {
      let best: [string, number] | undefined;
      for (const [client, at] of clients)
        if (
          !best ||
          at < best[1] ||
          (at === best[1] && clientNumber(client) < clientNumber(best[0]))
        )
          best = [client, at];
      winner = best?.[0];
    }
    memo.set(group, winner);
    return winner;
  };
}

/**
 * A piece's text with the texts of its own pieces after it (a later format of part of it), or
 * null when it or one of them is edited, missing, or in a cycle (a hostile record). Memoized
 * per read, so each piece is looked at once.
 */
function fullText(
  subs: ReadonlyMap<string, readonly string[]>,
  lineage: (key: string) => Lineage | undefined,
  text: (key: string) => string
): (key: string) => string | null {
  const memo = new Map<string, string | null>();
  const visiting = new Set<string>();
  const full = (key: string): string | null => {
    const known = memo.get(key);
    if (known !== undefined || memo.has(key)) return known ?? null;
    const p = lineage(key);
    if (!p || p.edited || visiting.has(key) || visiting.size > MAX_PIECES) return null;
    visiting.add(key);
    let out: string | null = text(key);
    for (const sub of subs.get(key) ?? []) {
      const more = full(sub);
      if (more === null) {
        out = null;
        break;
      }
      out += more;
    }
    visiting.delete(key);
    memo.set(key, out);
    return out;
  };
  return full;
}

/**
 * The pieces split from each key, in the order of their text (where each starts in the run it
 * came from), not in the order they were made: a later format of an earlier part of a piece
 * makes a piece that comes before the ones made first.
 */
function subsOf(pieces: Iterable<[string, Lineage]>): Map<string, string[]> {
  const subs = new Map<string, [string, number][]>();
  for (const [key, p] of pieces) {
    const list = subs.get(p.from);
    if (list) list.push([key, p.at]);
    else subs.set(p.from, [[key, p.at]]);
  }
  const out = new Map<string, string[]>();
  for (const [from, list] of subs)
    out.set(
      from,
      list.sort((x, y) => x[1] - y[1] || (x[0] < y[0] ? -1 : x[0] > y[0] ? 1 : 0)).map(([k]) => k)
    );
  return out;
}

/**
 * Whether `text` is `original` with characters deleted. A peer's delete of copied text reaches
 * every copy (`text-repair.ts`, "Delete"), so a piece that lost characters that way is still
 * unedited; text typed into it is not.
 */
function onlyDeletes(text: string, original: string): boolean {
  let at = 0;
  for (let i = 0; i < original.length && at < text.length; i += 1)
    if (original[i] === text[at]) at += 1;
  return at === text.length;
}

/**
 * A paragraph's (`a:p`) children with the pieces of concurrent splits of one run reduced to one
 * peer's. Only a group that more than one peer split at once is touched. In it, an unedited
 * piece is dropped when another peer won the group (the earliest first cut, then the smaller
 * client id), with the pieces it was split into. A piece counts as unedited when its text and
 * the texts of its own pieces (a later format of part of it) still add up to the text it took,
 * less characters a peer deleted, and its writer never edited it (`se`). The dropped keys go into `dropped`, which the mirror
 * keeps listed on its next write (`withDropped`).
 */
export function dropLosingSplits(
  room: Room,
  parent: { readonly ns: string; readonly local: string },
  children: readonly Spec[],
  dropped: Set<string>,
  unseenWinners: ReadonlySet<string> = NONE
): readonly Spec[] {
  if (parent.ns !== NS_A || parent.local !== 'p') return children;
  const pieces = new Map<string, Lineage>();
  const texts = new Map<string, string>();
  for (const c of children) {
    if ('text' in c) continue;
    const p = lineageOf(room.nodes.get(c.key));
    if (!p) continue;
    // More pieces than any writer makes: keep every piece as listed.
    if (pieces.size >= MAX_PIECES) return children;
    pieces.set(c.key, p);
    texts.set(c.key, specText(c));
  }
  if (pieces.size === 0) return children;
  const subs = subsOf(pieces);
  const full = fullText(
    subs,
    (k) => pieces.get(k),
    (k) => texts.get(k) ?? ''
  );
  const winnerOf = winners(room);
  const losers = new Set<string>();
  for (const [key, p] of pieces) {
    const winner = winnerOf(p.group);
    // When every piece of the winner went with a subtree that a peer removed without seeing
    // them (a paragraph join while the formats were in flight), the loser's pieces hold the text.
    if (winner === undefined || clientOf(key) === winner || unseenWinners.has(p.group)) continue;
    const text = full(key);
    if (text !== null && onlyDeletes(text, p.original)) losers.add(key);
  }
  if (losers.size === 0) return children;
  // The pieces a dropped piece was split into go with it: its text counted them.
  const queue = [...losers];
  while (queue.length > 0)
    for (const sub of subs.get(queue.pop()!) ?? [])
      if (!losers.has(sub)) {
        losers.add(sub);
        queue.push(sub);
      }
  // The mirror keeps a dropped piece listed in the room (`withDropped`), so each read checks it
  // again: if its writer types into it while the merge is in flight, it shows.
  for (const key of losers) dropped.add(key);
  const kept = children.filter((c) => 'text' in c || !losers.has(c.key));
  return withBreaksInPlace(room, kept, losers, pieces, winnerOf);
}

/**
 * A line break that a losing writer typed at its cut goes before the winner's piece that starts
 * at the same place of the run's text (or the first one after it), so the dropped pieces do not
 * move the break to after the winner's word. A break with no such piece stays where it is.
 */
function withBreaksInPlace(
  room: Room,
  children: readonly Spec[],
  losers: ReadonlySet<string>,
  pieces: ReadonlyMap<string, Lineage>,
  winnerOf: (group: string) => string | undefined
): readonly Spec[] {
  const droppedBy = new Set<string>();
  for (const key of losers) {
    const p = pieces.get(key);
    if (p) droppedBy.add(`${p.group}\u0000${clientOf(key)}`);
  }
  const moves = new Map<string, string>();
  for (const c of children) {
    if ('text' in c || c.local !== 'br') continue;
    const rec = room.nodes.get(c.key);
    if (!(rec instanceof Y.Map)) continue;
    const group = rec.get(LINEAGE.splitBreak);
    const at = rec.get(LINEAGE.splitAt);
    if (typeof group !== 'string' || !isCount(at)) continue;
    if (!droppedBy.has(`${group}\u0000${clientOf(c.key)}`)) continue;
    let best: [string, number] | undefined;
    for (const d of children) {
      const p = 'text' in d ? undefined : pieces.get(d.key);
      if (!p || p.group !== group || clientOf(d.key) !== winnerOf(group) || p.at < at) continue;
      if (!best || p.at < best[1]) best = [d.key, p.at];
    }
    if (best) moves.set(c.key, best[0]);
  }
  if (moves.size === 0) return children;
  const before = new Map<string, Spec[]>();
  for (const c of children) {
    const target = moves.get(c.key);
    if (target === undefined) continue;
    const list = before.get(target);
    if (list) list.push(c);
    else before.set(target, [c]);
  }
  const out: Spec[] = [];
  for (const c of children) {
    if (moves.has(c.key)) continue;
    for (const br of before.get(c.key) ?? []) out.push(br);
    out.push(c);
  }
  return out;
}

/**
 * A child list to write: `desired` (this replica's children), with the keys the last read
 * dropped kept at their places in `current` (the room's list), so a write never removes a
 * dropped piece that its writer may still edit.
 */
export function withDropped(
  current: readonly string[],
  desired: readonly string[],
  dropped: ReadonlySet<string> | undefined
): readonly string[] {
  if (!dropped || dropped.size === 0) return desired;
  const keep = new Set(desired);
  const out = [...desired];
  let after = -1;
  for (const key of current) {
    const at = out.indexOf(key);
    if (at >= 0 && keep.has(key)) {
      after = at;
      continue;
    }
    if (!dropped.has(key) || keep.has(key)) continue;
    out.splice(after + 1, 0, key);
    after += 1;
  }
  return out;
}

/**
 * A part's spec with each losing split piece that its writer edited moved right after the
 * winner's last piece of that split, wherever that piece is now (a peer may have joined or
 * split the paragraph while the edit was in flight), or after its run of origin when the winner
 * has none left. A piece no list holds any more is found through the piece index, so its
 * writer's text is never lost; every replica reads the same room, so all place it alike. A
 * piece that goes after another moved piece follows it; a piece whose place is never shown (a
 * cycle of hostile records) stays where it is.
 */
export function placeEditedPieces(
  room: Room,
  partName: string,
  spec: ReplicaElementSpec,
  read: ReadRecord,
  lost: ReadonlySet<string> = NONE
): ReplicaElementSpec {
  const index = room.doc.getMap(PIECES_KEY);
  if (index.size === 0) return spec;
  // Where each shown key is: its parent, its place there, and its place in document order.
  const where = new Map<string, { parent: ReplicaElementSpec; at: number; order: number }>();
  const walk = (s: ReplicaElementSpec): void => {
    (s.children ?? []).forEach((c, at) => {
      where.set(c.key, { parent: s, at, order: where.size });
      if (!('text' in c)) walk(c);
    });
  };
  walk(spec);
  const pieces = new Map<string, Lineage>();
  for (const [key, part] of index.entries()) {
    if (part !== partName) continue;
    // Only live pieces count: an undo deletes a split's records, but the index keeps its
    // entries. A piece someone removed stays removed.
    const rec = room.nodes.get(key);
    if (!(rec instanceof Y.Map) || isRemoved(rec as Y.Map<unknown>)) continue;
    const piece = lineageOf(rec);
    if (!piece) continue;
    // More live pieces than any writer makes (a hostile room): the shown ones stay where their
    // lists hold them, and the ones no list shows are still placed, up to a second cap.
    if (pieces.size >= MAX_PIECES && (where.has(key) || pieces.size >= 2 * MAX_PIECES)) continue;
    pieces.set(key, piece);
  }
  const winnerOf = winners(room);
  // A piece's text, read once and only when a rule needs it.
  const texts = new Map<string, string>();
  const textOf = (key: string): string => {
    let text = texts.get(key);
    if (text === undefined) {
      const w = where.get(key);
      text = specText(w ? w.parent.children![w.at] : (read(key, new Set(), 1) ?? undefined));
      texts.set(key, text);
    }
    return text;
  };
  const full = fullText(subsOf(pieces), (k) => pieces.get(k), textOf);
  // The winner's last shown piece of each split group, in document order: a peer may have
  // split the paragraph, so the pieces can sit in different paragraphs.
  const lastShown = new Map<string, { key: string; order: number }>();
  for (const [key, p] of pieces) {
    if (clientOf(key) !== winnerOf(p.group)) continue;
    const order = where.get(key)?.order;
    if (order === undefined) continue;
    const best = lastShown.get(p.group);
    if (!best || order > best.order) lastShown.set(p.group, { key, order });
  }
  const targets = new Map<string, string>();
  // Pieces that go right after the piece (or run) they came from, in the order of its text.
  const follow = (key: string): void => {
    const from = pieces.get(key)!.from;
    if (from !== key) targets.set(key, from);
  };
  const rescued: string[] = [];
  for (const key of [...pieces.keys()].sort()) {
    const p = pieces.get(key)!;
    const winner = winnerOf(p.group);
    // A piece that went with a subtree that a peer removed without seeing it (a paragraph join
    // while the format was in flight): an uncontested one, or the winner's (its losers are
    // dropped), goes back after the run it came from.
    if (winner === undefined || clientOf(key) === winner) {
      if (lost.has(key) && !where.has(key)) rescued.push(key);
      continue;
    }
    const text = full(key);
    if (text !== null && onlyDeletes(text, p.original)) continue;
    const after = lastShown.get(p.group)?.key ?? (where.has(p.from) ? p.from : undefined);
    if (after !== undefined && after !== key) targets.set(key, after);
  }
  const rescuedSet = new Set(rescued);
  for (const key of rescued) {
    const from = pieces.get(key)!.from;
    if (where.has(from) || rescuedSet.has(from)) follow(key);
  }
  // The pieces a moved piece was split into (a later format of part of it) move with it.
  const subs = subsOf(pieces);
  const queue = [...targets.keys()];
  while (queue.length > 0)
    for (const sub of subs.get(queue.pop()!) ?? [])
      if (!targets.has(sub) && where.has(sub)) {
        follow(sub);
        queue.push(sub);
      }
  if (targets.size === 0) return spec;
  // Take each moved piece out of its place (or read it when no list holds it).
  const taken = new Map<string, Spec>();
  const seen = new Set(where.keys());
  for (const key of targets.keys()) {
    const w = where.get(key);
    const node = w ? w.parent.children![w.at] : read(key, seen, 1);
    if (node) taken.set(key, node);
  }
  // A target shows when no piece moves it, or when it moves after a target that shows.
  const shows = new Map<string, boolean>();
  const showsAt = (key: string): boolean => {
    const known = shows.get(key);
    if (known !== undefined) return known;
    const target = targets.get(key);
    // A cycle reads as not shown: set before the walk.
    shows.set(key, false);
    const result = target === undefined ? where.has(key) : taken.has(key) && showsAt(target);
    shows.set(key, result);
    return result;
  };
  const after = new Map<string, Spec[]>();
  const moved = new Set<string>();
  for (const [key, target] of targets) {
    const node = taken.get(key);
    if (!node || !showsAt(key)) continue;
    moved.add(key);
    const list = after.get(target);
    if (list) list.push(node);
    else after.set(target, [node]);
  }
  if (moved.size === 0) return spec;
  // After one target: its own pieces in the order of its text, then the others by key.
  const rank = (target: string, key: string): number => {
    const p = pieces.get(key);
    return p?.from === target ? p.at : Infinity;
  };
  for (const [target, list] of after)
    list.sort(
      (x, y) =>
        rank(target, x.key) - rank(target, y.key) || (x.key < y.key ? -1 : x.key > y.key ? 1 : 0)
    );
  // Each moved piece goes after its target, and the pieces that go after it follow it.
  const emitted = new Set<string>();
  const emit = (key: string, out: Spec[]): void => {
    for (const node of after.get(key) ?? []) {
      if (emitted.has(node.key)) continue;
      emitted.add(node.key);
      out.push(node);
      emit(node.key, out);
    }
  };
  const rebuild = (s: ReplicaElementSpec): ReplicaElementSpec => {
    const children: Spec[] = [];
    for (const c of s.children ?? []) {
      if (moved.has(c.key)) continue;
      children.push('text' in c ? c : rebuild(c));
      emit(c.key, children);
    }
    return { ...s, children };
  };
  return rebuild(spec);
}

/** What a part read knows about its split pieces beyond one paragraph (`partLineage`). */
export interface PartLineage {
  /**
   * Split groups whose winner's pieces all went with a subtree that a peer removed without
   * seeing them, and none of them can go back (its run of origin is gone too): no loser of such
   * a group is dropped. When a winner's piece can go back, `placeEditedPieces` puts it after its
   * run of origin, and the losers are dropped as usual.
   */
  readonly unseenWinners: ReadonlySet<string>;
  /** Live pieces of the part that no list under its root holds and no removal saw. */
  readonly lost: ReadonlySet<string>;
}

export const NO_PART_LINEAGE: PartLineage = Object.freeze({
  unseenWinners: NONE,
  lost: NONE,
});

/**
 * One walk of a part before its read: which split groups lost their winner's pieces to a
 * removal that did not see them, and which pieces no list under the part's root holds. A
 * winner's piece that a removal saw, or whose records an undo deleted, still owns its group.
 * `alsoReached` holds keys that the read shows outside the room's lists (nodes that the move
 * rescue put back). Over the piece cap (only a hostile room), nothing is known: every group
 * reads as before.
 */
export function partLineage(
  room: Room,
  partName: string,
  rootKey: string,
  alsoReached: ReadonlySet<string> = NONE
): PartLineage {
  const index = room.doc.getMap(PIECES_KEY);
  if (index.size === 0) return NO_PART_LINEAGE;
  const winnerOf = winners(room);
  const owning = new Set<string>();
  const candidates = new Map<string, Lineage>();
  let live = 0;
  for (const [key, part] of index.entries()) {
    if (part !== partName) continue;
    const rec = room.nodes.get(key);
    if (!(rec instanceof Y.Map) || rec.get(F.part) !== partName) continue;
    const p = lineageOf(rec);
    if (!p) continue;
    // A removed piece still makes its winner own the group, but only live pieces count
    // against the cap: the index only grows, and users delete formatted words.
    if (isRemoved(rec as Y.Map<unknown>)) {
      if (clientOf(key) === winnerOf(p.group)) owning.add(p.group);
      continue;
    }
    if ((live += 1) > MAX_PIECES) return NO_PART_LINEAGE;
    candidates.set(key, p);
  }
  if (candidates.size === 0) return NO_PART_LINEAGE;
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
  const lost = new Set<string>();
  const lostWinners = new Map<string, string[]>();
  for (const [key, p] of candidates) {
    const winner = clientOf(key) === winnerOf(p.group);
    if (reached.has(key) || alsoReached.has(key)) {
      if (winner) owning.add(p.group);
      continue;
    }
    lost.add(key);
    if (!winner) continue;
    const list = lostWinners.get(p.group);
    if (list) list.push(key);
    else lostWinners.set(p.group, [key]);
  }
  // A lost piece can go back when its run of origin shows, or is a lost piece that can.
  const placeable = new Map<string, boolean>();
  const canPlace = (key: string): boolean => {
    const known = placeable.get(key);
    if (known !== undefined) return known;
    // A cycle (a hostile room) reads as not placeable: set before the walk.
    placeable.set(key, false);
    const from = candidates.get(key)?.from;
    const result =
      from !== undefined &&
      (reached.has(from) || alsoReached.has(from) || (lost.has(from) && canPlace(from)));
    placeable.set(key, result);
    return result;
  };
  const unseenWinners = new Set<string>();
  for (const [group, keys] of lostWinners)
    if (!owning.has(group) && !keys.some(canPlace)) unseenWinners.add(group);
  return { unseenWinners, lost };
}

/** Reads a record of the part as a spec, skipping keys in `seen`, at a read depth. */
export type ReadRecord = (
  key: string,
  seen: Set<string>,
  depth: number
) => ReplicaElementSpec | null;
