/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// The identity of text that a split moves (openspec change `collaboration-text-lineage`, "Character
// identity"). After DOCX Editor's `packages/pro/src/collaboration/document/split-text-sources.ts`
// and its paragraph text design ("Moved text keeps an identity"), adapted to PPTX Editor's run
// records. DOCX Editor commit adf0e178.
//
// A split (`splitRunAt` in `packages/core/src/automation/commands/text-model.ts`, alone for a format or
// before an Enter moves the new run into a new paragraph) deletes the text after the cut from
// the run's `Y.Text` and inserts it into a new run's `Y.Text`. The new characters are copies:
// the copy's text record names the Yjs item id of each character it copies (`ci`), so a copy
// of a copy resolves to the character first typed (its root). `cp` marks a copy that went into
// a new paragraph (an Enter). `text-repair.ts` uses the marks after a peer's change: text a peer
// typed among the moved characters follows them, a peer's delete reaches the copies, and two
// copies of one character show once. `cs` records what the split's writer had seen, so only
// text it had not seen follows. A "Follow" move marks the copy it moved text into (`cf.`), so a
// peer's delete reaches the moved text too. A replica without `cs` or `cf.` ignores them and
// keeps the behavior it had before them; every repair is a room write, so the room converges.
//
// Restored characters. An undo or a redo puts each character its step deleted back as a new
// character (a new item id). The replica that undid or redid marks each one with the character
// it stands for (`RESTORED_KEY`), so the identity of a copy survives undo and redo: a copy whose
// characters came back still resolves to the characters first typed.
//
// Removed runs and paragraphs. A run that a write removes from a paragraph, or a paragraph
// that a write removes from a text body, loses the characters its writer saw, and the parent
// records its place (`rr.<key>`: the key of the sibling before it, `markRemoved`). A read puts
// it back while it holds a peer's text (`removed-content.ts`). The place stays: a peer's undo
// can unlist the node again without a write that records it, so no write can tell that the
// place is no longer needed.
//
// Everything read here is untrusted: a wrong value is ignored, and every walk is bounded.

import * as Y from 'yjs';
import type { Element, Node, NodeId } from '@pptx-editor.dev/core/store';
import { F, type Room } from './schema.ts';

/** The record fields this module writes. */
export const IDENTITY = Object.freeze({
  /** On a copy's text record: `[client, clock, ...spans]`, see `copyOf`. */
  copyOf: 'ci',
  /** On a copy's text record: true when the copy went into a new paragraph (an Enter). */
  newParagraph: 'cp',
  /**
   * On a copy's text record: what the split's writer had seen of each typist of the copied
   * text, `[client, clock, ...]` (its state vector entries), see `seenBy`.
   */
  seen: 'cs',
  /**
   * On a copy's text record, per "Follow" move (`text-repair.ts`) into it: `cf.<client>:<clock>`
   * holds `[client, clock, ...spans]` as `copyOf` does: the moved text, and the characters it
   * replaces. A peer's delete of those characters reaches the moved text (`repairFollows`).
   */
  followed: 'cf.',
  /**
   * On a paragraph record, per removed run, and on the text body record of a table cell, per
   * removed empty paragraph: the key of the sibling before it, or ''.
   */
  removedRun: 'rr.',
});

/** The room map of copies: copy text key → part name, so a repair finds them without a scan. */
export const COPIES_KEY = 'pptx-copies-v1';

/** The longest copy a mark describes; a longer mark (only a hostile one) is ignored. */
const MAX_COPY = 1 << 20;
/** The typists one `cs` mark lists; a longer mark (only a hostile one) is ignored. */
const MAX_SEEN = 4_096;
/** The `cf.` marks of one record that a read looks at, the first ones in field order. */
const MAX_FIELDS = 1_024;
/** The longest chain of copies of copies an identity read follows. */
export const MAX_CHAIN = 32;
const NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const NS_P = 'http://schemas.openxmlformats.org/presentationml/2006/main';

/** Whether an element is a text body: `a:txBody` (a table cell) or `p:txBody` (a shape). */
export const isTextBody = (el: { readonly ns: string; readonly local: string }): boolean =>
  el.local === 'txBody' && (el.ns === NS_A || el.ns === NS_P);

/** A Yjs item id as a map key. */
export const idKey = (client: number, clock: number): string => `${client}:${clock}`;

/** A copy mark: the copy's own characters, and the id of the character each one copies. */
export interface Copy {
  /** The client and first clock of the copy's characters (one insert, so clocks run on). */
  readonly client: number;
  readonly clock: number;
  readonly length: number;
  /** The copied character of the copy's character `j`. */
  source(j: number): Y.ID;
  readonly newParagraph: boolean;
}

const isCount = (v: unknown): v is number => Number.isSafeInteger(v) && (v as number) >= 0;

const copies = new WeakMap<Y.Map<unknown>, { value: unknown; copy: Copy | null }>();

/** The copy mark of a text record, or null when it has none or a wrong one. */
export function copyOf(rec: unknown): Copy | null {
  if (!(rec instanceof Y.Map)) return null;
  const value: unknown = rec.get(IDENTITY.copyOf);
  const cached = copies.get(rec as Y.Map<unknown>);
  if (cached && cached.value === value) return cached.copy;
  const copy = parseCopy(value, rec.get(IDENTITY.newParagraph) === true);
  copies.set(rec as Y.Map<unknown>, { value, copy });
  return copy;
}

function parseCopy(value: unknown, newParagraph: boolean): Copy | null {
  if (!Array.isArray(value) || value.length < 5 || (value.length - 2) % 3 !== 0) return null;
  if (value.length > 2 + 3 * MAX_COPY) return null;
  const [client, clock] = value as unknown[];
  if (!isCount(client) || !isCount(clock)) return null;
  const starts: number[] = [];
  let length = 0;
  for (let i = 2; i < value.length; i += 3) {
    const c: unknown = value[i];
    const k: unknown = value[i + 1];
    const n: unknown = value[i + 2];
    if (!isCount(c) || !isCount(k) || !isCount(n) || n === 0) return null;
    starts.push(length);
    length += n;
    if (length > MAX_COPY) return null;
  }
  const spans = value as number[];
  return {
    client,
    clock,
    length,
    newParagraph,
    source(j) {
      // Binary search for the span that holds `j`.
      let lo = 0;
      let hi = starts.length - 1;
      while (lo < hi) {
        const mid = (lo + hi + 1) >> 1;
        if (starts[mid]! <= j) lo = mid;
        else hi = mid - 1;
      }
      const at = 2 + 3 * lo;
      return Y.createID(spans[at]!, spans[at + 1]! + (j - starts[lo]!));
    },
  };
}

const follows = new WeakMap<object, Copy | null>();

/** The "Follow" marks (`cf.`) of a text record, valid ones only, at most `MAX_FIELDS` read. */
export function followsOf(rec: unknown): Copy[] {
  const out: Copy[] = [];
  if (!(rec instanceof Y.Map)) return out;
  // The fields in one order on every replica, so each reads the same marks.
  const fields: string[] = [];
  for (const field of (rec as Y.Map<unknown>).keys())
    if (field.startsWith(IDENTITY.followed)) fields.push(field);
  for (const field of fields.sort().slice(0, MAX_FIELDS)) {
    const value: unknown = (rec as Y.Map<unknown>).get(field);
    if (!Array.isArray(value)) continue;
    let copy = follows.get(value);
    if (copy === undefined) follows.set(value, (copy = parseCopy(value, false)));
    if (copy) out.push(copy);
  }
  return out;
}

/** The mark of `rec` (its copy mark or a "Follow" mark) whose own characters hold `id`. */
function markHolding(rec: unknown, id: Y.ID): Copy | null {
  const holds = (c: Copy | null): c is Copy =>
    c !== null && c.client === id.client && id.clock >= c.clock && id.clock < c.clock + c.length;
  const copy = copyOf(rec);
  if (holds(copy)) return copy;
  return followsOf(rec).find(holds) ?? null;
}

/**
 * Mark text that a "Follow" repair moved into a copy's record: the `length` characters this
 * replica inserted from `clock` replace the characters `ids`.
 */
export function markFollow(
  rec: Y.Map<unknown>,
  client: number,
  clock: number,
  ids: readonly Y.ID[]
): void {
  rec.set(`${IDENTITY.followed}${idKey(client, clock)}`, [client, clock, ...spansOf(ids)]);
}

/**
 * The text records whose "Follow" marks a remote transaction added or brought: records it
 * created that hold one, and records where it set one.
 */
export function followsIn(room: Room, transaction: Y.Transaction): Set<string> {
  const out = new Set<string>();
  const nodes = room.nodes;
  for (const [type, keys] of transaction.changed) {
    if ((type as unknown) === nodes) {
      for (const k of keys) if (k !== null && followsOf(nodes.get(k)).length > 0) out.add(k);
      continue;
    }
    if (!(type instanceof Y.Map) || type.parent !== nodes) continue;
    const key: unknown = type.get(F.key);
    if (typeof key !== 'string') continue;
    for (const k of keys) if (k?.startsWith(IDENTITY.followed)) out.add(key);
  }
  return out;
}

/**
 * The room map of restore marks: `<client>:<clock>` of the first character an undo or a redo
 * put back → `[client, clock, length]`, the character it stands for and how many run on. One
 * entry per span, so a write adds only its own spans, and two peers never write one key.
 */
export const RESTORED_KEY = 'pptx-restored-v1';

/** The longest span a restore mark describes: one step puts back no more (`undo-restore.ts`). */
export const MAX_RESTORED = 1 << 16;
const RESTORED_ENTRY = /^(\d{1,16}):(\d{1,16})$/;

/** Restore spans by the client of the characters put back, each list sorted by first clock. */
type RestoredIndex = Map<number, (readonly [number, number, number, number])[]>;
const restoredIndex = new WeakMap<Y.Map<unknown>, RestoredIndex>();
const watched = new WeakSet<Y.Map<unknown>>();

function indexRestored(map: Y.Map<unknown>): RestoredIndex {
  let index = restoredIndex.get(map);
  if (index) return index;
  if (!watched.has(map)) {
    watched.add(map);
    map.observe(() => restoredIndex.delete(map));
  }
  index = new Map();
  for (const [key, value] of map.entries()) {
    const m = RESTORED_ENTRY.exec(key);
    if (!m || !Array.isArray(value) || value.length !== 3 || !value.every(isCount)) continue;
    const [client, clock, length] = value as [number, number, number];
    if (length === 0 || length > MAX_RESTORED) continue;
    const at = Number(m[1]);
    const first = Number(m[2]);
    if (!Number.isSafeInteger(at) || !Number.isSafeInteger(first)) continue;
    const list = index.get(at);
    const span = [first, client, clock, length] as const;
    if (list) list.push(span);
    else index.set(at, [span]);
  }
  for (const list of index.values()) list.sort((x, y) => x[0] - y[0]);
  restoredIndex.set(map, index);
  return index;
}

/** The character a restored character `id` stands for, or null when no mark names it. */
export function restoredOf(doc: Y.Doc, id: Y.ID): Y.ID | null {
  const map = doc.getMap(RESTORED_KEY);
  if (map.size === 0) return null;
  const spans = indexRestored(map).get(id.client);
  if (!spans) return null;
  let lo = 0;
  let hi = spans.length - 1;
  while (lo < hi) {
    const mid = (lo + hi + 1) >> 1;
    if (spans[mid]![0] <= id.clock) lo = mid;
    else hi = mid - 1;
  }
  const span = spans[lo];
  if (!span || id.clock < span[0] || id.clock >= span[0] + span[3]) return null;
  return Y.createID(span[1], span[2] + (id.clock - span[0]));
}

/**
 * Mark characters that this replica's undo or redo put back: `spans` of `[clock, client, clock,
 * length]`, each first new character and the character it stands for. Call inside a
 * transaction.
 */
export function markRestored(doc: Y.Doc, spans: readonly number[]): void {
  if (spans.length === 0) return;
  const map = doc.getMap(RESTORED_KEY);
  for (let i = 0; i + 3 < spans.length; i += 4)
    map.set(idKey(doc.clientID, spans[i]!), [spans[i + 1]!, spans[i + 2]!, spans[i + 3]!]);
  // A read later in this transaction sees the new marks; the observer runs only after it.
  restoredIndex.delete(map);
}

/** Compress ids into `[client, clock, length]` spans. */
function spansOf(ids: readonly Y.ID[]): number[] {
  const out: number[] = [];
  for (const id of ids) {
    const n = out.length;
    if (n > 0 && out[n - 3] === id.client && out[n - 2]! + out[n - 1]! === id.clock)
      out[n - 1]! += 1;
    else out.push(id.client, id.clock, 1);
  }
  return out;
}

/**
 * What the writer of a split had seen of the copied text's typists: the `cs` entries for every
 * client with an item in `text`, read from this replica's store.
 */
function seenOf(doc: Y.Doc, text: Y.Text): number[] | null {
  const clients = new Set<number>();
  for (let item = text._start; item !== null; item = item.right) {
    clients.add(item.id.client);
    if (clients.size > MAX_SEEN) return null;
  }
  const out: number[] = [];
  for (const client of [...clients].sort((x, y) => x - y))
    out.push(client, Y.getState(doc.store, client));
  return out;
}

const seens = new WeakMap<Y.Map<unknown>, { value: unknown; seen: Map<number, number> | null }>();

/**
 * Whether the writer of a copy had seen a character of the copied text, from the copy's `cs`
 * mark: a character it had seen and did not move was deleted in its view. Null when the record
 * has no mark or a wrong one (a room written before the mark, or a hostile one).
 */
export function seenBy(rec: unknown): ((id: Y.ID) => boolean) | null {
  if (!(rec instanceof Y.Map)) return null;
  const value: unknown = rec.get(IDENTITY.seen);
  let cached = seens.get(rec as Y.Map<unknown>);
  if (!cached || cached.value !== value) {
    cached = { value, seen: parseSeen(value) };
    seens.set(rec as Y.Map<unknown>, cached);
  }
  const seen = cached.seen;
  return seen ? (id) => id.clock < (seen.get(id.client) ?? 0) : null;
}

function parseSeen(value: unknown): Map<number, number> | null {
  if (!Array.isArray(value) || value.length % 2 !== 0 || value.length > 2 * MAX_SEEN) return null;
  const seen = new Map<number, number>();
  for (let i = 0; i < value.length; i += 2) {
    const client: unknown = value[i];
    const clock: unknown = value[i + 1];
    if (!isCount(client) || !isCount(clock) || seen.has(client)) return null;
    seen.set(client, clock);
  }
  return seen;
}

/** The live characters of `text` from `start`, `count` of them, as item ids. */
export function liveIds(text: Y.Text, start: number, count: number): Y.ID[] {
  const out: Y.ID[] = [];
  let index = 0;
  for (let item = text._start; item !== null && out.length < count; item = item.right) {
    if (item.deleted || !item.countable) continue;
    for (let j = 0; j < item.length && out.length < count; j += 1, index += 1)
      if (index >= start) out.push(Y.createID(item.id.client, item.id.clock + j));
  }
  return out;
}

/** The item that holds `id`, or null when this replica does not hold it (or it is no item). */
export function itemOf(doc: Y.Doc, id: Y.ID): Y.Item | null {
  if (Y.getState(doc.store, id.client) <= id.clock) return null;
  const struct = Y.getItem(doc.store, id);
  return struct instanceof Y.Item ? struct : null;
}

/** The text record that holds an item, or null. */
export function recordOfItem(item: Y.Item): Y.Map<unknown> | null {
  const text = item.parent;
  const rec = text instanceof Y.Text ? text.parent : null;
  return rec instanceof Y.Map ? (rec as Y.Map<unknown>) : null;
}

/** Whether the character `id` is deleted (or this replica does not hold it). */
export function isGone(doc: Y.Doc, id: Y.ID): boolean {
  const item = itemOf(doc, id);
  return !item || item.deleted;
}

/**
 * The character that `id` copies, through every copy of a copy, every "Follow" move, and every
 * character an undo or a redo put back: the character first typed. With `restores` false, a
 * character put back is its own root: a delete reaches the copies of what it deleted, never
 * what an undo put back. Memoized per call site with `memo` (one `memo` per value of
 * `restores`).
 */
export function rootOf(doc: Y.Doc, id: Y.ID, memo?: Map<string, Y.ID>, restores = true): Y.ID {
  const key = idKey(id.client, id.clock);
  const known = memo?.get(key);
  if (known) return known;
  let at = id;
  for (let depth = 0; depth < MAX_CHAIN; depth += 1) {
    const item = itemOf(doc, at);
    const copy = item ? markHolding(recordOfItem(item), at) : null;
    if (copy) {
      at = copy.source(at.clock - copy.clock);
      continue;
    }
    // A character an undo or a redo put back stands for the character it restored.
    const back = restores ? restoredOf(doc, at) : null;
    if (!back) break;
    at = back;
  }
  memo?.set(key, at);
  return at;
}

/** What a write did to one text record's `Y.Text`. */
export interface TextWrite {
  /** The record was made by this write. */
  readonly created: boolean;
  /** The characters the write deleted, in text order. */
  readonly deleted: Y.ID[];
  /** The characters it inserted in one piece: first id and length. */
  readonly inserted?: { readonly client: number; readonly clock: number; readonly length: number };
}

const isElement = (n: Node | undefined): n is Element => n?.type === 'element';
const kids = (el: Element): Element[] => el.children.filter(isElement);
const isA = (el: Element, local: string): boolean => el.ns === NS_A && el.local === local;

/** The text node of a run's `a:t`, or undefined. */
function textNodeOf(run: Element): Node | undefined {
  const t = kids(run).find((c) => isA(c, 't'));
  return t?.children.find((c) => c.type === 'text');
}

const runText = (run: Element): string => {
  const n = textNodeOf(run);
  return n?.type === 'text' ? n.value : '';
};

/**
 * After a local write: mark each run a split made with the characters it copies (`ci`), and
 * return the ids of the characters the splits moved (they are moves, not deletes). A split is a
 * run whose text became a strict prefix of its old text, with new runs whose texts add up to the
 * rest right after it, or (after an Enter) at the start of a new paragraph right after its own.
 * An Enter over a selection moves only the end of the rest; the selected characters are deleted.
 */
export function markCopies(
  room: Room,
  partName: string,
  root: Element,
  before: ReadonlyMap<NodeId, Node>,
  keyOf: ReadonlyMap<NodeId, string>,
  texts: ReadonlyMap<string, TextWrite>
): Set<string> {
  const moved = new Set<string>();
  if (before.size === 0 || texts.size === 0) return moved;
  const index = room.doc.getMap(COPIES_KEY);
  const textKey = (run: Element): string | undefined => {
    const n = textNodeOf(run);
    return n ? keyOf.get(n.id) : undefined;
  };
  const visit = (p: Element, siblings: readonly Element[], at: number): void => {
    const list = kids(p);
    list.forEach((run, i) => {
      if (!isA(run, 'r')) return;
      const old = before.get(run.id);
      if (!isElement(old) || old === run) return;
      const base = runText(old);
      const kept = runText(run);
      if (base.length <= kept.length || !base.startsWith(kept)) return;
      const origin = textKey(run);
      const write = origin === undefined ? undefined : texts.get(origin);
      const rest = base.length - kept.length;
      if (!write || write.deleted.length !== rest) return;
      const pieces: { run: Element; newParagraph: boolean }[] = [];
      let got = 0;
      const collect = (from: readonly Element[], start: number, newParagraph: boolean): boolean => {
        for (let j = start; j < from.length && got < rest; j += 1) {
          const next = from[j]!;
          if (!before.has(next.id) && isA(next, 'br')) continue;
          if (isA(next, 'pPr')) continue;
          // The paragraph's end: nothing after the split run blocks an Enter.
          if (isA(next, 'endParaRPr')) return true;
          if (!isA(next, 'r') || before.has(next.id)) return false;
          pieces.push({ run: next, newParagraph });
          got += runText(next).length;
        }
        return true;
      };
      const onlyEnd = collect(list, i + 1, false);
      // After an Enter: the rest went to the start of the new paragraph right after this one.
      const nextP = siblings[at + 1];
      const restAfter = list.slice(i + 1).every((c) => before.has(c.id) || isA(c, 'endParaRPr'));
      if (got < rest && onlyEnd && restAfter && nextP && isA(nextP, 'p') && !before.has(nextP.id))
        collect(kids(nextP), 0, true);
      // An Enter over a selection deletes the selected text first: the new paragraph holds the
      // end of the old text, and the characters before it are deleted, not moved.
      const enter = pieces.length > 0 && pieces.every((x) => x.newParagraph);
      const tail = pieces.map((x) => runText(x.run)).join('');
      if (got > rest || (got < rest && !(enter && base.endsWith(tail)))) return;
      const source = origin === undefined ? undefined : room.nodes.get(origin);
      const sourceText = source instanceof Y.Map ? source.get(F.text) : undefined;
      const seen = sourceText instanceof Y.Text ? seenOf(room.doc, sourceText) : null;
      let offset = rest - got;
      for (const piece of pieces) {
        const key = textKey(piece.run);
        const length = runText(piece.run).length;
        const own = key === undefined ? undefined : texts.get(key);
        const rec = key === undefined ? undefined : room.nodes.get(key);
        const ids = write.deleted.slice(offset, offset + length);
        offset += length;
        if (length === 0 || !own?.created || own.inserted?.length !== length) continue;
        if (!(rec instanceof Y.Map) || key === undefined) continue;
        rec.set(IDENTITY.copyOf, [own.inserted.client, own.inserted.clock, ...spansOf(ids)]);
        if (piece.newParagraph) rec.set(IDENTITY.newParagraph, true);
        if (seen) rec.set(IDENTITY.seen, seen);
        index.set(key, partName);
        for (const id of ids) moved.add(idKey(id.client, id.clock));
      }
    });
  };
  const walk = (el: Element): void => {
    if (before.get(el.id) === el) return;
    const children = kids(el);
    children.forEach((c, at) => {
      if (isA(c, 'p') && isElement(before.get(c.id))) visit(c, children, at);
    });
    for (const c of children) walk(c);
  };
  walk(root);
  return moved;
}

/**
 * After a local write: each run or paragraph the write removed loses the characters its writer
 * saw, and a removed run records its place in the paragraph it left (`rr.<key>`). Returns the
 * ids of the deleted characters.
 */
export function markRemoved(
  room: Room,
  root: Element,
  before: ReadonlyMap<NodeId, Node>,
  keyOf: ReadonlyMap<NodeId, string>
): Y.ID[] {
  const deleted: Y.ID[] = [];
  if (before.size === 0) return deleted;
  const present = new Set<NodeId>();
  const left: { node: Element; from: Element; prev: string }[] = [];
  const walk = (el: Element): void => {
    const old = before.get(el.id);
    if (old === el) return;
    for (const c of el.children) present.add(c.id);
    if (isElement(old)) {
      const kept = new Set(el.children.map((c) => c.id));
      let prev = '';
      for (const c of old.children) {
        if (!kept.has(c.id) && isElement(c) && (isA(c, 'r') || isA(c, 'p')))
          left.push({ node: c, from: el, prev });
        prev = keyOf.get(c.id) ?? prev;
      }
    }
    for (const c of el.children) if (isElement(c)) walk(c);
  };
  walk(root);
  for (const { node, from, prev } of left) {
    if (present.has(node.id)) continue;
    const key = keyOf.get(node.id);
    const parent = keyOf.get(from.id);
    const parentRec = parent === undefined ? undefined : room.nodes.get(parent);
    const place = (isA(node, 'r') && isA(from, 'p')) || (isA(node, 'p') && isTextBody(from));
    if (place && key !== undefined && parentRec instanceof Y.Map)
      parentRec.set(`${IDENTITY.removedRun}${key}`, prev);
    // The characters this writer saw go; a peer's characters typed at the same time stay.
    // A node the write keeps elsewhere (a run a join moved out of a removed paragraph) keeps
    // its text.
    const stack: Node[] = [node];
    while (stack.length > 0) {
      const n = stack.pop()!;
      if (n !== node && present.has(n.id)) continue;
      if (isElement(n)) {
        for (const c of n.children) stack.push(c);
        continue;
      }
      const k = keyOf.get(n.id);
      const rec = k === undefined ? undefined : room.nodes.get(k);
      const text = rec instanceof Y.Map ? rec.get(F.text) : undefined;
      if (!(text instanceof Y.Text) || n.type !== 'text' || n.value.length === 0) continue;
      // Only when the room holds what the writer's tree held: those are the characters it saw.
      if (text.toString() !== n.value) continue;
      for (const id of liveIds(text, 0, text.length)) deleted.push(id);
      text.delete(0, text.length);
    }
  }
  return deleted;
}
