/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// The order that chooses which of two shown places keeps a character (openspec change
// `collaboration-text-lineage`, "Character identity"). Every "Once" rule uses it, so every
// replica keeps the same place, whichever rule runs there first:
//
// - Two copies of one character show, and an Enter made one of them (`repairCopies` in
//   `text-repair.ts`).
// - A character an undo put back shows beside a copy (`repairRestored` and `dropRestoredTwins`
//   in `undo-restore.ts`, and "Follow" in `text-repair.ts` for the undoing replica's own text).
//
// A place includes its source position, original copy key, creator, and Enter flag.
// Inserted text is located between its copied source neighbors before comparing boundaries.
// Equal boundaries prefer Enter, then stable creator and key ties. Undo retains that provenance;
// explicit deletion restoration starts a new deletion lifecycle. A later copy of restored text
// supersedes that restore, while copies of an earlier restore do not erase a later undo.

import * as Y from 'yjs';
import type { Element, Node, NodeId, PptxPackage } from '@pptx-editor.dev/core/store';
import type { Room } from './schema.ts';
import {
  copyOf,
  itemOf,
  MAX_CHAIN,
  recordOfItem,
  restoredOf,
  restoredStartOf,
  restoredPlaceOf,
  RESTORED_STEP_KEY,
  rootOf,
  type Copy,
} from './text-identity.ts';
import { holderItem } from './yjs-items.ts';

/** The characters `restoredStart` walks to the left. */
const MAX_WALK = 4_096;

/** What a repair reads of this replica: its keys and the package the store shows. */
export interface RepairView {
  readonly keyOf: (partName: string, id: NodeId) => string | undefined;
  readonly package: PptxPackage;
}

/** Where the store shows a text record: its paragraph (with the paragraph's parent), its run. */
export interface Place {
  readonly order: number;
  readonly p?: { readonly key: string; readonly parent: string };
  readonly run?: string;
}

/**
 * A place that may keep a character: a text record's key, the character its split starts at in
 * the text it split, and that character's root.
 */
export interface Contender {
  readonly key: string;
  readonly from: Y.ID;
  readonly start: Y.ID;
  readonly enter?: boolean;
  readonly client?: number;
}

const NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const isElement = (n: Node): n is Element => n.type === 'element';

/** Where a character stands in its text, deleted characters included, to order two starts. */
export function placeOf(doc: Y.Doc, id: Y.ID): { text: unknown; at: number } | null {
  const item = itemOf(doc, id);
  if (!item) return null;
  let at = 0;
  let remaining = MAX_WALK;
  for (
    let it = item.parent instanceof Y.Text ? item.parent._start : null;
    it && --remaining >= 0;
    it = it.right
  ) {
    if (it === item) return { text: item.parent, at: at + (id.clock - item.id.clock) };
    at += it.length;
  }
  return null;
}

/** Locate inserted text between the source characters of its copied text record. */
function sourcePlace(doc: Y.Doc, id: Y.ID): { text: unknown; at: number } | null {
  let source = id;
  let offset = 0;
  let remaining = MAX_WALK;
  for (let depth = 0; depth < MAX_CHAIN; depth += 1) {
    const item = itemOf(doc, source);
    const here = placeOf(doc, source);
    if (!item || !here) return null;
    const copy = copyOf(recordOfItem(item));
    if (!copy) return { text: here.text, at: here.at + offset };
    const j = source.clock - copy.clock;
    if (source.client === copy.client && j >= 0 && j < copy.length) {
      source = copy.source(j);
      continue;
    }
    let left: number | null = null;
    let right: number | null = null;
    for (let k = 0; k < copy.length && remaining-- > 0; k += 1) {
      const at = placeOf(doc, Y.createID(copy.client, copy.clock + k));
      if (!at || at.text !== here.text) continue;
      if (at.at < here.at) left = k;
      else {
        right = k;
        break;
      }
    }
    const anchor = left ?? right;
    if (anchor === null) return { text: here.text, at: here.at + offset };
    offset = (left !== null ? 0.5 : -0.5) + offset / 2;
    source = copy.source(anchor);
  }
  return null;
}

/** The place in the store's view of each text record of a part, in document order. */
export function placesOf(view: RepairView, partName: string): Map<string, Place> {
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

/** The order of one repair: `placed` is the store's view of the part (`placesOf`). */
export interface CopyOrder {
  readonly shown: (key: string) => boolean;
  /**
   * > 0 when `a`'s split starts later than `b`'s: in the text both split, else in the text of
   * both roots; 0 when they cannot be told apart.
   */
  readonly laterStart: (a: Contender, b: Contender) => number;
  /** > 0 when `a`'s root comes later than `b`'s in one text; 0 when they cannot be told apart. */
  readonly laterRoot: (a: Contender, b: Contender) => number;
  /** Whether `a` keeps a shared character, using source position and stable copy provenance. */
  readonly ahead: (a: Contender, b: Contender) => boolean;
}

export function copyOrder(doc: Y.Doc, placed: ReadonlyMap<string, Place>): CopyOrder {
  const places = new Map<string, { text: unknown; at: number } | null>();
  const place = (id: Y.ID) => {
    const k = `${id.client}:${id.clock}`;
    if (!places.has(k)) places.set(k, placeOf(doc, id));
    return places.get(k)!;
  };
  const later = (a: Y.ID, b: Y.ID): number | null => {
    const x = place(a);
    const y = place(b);
    return x && y && x.text === y.text ? x.at - y.at : null;
  };
  const sourcePlaces = new Map<string, ReturnType<typeof sourcePlace>>();
  const source = (id: Y.ID) => {
    const key = `${id.client}:${id.clock}`;
    if (!sourcePlaces.has(key)) sourcePlaces.set(key, sourcePlace(doc, id));
    return sourcePlaces.get(key);
  };
  const laterRoot = (a: Contender, b: Contender): number => {
    const x = source(a.start);
    const y = source(b.start);
    return x && y && x.text === y.text ? x.at - y.at : 0;
  };
  const laterStart = (a: Contender, b: Contender): number =>
    later(a.from, b.from) ?? laterRoot(a, b);
  const orderOf = (c: Contender) => placed.get(c.key)?.order ?? -1;
  const tie = (a: Contender, b: Contender): number => {
    if (later(a.start, b.start) !== 0) return 0;
    return (
      Number(!!a.enter) - Number(!!b.enter) ||
      (a.enter && b.enter ? (a.client ?? 0) - (b.client ?? 0) : 0)
    );
  };
  const stableTie = (a: Contender, b: Contender): number =>
    a.client !== undefined && b.client !== undefined
      ? b.client - a.client || (a.key < b.key ? 1 : a.key > b.key ? -1 : 0)
      : orderOf(a) - orderOf(b);
  const ahead = (a: Contender, b: Contender): boolean =>
    (laterRoot(a, b) || laterStart(a, b) || tie(a, b) || stableTie(a, b)) >= 0;
  return { laterStart, laterRoot, ahead, shown: (key) => placed.has(key) };
}

/** Whether an Enter made a copy, or a copy it copies (a format of the moved text). */
export function fromEnter(doc: Y.Doc, copy: Copy): boolean {
  if (copy.newParagraph) return true;
  let at = copy.source(0);
  for (let depth = 0; depth < MAX_CHAIN; depth += 1) {
    const item = itemOf(doc, at);
    const c = item ? copyOf(recordOfItem(item)) : null;
    const j = c ? at.clock - c.clock : -1;
    if (!c || at.client !== c.client || j < 0 || j >= c.length) return false;
    if (c.newParagraph) return true;
    at = c.source(j);
  }
  return false;
}

/** The key of a text's record in `nodes`, or null. */
export function keyOfText(room: Room, text: Y.Text): string | null {
  const holder = holderItem(text.parent);
  return holder && holder.parent === (room.nodes as unknown) && typeof holder.parentSub === 'string'
    ? holder.parentSub
    : null;
}

/**
 * The first character of the put-back text that holds character `j` of `item`: the live
 * characters right before it that `isRestored` holds, without a live character between.
 */
export function restoredStart(item: Y.Item, j: number, isRestored: (id: Y.ID) => boolean): Y.ID {
  let first = Y.createID(item.id.client, item.id.clock + j);
  const doc = item.parent instanceof Y.Text ? item.parent.doc : null;
  const recorded = doc ? restoredStartOf(doc, first) : null;
  if (recorded) return recorded;
  let looked = 0;
  let at: Y.Item | null = item;
  let k = j - 1;
  while (at !== null && looked++ < MAX_WALK) {
    if (!at.deleted && at.countable)
      for (; k >= 0; k -= 1, looked += 1) {
        const id = Y.createID(at.id.client, at.id.clock + k);
        if (!isRestored(id)) return first;
        first = id;
      }
    at = at.left;
    k = at ? at.length - 1 : -1;
  }
  return first;
}

/** A shown copy that holds a character: its record key and its mark. */
export interface HeldBy {
  readonly key: string;
  readonly copy: Copy;
}

/**
 * Whether a character put back into `text` keeps its root over the shown `copies` that hold it,
 * so they lose it: when an Enter made one of them or the run, and no copy comes before the
 * put-back text's place (`start`, from `restoredStart`) in the order.
 */
export function restoredKeeps(
  room: Room,
  order: CopyOrder,
  text: Y.Text,
  start: Y.ID,
  copies: readonly HeldBy[],
  memo: Map<string, Y.ID>,
  character: Y.ID = start
): boolean {
  const doc = room.doc;
  const key = keyOfText(room, text);
  if (key === null || !order.shown(key) || copies.length === 0) return false;
  const root = rootOf(doc, character, memo);
  const ancestorOf = (older: Y.ID, newer: Y.ID): boolean => {
    let at = newer;
    for (let n = 0; n < MAX_CHAIN; n += 1) {
      const prior = doc.getMap(RESTORED_STEP_KEY).get(`${at.client}:${at.clock}`);
      const next =
        Array.isArray(prior) && Number.isSafeInteger(prior[0]) && Number.isSafeInteger(prior[1])
          ? Y.createID(prior[0], prior[1])
          : restoredOf(doc, at);
      if (!next || Y.compareIDs(next, at)) return false;
      if (Y.compareIDs(next, older)) return true;
      at = next;
    }
    return false;
  };
  // A split of text restored by an earlier undo supersedes that restored place. A peer can
  // receive that split before the earlier restore, so root position alone cannot order them.
  for (const { copy } of copies)
    for (let j = 0; j < copy.length; j += 1) {
      let source = copy.source(j);
      if (!Y.compareIDs(rootOf(doc, source, memo), root)) continue;
      for (let depth = 0; depth < MAX_CHAIN; depth += 1) {
        if (Y.compareIDs(source, character)) return false;
        const restored = restoredOf(doc, source);
        if (restored) {
          if (!ancestorOf(source, character)) return false;
          const prior = doc.getMap(RESTORED_STEP_KEY).get(`${source.client}:${source.clock}`);
          source =
            Array.isArray(prior) && Number.isSafeInteger(prior[0]) && Number.isSafeInteger(prior[1])
              ? Y.createID(prior[0], prior[1])
              : restored;
          continue;
        }
        const item = itemOf(doc, source);
        const parent = item ? copyOf(recordOfItem(item)) : null;
        const offset = parent ? source.clock - parent.clock : -1;
        if (!parent || source.client !== parent.client || offset < 0 || offset >= parent.length)
          break;
        source = parent.source(offset);
      }
    }
  const step = doc.getMap(RESTORED_STEP_KEY).get(`${character.client}:${character.clock}`);
  if (Array.isArray(step) && step[2] === false) return true;
  // The put-back text stands where the split its undo reverted started: at the character its
  // first character stands for.
  const provenance = restoredPlaceOf(doc, character);
  const at: Contender = {
    key: provenance?.key ?? key,
    from: standsFor(doc, start),
    start: rootOf(doc, start, memo),
    enter: provenance?.enter,
    client: provenance?.client,
  };
  return !copies.some((c) => order.ahead(contenderOf(doc, c, memo), at));
}

/** A shown copy's place in the order. */
export function contenderOf(doc: Y.Doc, c: HeldBy, memo: Map<string, Y.ID>): Contender {
  const from = c.copy.source(0);
  return {
    key: c.key,
    from: standsFor(doc, from),
    start: rootOf(doc, from, memo),
    enter: fromEnter(doc, c.copy),
    client: c.copy.client,
  };
}

/**
 * The character that `id` stands for through restore marks only: an undo or a redo puts a
 * character back as a new one, and both stand at the place of the first.
 */
function standsFor(doc: Y.Doc, id: Y.ID): Y.ID {
  let at = id;
  for (let depth = 0; depth < MAX_CHAIN; depth += 1) {
    const back = restoredOf(doc, at);
    if (!back) break;
    at = back;
  }
  return at;
}
