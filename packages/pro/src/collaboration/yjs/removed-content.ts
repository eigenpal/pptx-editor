/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Removed runs and paragraphs that come back with a peer's text (openspec change
// `collaboration-text-lineage`, "Character identity").
//
// A run that a write removes from a paragraph loses the characters its writer saw (they are
// deleted from its `Y.Text`), and the paragraph records the run's place (`rr.<key>`: the key of
// the sibling before it, `markRemoved` in `text-identity.ts`). A peer that typed into the run at
// the same time typed characters the removal did not see: a paragraph read puts such a run back
// at its place, with those characters only (`withRemovedRuns`). A paragraph that a write removes
// from a text body (`a:txBody` of a table cell, `p:txBody` of a shape: a delete across
// paragraphs, a join) records its place on the body the same way, and loses the characters its
// writer saw. A body read puts it back at its place while it holds a peer's text: a character,
// or a line break or run that the removal did not see. A node that a peer moved or split into it
// without the remover seeing it goes with the move or the split (`move-rescue.ts`,
// `placeEditedPieces`), unless the move rescue finds no other place for it (`stranded`); a node
// that a write moved out of it shows where it went (`movedAway`). An undo of the removal lists
// the paragraph again with all of its text, and an undo of the peer's typing empties it, so it
// goes again.
//
// Every replica must put back the same nodes. A read never stops a walk over the fields of a
// `Y.Map` part way: their order differs between replicas. It sorts what it found and then
// applies its limits. One read of a part takes one step budget (`MAX_READ_STEPS`) in one order
// on every replica: each record a read names costs a step for each of its fields, read once per
// pass (`fieldsOf`), and each record it looks at costs a step.
//
// Everything read here is untrusted: a wrong value is ignored, and every walk is bounded.

import * as Y from 'yjs';
import type { ReplicaElementSpec, ReplicaTextSpec } from '@pptx-editor.dev/core/store';
import { F, type Room } from './schema.ts';
import { isRemovalField, LINEAGE, MAX_REMOVERS, MAX_SNAPSHOT } from './text-lineage.ts';
import { IDENTITY } from './text-identity.ts';

/** The children of a removed node, and of its `a:t`, that a look for its text reads. */
const MAX_RUN_CHILDREN = 1_024;
/** The removed runs or paragraphs that one read puts back, and that one look reads. */
const MAX_REMOVED_RUNS = 1_024;
/**
 * The steps that one read of a part takes to put back removed runs and paragraphs
 * (`withRemovedRuns`). Each step looks at one record, one field, or one key; a read that runs
 * out puts back no more. Every replica takes the same steps in the same order, so every replica
 * puts back the same nodes.
 */
const MAX_READ_STEPS = 1 << 20;

const NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main';

type Spec = ReplicaElementSpec | ReplicaTextSpec;

/** What the fields of one record say, read once per pass (`fieldsOf`). */
interface Fields {
  /** A removal took the record: it has a removal snapshot of any writer. */
  readonly removed: boolean;
  /** Each node a write removed from the record (`rr.<key>`), with the sibling before it. */
  readonly places: ReadonlyMap<string, string>;
  /** The keys of `places`, sorted, the first `MAX_REMOVED_RUNS`. */
  readonly removedRuns: readonly string[];
  /** The removal snapshot fields (`rw.<client>`), sorted, the first `MAX_REMOVERS`. */
  readonly removals: readonly string[];
}

/** The step budget and the memos of one read of a part (`MAX_READ_STEPS`). */
interface Walk {
  left: number;
  readonly fields: WeakMap<Y.Map<unknown>, Fields>;
  readonly runs: Map<string, boolean>;
  readonly paragraphs: Map<string, boolean>;
  readonly seen: Map<string, ReadonlySet<string>>;
  /** Each node put back, with the parent it shows in: it comes back in one parent only. */
  readonly placed: Map<string, string>;
}

const newWalk = (): Walk => ({
  left: MAX_READ_STEPS,
  fields: new WeakMap(),
  runs: new Map(),
  paragraphs: new Map(),
  seen: new Map(),
  placed: new Map(),
});

/**
 * What the fields of a record say, read once per pass. A record a read names (`charge`) costs
 * a step for each of its fields; the read of a parent itself reads its own fields anyway.
 */
function fieldsOf(rec: Y.Map<unknown>, walk: Walk, charge: boolean): Fields {
  const known = walk.fields.get(rec);
  if (known) return known;
  let removed = false;
  let count = 0;
  const places = new Map<string, string>();
  const removals: string[] = [];
  for (const [field, value] of rec.entries()) {
    count += 1;
    if (isRemovalField(field)) {
      removed = true;
      removals.push(field);
    } else if (field.startsWith(IDENTITY.removedRun) && typeof value === 'string')
      places.set(field.slice(IDENTITY.removedRun.length), value);
  }
  if (charge) step(walk, count);
  const fields: Fields = {
    removed,
    places,
    removedRuns: [...places.keys()].sort().slice(0, MAX_REMOVED_RUNS),
    removals: removals.sort().slice(0, MAX_REMOVERS),
  };
  walk.fields.set(rec, fields);
  return fields;
}

/** Takes `n` steps; false once the read ran out. */
const step = (walk: Walk, n = 1): boolean => (walk.left -= n) >= 0;

/** The first `MAX_RUN_CHILDREN` keys of a record's child list, each once. */
function childKeys(rec: Y.Map<unknown>): string[] {
  const list = rec.get(F.children);
  if (!(list instanceof Y.Array)) return [];
  const out = new Set<string>();
  for (const k of list.slice(0, MAX_RUN_CHILDREN) as unknown[])
    if (typeof k === 'string') out.add(k);
  return [...out];
}

/** Whether a run record holds a character in its `a:t` text. */
function runHoldsText(room: Room, key: string, rec: Y.Map<unknown>, walk: Walk): boolean {
  const known = walk.runs.get(key);
  if (known !== undefined) return known;
  const holds = childKeys(rec).some((tKey) => {
    if (!step(walk)) return false;
    const t = room.nodes.get(tKey);
    if (!(t instanceof Y.Map)) return false;
    const shell = t.get(F.shell) as { ns?: unknown; local?: unknown } | undefined;
    if (shell?.ns !== NS_A || shell.local !== 't') return false;
    return childKeys(t as Y.Map<unknown>).some((xKey) => {
      if (!step(walk)) return false;
      const x = room.nodes.get(xKey);
      const text = x instanceof Y.Map ? x.get(F.text) : undefined;
      return text instanceof Y.Text && text.length > 0;
    });
  });
  walk.runs.set(key, holds);
  return holds;
}

/** Whether a run as read holds a character in its `a:t`. */
const specHoldsText = (run: ReplicaElementSpec): boolean =>
  (run.children ?? []).some(
    (t) =>
      !('text' in t) &&
      t.local === 't' &&
      (t.children ?? []).some((x) => 'text' in x && x.text.length > 0)
  );

/**
 * What the writes that removed a paragraph saw: its removal snapshots (`rw.<client>`,
 * `text-lineage.ts`), the first `MAX_REMOVERS` by writer of `MAX_SNAPSHOT` keys each (a longer
 * one is not read), and the runs a write removed from it before (they record their place
 * there). Built on first use, once per read.
 */
function removalSeen(key: string, rec: Y.Map<unknown>, walk: Walk): ReadonlySet<string> {
  const known = walk.seen.get(key);
  if (known) return known;
  const out = new Set<string>();
  const fields = fieldsOf(rec, walk, true);
  for (const field of fields.removals) {
    const value: unknown = rec.get(field);
    if (!Array.isArray(value) || value.length > MAX_SNAPSHOT || !step(walk, value.length)) continue;
    for (const k of value as unknown[]) if (typeof k === 'string') out.add(k);
  }
  for (const k of fields.removedRuns) out.add(k);
  walk.seen.set(key, out);
  return out;
}

/** Each key of a part → the records whose child lists list it. */
export type Listings = () => ReadonlyMap<string, readonly string[]>;

/**
 * The listings of every part of the room, built on first use with one walk over the room's
 * records: one read pass shares them (`Mirror.rebuild`).
 */
export function roomListings(room: Room): (partName: string) => Listings {
  let index: Map<string, Map<string, string[]>> | undefined;
  const none: ReadonlyMap<string, readonly string[]> = new Map();
  const build = (): Map<string, Map<string, string[]>> => {
    const out = new Map<string, Map<string, string[]>>();
    for (const [key, rec] of room.nodes.entries()) {
      if (!(rec instanceof Y.Map)) continue;
      const part: unknown = rec.get(F.part);
      const list: unknown = rec.get(F.children);
      if (typeof part !== 'string' || !(list instanceof Y.Array)) continue;
      let byKey = out.get(part);
      if (!byKey) out.set(part, (byKey = new Map()));
      for (const c of list.toArray() as unknown[]) {
        if (typeof c !== 'string') continue;
        const at = byKey.get(c);
        if (at) at.push(key);
        else byKey.set(c, [key]);
      }
    }
    return out;
  };
  return (partName) => () => (index ??= build()).get(partName) ?? none;
}

/**
 * What a read knows about removed paragraphs that come back (`withRemovedRuns`): the ones it
 * put back, and the ones whose moved-in nodes the move rescue could not place anywhere else
 * (`stranded`, from `rescueMoved`), so they come back with those nodes.
 */
export interface ParagraphRescue {
  readonly rescued: Set<string>;
  readonly stranded: ReadonlySet<string>;
  /** The part's listings (`roomListings`), shared by the reads of one pass. */
  readonly listings: Listings;
  /** The step budget and the memos of the pass (`newParagraphRescue`). */
  readonly walk: Walk;
}

/** The state of one read pass of a part: give each pass a new one. */
export const newParagraphRescue = (
  stranded: ReadonlySet<string>,
  listings: Listings
): ParagraphRescue => ({ rescued: new Set(), stranded, listings, walk: newWalk() });

/** One paragraph's view of a `ParagraphRescue`, with the read's step budget. */
interface Rescue {
  readonly stranded: boolean;
  readonly listings: Listings;
  readonly walk: Walk;
}

/**
 * Whether a write moved the node `key` out of the record `self` to a place the room still
 * lists: its move mark names `self` as the parent it left, and another record lists it. A mark
 * alone is not enough: an undo of a peer's move can take the node back while the mark of an
 * earlier move names the parent again, and a later delete there leaves it listed nowhere.
 */
function movedAway(child: Y.Map<unknown>, key: string, self: unknown, listings: Listings): boolean {
  if (child.get(LINEAGE.movedFrom) !== self) return false;
  return (listings().get(key) ?? []).some((parent) => parent !== self);
}

/**
 * Whether a child of a removed paragraph can hold a peer's text there: not a child that a write
 * moved out of it (it shows where it went), and not one that a peer moved or split into it
 * without the remover seeing it, unless the paragraph is `stranded`. A line break counts only
 * when the remover did not see it.
 */
function peerChild(
  child: Y.Map<unknown>,
  key: string,
  self: unknown,
  seen: () => ReadonlySet<string>,
  rescue: Rescue
): 'br' | 'run' | null {
  const shell = child.get(F.shell) as { ns?: unknown; local?: unknown } | undefined;
  if (shell?.ns !== NS_A) return null;
  if (movedAway(child, key, self, rescue.listings)) return null;
  const from: unknown = child.get(LINEAGE.movedFrom);
  // A node moved in, or a split piece made, without the remover seeing it: its text goes with
  // the move or the split (`move-rescue.ts`, `placeEditedPieces` in `text-lineage.ts`), unless
  // the move rescue found no other place for it.
  const split = typeof child.get(LINEAGE.splitFrom) === 'string';
  const moved = split || (typeof from === 'string' && !rescue.stranded);
  if (moved && !seen().has(key)) return null;
  if (shell.local === 'br') return seen().has(key) ? null : 'br';
  return shell.local === 'r' || shell.local === 'fld' ? 'run' : null;
}

/**
 * Whether a paragraph as read holds a peer's text (`peerChild`). Its remover deleted the
 * characters it saw, so a character left is a peer's. A run that a join moved away can stay
 * listed in the removed paragraph; the read shows it where it went, so the read decides.
 */
function paragraphSpecHoldsText(
  room: Room,
  p: ReplicaElementSpec,
  rec: Y.Map<unknown>,
  rescue: Rescue
): boolean {
  const self: unknown = rec.get(F.key);
  const seen = () => removalSeen(p.key, rec, rescue.walk);
  for (const c of (p.children ?? []).slice(0, MAX_RUN_CHILDREN)) {
    if ('text' in c) continue;
    if (!step(rescue.walk)) return false;
    const child = room.nodes.get(c.key);
    const kind = child instanceof Y.Map ? peerChild(child, c.key, self, seen, rescue) : null;
    if (kind === 'br' || (kind === 'run' && specHoldsText(c))) return true;
  }
  return false;
}

/**
 * Whether a paragraph record holds a peer's text (`peerChild`), also in a run that a write
 * removed from it while a peer typed into it (that run comes back with it).
 */
function paragraphHoldsText(room: Room, key: string, rec: Y.Map<unknown>, rescue: Rescue) {
  const known = rescue.walk.paragraphs.get(key);
  if (known !== undefined) return known;
  const self: unknown = rec.get(F.key);
  const removed = new Set(fieldsOf(rec, rescue.walk, true).removedRuns);
  const seen = () => removalSeen(key, rec, rescue.walk);
  const holds = [...new Set([...childKeys(rec), ...removed])].some((k) => {
    if (!step(rescue.walk)) return false;
    const c = room.nodes.get(k);
    if (!(c instanceof Y.Map)) return false;
    const kind = peerChild(c as Y.Map<unknown>, k, self, seen, rescue);
    if (kind === 'br') return true;
    // A removed run counts as `withRemovedRuns` puts it back: not one a peer moved away.
    const away = removed.has(k) && !fieldsOf(c as Y.Map<unknown>, rescue.walk, true).removed;
    return kind === 'run' && !away && runHoldsText(room, k, c as Y.Map<unknown>, rescue.walk);
  });
  rescue.walk.paragraphs.set(key, holds);
  return holds;
}

/** The nodes of a spec, counted up to `limit`. */
function sizeOf(spec: ReplicaElementSpec, limit: number): number {
  let n = 0;
  const stack: Spec[] = [spec];
  while (stack.length > 0 && n < limit) {
    const s = stack.pop()!;
    n += 1;
    if (!('text' in s)) for (const c of s.children ?? []) stack.push(c);
  }
  return n;
}

/**
 * A paragraph's children with each run that a write removed while a peer typed into it put
 * back at its place, after the sibling it followed; or a text body's children with each
 * paragraph that a write removed while a peer typed into it. Only a run or a paragraph that
 * holds a peer's text comes back: its remover deleted the characters it saw. A child that a
 * write moved out of this parent, or that the read shows here, does not. Each paragraph that
 * comes back goes into `rescue.rescued`. The read looks at the removed nodes in key order, at
 * most `MAX_REMOVED_RUNS` come back, and it stops when it ran out of steps.
 */
export function withRemovedRuns(
  room: Room,
  partName: string,
  paragraph: Y.Map<unknown>,
  children: readonly Spec[],
  read: (key: string) => ReplicaElementSpec | null,
  rescue: ParagraphRescue = newParagraphRescue(new Set(), roomListings(room)(partName))
): readonly Spec[] {
  const walk = rescue.walk;
  const places = fieldsOf(paragraph, walk, false).places;
  if (places.size === 0) return children;
  const shown = new Set(children.map((c) => c.key));
  const self: unknown = paragraph.get(F.key);
  // A paragraph puts back runs after its properties; a text body puts back paragraphs after
  // its body properties and list style.
  const body = (paragraph.get(F.shell) as { local?: unknown } | undefined)?.local === 'txBody';
  const kind = body ? 'p' : 'r';
  const lead = body ? ['bodyPr', 'lstStyle'] : ['pPr'];
  const back = new Map<string, ReplicaElementSpec>();
  // Each read of a node costs a step for each of its nodes.
  const readOnce = (key: string): ReplicaElementSpec | null => {
    const spec = read(key);
    return spec && step(walk, sizeOf(spec, MAX_READ_STEPS)) ? spec : null;
  };
  const putBack = (key: string, spec: ReplicaElementSpec) => {
    back.set(key, spec);
    walk.placed.set(key, String(self));
  };
  for (const key of [...places.keys()].sort()) {
    if (back.size >= MAX_REMOVED_RUNS || !step(walk)) break;
    // A node comes back in one parent only, the first that the read reaches.
    const placedIn = walk.placed.get(key);
    if (shown.has(key) || (placedIn !== undefined && placedIn !== String(self))) continue;
    const rec = room.nodes.get(key);
    if (!(rec instanceof Y.Map) || rec.get(F.part) !== partName) continue;
    const shell = rec.get(F.shell) as { ns?: unknown; local?: unknown } | undefined;
    if (shell?.ns !== NS_A || shell.local !== kind) continue;
    // A run a peer moved out of this paragraph shows where it went.
    if (!fieldsOf(rec as Y.Map<unknown>, walk, true).removed) continue;
    if (movedAway(rec, key, self, rescue.listings)) continue;
    if (body) {
      // A paragraph is read only when its records hold a peer's text, and the read decides.
      const one = { stranded: rescue.stranded.has(key), listings: rescue.listings, walk };
      if (!paragraphHoldsText(room, key, rec as Y.Map<unknown>, one)) continue;
      const spec = readOnce(key);
      if (!spec || !paragraphSpecHoldsText(room, spec, rec as Y.Map<unknown>, one)) continue;
      putBack(key, spec);
      rescue.rescued.add(key);
      continue;
    }
    // The lists of a run can lack the text when an undo took them; the read shows what a peer
    // typed there again (`undo-unlisted.ts`), so the read run decides.
    const listed = runHoldsText(room, key, rec as Y.Map<unknown>, walk);
    const spec = readOnce(key);
    if (spec && (listed || specHoldsText(spec))) putBack(key, spec);
  }
  if (back.size === 0) return children;
  // Each run goes after the sibling it followed; a sibling that is gone too passes it on.
  const anchorOf = (key: string): string => {
    let at = places.get(key) ?? '';
    for (let i = 0; i < MAX_REMOVED_RUNS && at !== '' && !shown.has(at) && !back.has(at); i += 1)
      at = places.get(at) ?? '';
    return shown.has(at) || back.has(at) ? at : '';
  };
  const after = new Map<string, ReplicaElementSpec[]>();
  for (const [key, spec] of back) {
    const anchor = anchorOf(key);
    const list = after.get(anchor);
    if (list) list.push(spec);
    else after.set(anchor, [spec]);
  }
  const out: Spec[] = [];
  const emitted = new Set<string>();
  const emit = (anchor: string, depth: number): void => {
    if (depth > MAX_REMOVED_RUNS) return;
    for (const spec of after.get(anchor) ?? []) {
      if (emitted.has(spec.key)) continue;
      emitted.add(spec.key);
      out.push(spec);
      emit(spec.key, depth + 1);
    }
  };
  let started = false;
  for (const c of children) {
    // Children with no sibling before them go after the parent's properties.
    if (!started && ('text' in c || !lead.includes(c.local))) {
      started = true;
      emit('', 0);
    }
    out.push(c);
    emit(c.key, 0);
  }
  if (!started) emit('', 0);
  // A run whose chain of places never reaches the paragraph (only a hostile room) goes last.
  for (const spec of back.values()) if (!emitted.has(spec.key)) out.push(spec);
  return out;
}

/**
 * A part's spec without each paragraph in `rescued` (put back by `withRemovedRuns`) that no
 * longer holds a peer's text after the part-wide passes of the read moved that text on (a
 * split piece goes after the winner's piece, `placeEditedPieces`). Every replica reads the same
 * room, so every replica drops the same paragraphs.
 */
export function withoutEmptiedRescues(
  room: Room,
  spec: ReplicaElementSpec,
  rescue: ParagraphRescue
): ReplicaElementSpec {
  if (rescue.rescued.size === 0) return spec;
  // A budget of its own: the read may have spent the pass's budget.
  const budget = newWalk();
  const keep = (c: Spec): boolean => {
    if ('text' in c || !rescue.rescued.has(c.key)) return true;
    const rec = room.nodes.get(c.key);
    const one = { stranded: rescue.stranded.has(c.key), listings: rescue.listings, walk: budget };
    return rec instanceof Y.Map && paragraphSpecHoldsText(room, c, rec as Y.Map<unknown>, one);
  };
  const walk = (s: ReplicaElementSpec): ReplicaElementSpec => {
    const children = s.children ?? [];
    if (!children.some((c) => !keep(c))) {
      const next = children.map((c) => ('text' in c ? c : walk(c)));
      return next.every((c, i) => c === children[i]) ? s : { ...s, children: next };
    }
    return { ...s, children: children.filter(keep).map((c) => ('text' in c ? c : walk(c))) };
  };
  return walk(spec);
}
