/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// The package mirror: one replica's two-way mapping between its store and the room.
//
// Local → room (`seed`, `writeChange`): walk each changed part. A node whose object is the same
// as before the change is an untouched subtree (the store shares structure), so the walk skips
// it. Otherwise the node's record gets the smallest writes that make it equal: attribute sets
// and deletes, a text diff on its `Y.Text`, and a list diff on its child keys. Concurrent
// typing in one text node therefore merges character by character, as in DOCX Editor.
//
// Every new node gets a minted key (the seed marks its own, `client._N`). Two peers that both
// add one schema singleton leave two records; a read shows the first and keeps the others as
// its companions (`withCompanions` in `singleton-children.ts`), and puts the children of a
// listed parent back in schema order. A write keeps the companions listed after their winner
// and sends each value to the copy that holds it (`MergeView`).
//
// Room → local (`rebuild`): read the records of each touched part into a replica spec and
// rebuild the part with core's `buildReplicaPart`, which validates it and keeps this replica's
// node ids for keys it already knows.
//
// A key that two lists hold (two peers moved one node at once) shows at its first place in
// document order. The read notes the index of each other listing, and the session deletes
// exactly those before the next local write of the part (`dropHiddenListings`). A hidden
// listing must not outlive that write: a write that removes or moves the shown place would let
// it show for peers, while the writer's store, which the store op made, keeps none.

import * as Y from 'yjs';
import { readText } from './text-case.ts';
import { locateCharacter } from './text-position.ts';
import { withRemovedTableText } from './removed-table-content.ts';
import { repairKnownDeletes } from './text-repair.ts';
import { touchedParts } from './room-events.ts';
import { changedCopies } from './copy-events.ts';
import { pendingChildren } from './pending-content.ts';
import {
  buildReplicaPart,
  type Element,
  type Node,
  type NodeId,
  type PackagePart,
  type PptxPackage,
  type ReplicaElementSpec,
  type ReplicaTextSpec,
  type XmlPart,
} from '@pptx-editor.dev/core/store';
import { ATTR_SEP, F, MAX_MINT_COUNTER, ROOM_LIMITS, partEntryOf, type Room } from './schema.ts';
import {
  SEED_MARK,
  clientOf,
  groupOf,
  inSchemaOrder,
  withCompanions,
  type MergeView,
} from './singleton-children.ts';
import {
  dropLosingSplits,
  markLineage,
  NO_PART_LINEAGE,
  isRemoved,
  partLineage,
  placeEditedPieces,
  type PartLineage,
  withDropped,
} from './text-lineage.ts';
import { rescueMoved, sameLineage, specKeys } from './move-rescue.ts';
import { elementsByKey, markMergedCells, withMergedText, withoutKey } from './merged-cells.ts';
import {
  idKey,
  isTextBody,
  markCopies,
  markRemoved,
  rootOf,
  deletionRoot,
  DELETE_CLOCK_KEY,
} from './text-identity.ts';
import {
  newParagraphRescue,
  type ParagraphRescue,
  roomListings,
  withoutEmptiedRescues,
  withRemovedRuns,
} from './removed-content.ts';
import { noteTidiedListings, UNDONE_KEY, UndoneListings } from './undo-unlisted.ts';
import {
  indexTree,
  inRoomOrder,
  rescuedParents,
  type RescuedIn,
  shownParentsOnly,
  shownParentKey,
  shownPlaceKept,
  syncList,
  syncText,
  type PartWrite,
  type SkippedListings,
} from './mirror-support.ts';

const NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main';

/** The root characters one replica remembers deleting (`text-repair.ts`); a cap on memory. */
const MAX_DELETED = 1 << 20;

interface PartIds {
  readonly keyOf: Map<NodeId, string>;
  readonly idOf: Map<string, NodeId>;
}

/** The outcome of reading touched parts back from the room. */
export type RebuildResult =
  | { readonly ok: true; readonly parts: PackagePart[]; readonly removed: string[] }
  | { readonly ok: false; readonly reason: string };

const attrOfElement = (el: Element, local: string): string | undefined =>
  el.attrs.find((a) => a.ns === '' && a.local === local)?.value;

const shellOf = (el: Element) => ({
  ns: el.ns,
  local: el.local,
  ...(el.prefix ? { prefix: el.prefix } : {}),
  nsDecls: el.nsDecls.map((d) => ({ prefix: d.prefix, uri: d.uri })),
});

/** The last minted key number of one client on one Y.Doc, shared by its mirrors. */
export interface MintCounter {
  value: number;
}

export class PackageMirror {
  readonly #room: Room;
  readonly #ids = new Map<string, PartIds>();
  /** The merge of each property container that won a concurrent group, by its key. */
  readonly #views = new Map<string, MergeView>();
  /** The keys of `#views` per part, to replace them when the part rebuilds. */
  readonly #viewsOfPart = new Map<string, Set<string>>();
  /** Every hidden copy: it stays listed in its parent and in its own record. */
  readonly #hidden = new Set<string>();
  /** Per part, the split pieces and cell bodies its last read dropped; a write keeps them. */
  readonly #dropped = new Map<string, Set<string>>();
  /** Per part, the text bodies whose merged text its last read moved (`merged-cells.ts`). */
  readonly #mergedBodies = new Map<string, Set<string>>();
  /**
   * Per part, the nodes the last read showed again after an undo unlisted them
   * (`undo-unlisted.ts`). A write never lists them: they show while a peer's content is in them.
   */
  readonly #unlisted = new Map<string, ReadonlyMap<string, string>>();
  /**
   * Per part, the runs the last read showed only by the move rescue, each with the parent it
   * showed them in (`rescuedParents`). A write never lists them there: they stay listed where
   * the room has them, so an undo of the removal that sent them shows them there again.
   */
  readonly #rescuedIn = new Map<string, RescuedIn>();
  /** The parts whose last local write changed a parent of a rescued run, or its parent. */
  readonly #touchedRescue = new Set<string>();
  readonly #changedDrawing = new Set<string>();
  #reading = new Set<string>();
  /** Per part, the listings its last read skipped, by parent key (`dropHiddenListings`). */
  readonly #hiddenListings = new Map<string, Map<string, SkippedListings>>();
  /** The skipped listings of the part being read. */
  #hiddenNow = new Map<string, SkippedListings>();
  /** The part being read: its split pieces beyond one paragraph (`partLineage`). */
  #lineage: PartLineage = NO_PART_LINEAGE;
  /** The part being read: the nodes an undo unlisted that show again (`undo-unlisted.ts`). */
  #undone: UndoneListings | null = null;
  /** The part being read: removed paragraphs that come back (`withRemovedRuns`). */
  #rescue: ParagraphRescue | undefined;
  /** Each hidden copy → the winner it is a companion of. */
  readonly #winnerOf = new Map<string, string>();
  /** The part of each key this replica knows, for records that a remote undo deletes. */
  readonly #partOfKey = new Map<string, string>();
  readonly #counter: MintCounter;
  /** The root characters this replica's own writes deleted (not moved), for `text-repair.ts`. */
  readonly #deleted = new Set<string>();
  readonly #deletedAt = new Map<string, number>();

  /**
   * `counter` is the session's, so a second attach on one Y.Doc (same client id) counts on
   * from the keys the first attach minted, also those whose records an undo has deleted.
   */
  constructor(room: Room, counter: MintCounter = { value: 0 }) {
    this.#room = room;
    this.#counter = counter;
    // Another session on the same Y.Doc: count on from this client's highest key in the room.
    // A peer may write any key, so a number above the cap does not count.
    const prefix = `${room.doc.clientID.toString(36)}.`;
    for (const key of room.nodes.keys()) {
      if (!key.startsWith(prefix)) continue;
      const n = Number.parseInt(key.slice(prefix.length).replace(SEED_MARK, ''), 36);
      if (Number.isSafeInteger(n) && n <= MAX_MINT_COUNTER && n > counter.value) counter.value = n;
    }
  }

  /** The room has content to adopt. */
  get seeded(): boolean {
    return this.#room.parts.size > 0;
  }

  /** Replace the merge views of a part. */
  #setViews(name: string, views: Map<string, MergeView>): void {
    for (const key of this.#viewsOfPart.get(name) ?? []) this.#views.delete(key);
    for (const [key, view] of views) this.#views.set(key, view);
    this.#viewsOfPart.set(name, new Set(views.keys()));
    this.#hidden.clear();
    this.#winnerOf.clear();
    for (const [winner, view] of this.#views)
      for (const c of view.companions) {
        this.#hidden.add(c);
        this.#winnerOf.set(c, winner);
      }
  }

  /** Replace a part's key mapping, and the part of each of its keys. */
  #setIds(name: string, ids: PartIds | undefined): void {
    for (const key of this.#ids.get(name)?.idOf.keys() ?? []) this.#partOfKey.delete(key);
    if (!ids) {
      this.#ids.delete(name);
      return;
    }
    this.#ids.set(name, ids);
    for (const key of ids.idOf.keys()) this.#partOfKey.set(key, name);
  }

  #partIds(partName: string): PartIds {
    let ids = this.#ids.get(partName);
    if (!ids) {
      ids = { keyOf: new Map(), idOf: new Map() };
      this.#ids.set(partName, ids);
    }
    return ids;
  }

  /** True while `seed` writes the opened deck: its keys say so (`client._N`, `fromSeed`). */
  #seeding = false;

  /** A key no record in the room holds. */
  #mint(): string {
    const client = this.#room.doc.clientID.toString(36);
    const mark = this.#seeding ? SEED_MARK : '';
    let key: string;
    do {
      this.#counter.value += 1;
      key = `${client}.${mark}${this.#counter.value.toString(36)}`;
    } while (this.#room.nodes.has(key));
    return key;
  }

  /** Write every part of `pkg` into an empty room. Call inside a local transaction. */
  seed(pkg: PptxPackage): void {
    this.#seeding = true;
    try {
      for (const part of pkg.parts.values()) this.#writePart(part, undefined);
    } finally {
      this.#seeding = false;
    }
  }

  /** Whether the last read of one of `parts` skipped a listing (`dropHiddenListings`). */
  hasHiddenListings(parts: readonly string[]): boolean {
    return parts.some((name) => this.#hiddenListings.has(name));
  }

  /**
   * Delete the listings that the last read of `parts` skipped, so the room lists each node only
   * where every replica shows it. The read shows a key at its first place in document order and
   * skips every other listing; a write that removes or moves that place would otherwise let the
   * next listing show for every peer. A skipped listing never shows, so no read changes, and
   * the caller writes this in its own transaction, outside the undo scope: an undo never lists
   * the node twice again.
   *
   * A deleted listing of a node whose shown place this write keeps (`before` → `after`) is
   * recorded as an undo records one: the shown listing can be a peer's undo of a removal, and
   * that peer's redo deletes it again, so the node then shows at the deleted listing. A node
   * whose shown place this write removes or moves goes with it.
   *
   * Only the exact listings the read skipped go. A transaction that changes the part rebuilds
   * it and notes them again, so the room still holds them at their indexes; a list that
   * differs from what the read saw keeps every entry.
   */
  dropHiddenListings(parts: readonly string[], before: PptxPackage, after: PptxPackage): void {
    for (const name of parts) {
      const hidden = this.#hiddenListings.get(name);
      if (!hidden) continue;
      const keeps = shownPlaceKept(
        this.#ids.get(name)?.idOf,
        before.parts.get(name),
        after.parts.get(name)
      );
      const ids = this.#ids.get(name);
      const shownIn = shownParentKey(ids?.idOf, ids?.keyOf, before.parts.get(name));
      this.#hiddenListings.delete(name);
      for (const [parent, skipped] of hidden) {
        const list = this.#childrenOf(parent);
        if (!list || list.length !== skipped.length) continue;
        const at = [...skipped.at.keys()].sort((x, y) => y - x);
        if (at.some((i) => list.get(i) !== skipped.at.get(i))) continue;
        // A listing whose shown place this write keeps can show again later (`undo-unlisted.ts`).
        const kept = at.filter((i) => keeps(skipped.at.get(i)!));
        if (kept.length > 0)
          noteTidiedListings(this.#room, name, parent, list.toArray(), kept, shownIn);
        // From the end, one delete per run of adjacent indexes, so a delete does not move the
        // indexes still to delete and a hostile list of repeats costs one delete.
        for (let run = 0; run < at.length;) {
          let next = run + 1;
          while (next < at.length && at[next] === at[next - 1]! - 1) next++;
          list.delete(at[next - 1]!, next - run);
          run = next;
        }
      }
    }
  }

  /** Write one local change. Call inside a local transaction. */
  writeChange(
    before: PptxPackage,
    after: PptxPackage,
    parts: readonly string[],
    caseChange = false
  ): void {
    for (const name of parts) {
      this.#touchedRescue.delete(name);
      this.#changedDrawing.delete(name);
      const next = after.parts.get(name);
      if (!next) {
        this.#room.parts.delete(name);
        this.#room.binaries.delete(name);
        this.#setIds(name, undefined);
        continue;
      }
      this.#writePart(next, before.parts.get(name), caseChange);
    }
  }

  #writePart(part: PackagePart, before: PackagePart | undefined, caseChange = false): void {
    if (part.kind === 'binary') {
      if (before?.kind !== 'binary' || before.bytes !== part.bytes)
        this.#room.binaries.set(part.partName, part.bytes);
      if (partEntryOf(this.#room.parts.get(part.partName))?.kind !== 'binary')
        this.#room.parts.set(part.partName, { kind: 'binary' });
      return;
    }
    if (before?.kind === 'binary') this.#room.binaries.delete(part.partName);
    const write: PartWrite = {
      partName: part.partName,
      caseChange,
      ids: this.#partIds(part.partName),
      before: indexTree(before?.kind === 'xml' ? before.xml.root : undefined),
      texts: new Map(),
      materialize: new Set(),
      force: new Set(),
    };
    const afterNodes = indexTree(part.xml.root);
    const rescued = this.#rescuedIn.get(part.partName)?.inParent;
    const removedKeys = new Set<string>();
    for (const [id] of write.before) {
      if (afterNodes.has(id)) continue;
      const key = write.ids.keyOf.get(id);
      if (key) removedKeys.add(key);
      const children = key ? this.#childrenOf(key) : undefined;
      if (!children) continue;
      for (const child of children.toArray()) if (rescued?.has(child)) write.materialize.add(child);
    }
    for (const [key, value] of this.#room.doc.getMap(UNDONE_KEY)) {
      const entry = value as { s?: unknown } | undefined;
      if (typeof entry?.s === 'string' && removedKeys.has(entry.s)) write.materialize.add(key);
    }
    if (write.materialize.size > 0) {
      const force = (node: Node): boolean => {
        const children = node.type === 'element' ? node.children.map(force).some(Boolean) : false;
        const own = write.ids.keyOf.get(node.id);
        const needed = children || (own !== undefined && write.materialize.has(own));
        if (needed) write.force.add(node.id);
        return needed;
      };
      force(part.xml.root);
    }
    const root = this.#syncNode(part.xml.root, write);
    // A read moved merged text, and the room lists it where it was: the room takes the child
    // lists of those text bodies and cells as this store holds them, so an edit that uncovers
    // that place shows it once everywhere. Only the lists: other nodes keep the room's content.
    const moved = this.#mergedBodies.get(part.partName);
    this.#mergedBodies.delete(part.partName);
    for (const body of elementsByKey(part.xml.root, moved, write.ids.keyOf))
      this.#syncNode(body, { ...write, before: withoutKey(write.before, body.id) });
    markLineage(this.#room, part.partName, part.xml.root, write.before, write.ids.keyOf);
    markMergedCells(this.#room, part.xml.root, write.before, write.ids.keyOf);
    this.#markIdentity(part.partName, part.xml.root, write);
    const entry = partEntryOf(this.#room.parts.get(part.partName));
    if (entry?.kind !== 'xml' || entry.root !== root)
      this.#room.parts.set(part.partName, { kind: 'xml', root });
  }

  /** Copy marks for this write's splits, and the characters it deleted (`text-identity.ts`). */
  #markIdentity(partName: string, root: Element, write: PartWrite): void {
    if (this.#seeding) return;
    const moved = markCopies(
      this.#room,
      partName,
      root,
      write.before,
      write.ids.keyOf,
      write.texts
    );
    const removed = markRemoved(this.#room, root, write.before, write.ids.keyOf);
    const doc = this.#room.doc;
    const memo = new Map<string, Y.ID>();
    let deletionClock: number | undefined;
    const removedRoots = new Set<string>();
    const note = (id: Y.ID) => {
      if (this.#deleted.size >= MAX_DELETED || moved.has(idKey(id.client, id.clock))) return;
      const r = deletionRoot(doc, id, memo);
      deletionClock ??= this.#tickDeletion();
      this.#deleted.add(idKey(r.client, r.clock));
      this.#deletedAt.set(idKey(r.client, r.clock), deletionClock);
      const full = rootOf(doc, id);
      removedRoots.add(idKey(full.client, full.clock));
      this.#deleted.add(idKey(full.client, full.clock));
      this.#deletedAt.set(idKey(full.client, full.clock), deletionClock);
    };
    for (const t of write.texts.values()) for (const id of t.deleted) note(id);
    for (const id of removed) note(id);
    repairKnownDeletes(this.#room, partName, removedRoots);
  }

  #tickDeletion(): number {
    const doc = this.#room.doc;
    const clock = Y.getState(doc.store, doc.clientID);
    doc.getMap(DELETE_CLOCK_KEY).set(String(doc.clientID), clock);
    return clock;
  }

  get deletedAt(): ReadonlyMap<string, number> {
    return this.#deletedAt;
  }

  noteRejected(id: string, source: string): void {
    if (this.#deleted.size >= MAX_DELETED) return;
    this.#deleted.add(id);
    const clock = this.#deletedAt.get(source);
    if (clock !== undefined) this.#deletedAt.set(id, clock);
  }

  /** The root characters this replica deleted. */
  get deletedText(): ReadonlySet<string> {
    return this.#deleted;
  }

  /**
   * The parts among `names` whose last read showed a node only for a peer's content after an
   * undo (`undo-unlisted.ts`), or whose last write changed a parent that shows a node only by
   * the move rescue. A local write can take that content, and the store keeps the node, so
   * such a part is read again.
   */
  showingUnlisted(names: readonly string[]): string[] {
    return names.filter(
      (name) =>
        this.#changedDrawing.has(name) ||
        (this.#unlisted.get(name)?.size ?? 0) > 0 ||
        this.#touchedRescue.has(name)
    );
  }

  /** After an undo or redo: a delete spares copies of `restored` roots and reaches `deleted`. */
  noteStep(restored: Iterable<string>, deleted: Iterable<string>): void {
    for (const r of restored) {
      this.#deleted.delete(r);
      this.#deletedAt.delete(r);
    }
    let clock: number | undefined;
    for (const r of deleted)
      if (this.#deleted.size < MAX_DELETED) {
        clock ??= this.#tickDeletion();
        this.#deleted.add(r);
        this.#deletedAt.set(r, clock);
      }
  }

  /**
   * The copies (`text-identity.ts`) that a remote transaction added, by part: the text records
   * it created that carry a copy mark of another client.
   */
  newCopies(transaction: Y.Transaction, includeChangedText = true): Map<string, string[]> {
    return changedCopies(this.#room, transaction, includeChangedText);
  }

  #syncNode(node: Node, write: PartWrite): string {
    const { partName, ids, before } = write;
    let key = ids.keyOf.get(node.id);
    let record = key === undefined ? undefined : this.#room.nodes.get(key);
    if (
      key !== undefined &&
      record instanceof Y.Map &&
      before.get(node.id) === node &&
      !write.force.has(node.id)
    )
      return key;
    if (node.type === 'element' && (node.ns === NS_A || isTextBody(node)))
      this.#changedDrawing.add(partName);
    if (key !== undefined && this.#rescuedIn.get(partName)?.watch.has(key))
      this.#touchedRescue.add(partName);
    const created = key === undefined || !(record instanceof Y.Map);
    if (key === undefined || !(record instanceof Y.Map)) {
      key = this.#mint();
      const created = new Y.Map<unknown>();
      this.#room.nodes.set(key, created);
      created.set(F.part, partName);
      created.set(F.key, key);
      if (node.type === 'text') created.set(F.text, new Y.Text());
      else {
        created.set(F.attrs, new Y.Map<unknown>());
        created.set(F.children, new Y.Array<string>());
      }
      record = created;
      ids.keyOf.set(node.id, key);
      ids.idOf.set(key, node.id);
      this.#partOfKey.set(key, partName);
    }
    const rec = record as Y.Map<unknown>;
    if (node.type === 'text') {
      const text = rec.get(F.text);
      const done =
        text instanceof Y.Text ? syncText(text, node.value, created, write.caseChange) : undefined;
      if (done) write.texts.set(key, done);
      return key;
    }
    const shell = shellOf(node);
    if (JSON.stringify(rec.get(F.shell)) !== JSON.stringify(shell)) rec.set(F.shell, shell);
    const view = this.#views.get(key);
    // A merged container: what this replica adds goes to the copy it wrote, when it has one,
    // so its own undo reverts it and a peer's undo does not.
    const client = this.#room.doc.clientID.toString(36);
    const ownCopy = view?.merge
      ? ([key, ...view.companions].find((k) => clientOf(k) === client) ?? key)
      : key;
    const attrs = rec.get(F.attrs);
    if (attrs instanceof Y.Map) {
      const desired = new Map<string, Record<string, string>>();
      for (const a of node.attrs)
        desired.set(`${a.ns}${ATTR_SEP}${a.local}`, {
          ns: a.ns,
          local: a.local,
          value: a.value,
          ...(a.prefix ? { prefix: a.prefix } : {}),
        });
      // A merged container: an attribute shown from a copy is written to the copy's record,
      // and a removed attribute leaves every copy, so it does not show again.
      // A hidden plain singleton keeps its content: only a merged container's copies change.
      const copies = view?.merge ? view.companions.map((c) => this.#attrsOf(c)) : [];
      for (const k of [...attrs.keys()]) if (!desired.has(k)) attrs.delete(k);
      for (const copy of copies)
        for (const k of [...(copy?.keys() ?? [])]) if (!desired.has(k)) copy!.delete(k);
      for (const [k, v] of desired) {
        const holder = view?.attrFrom.get(k) ?? (attrs.has(k) ? key : undefined);
        const held = holder === undefined ? undefined : this.#attrsOf(holder);
        if (held && JSON.stringify(held.get(k)) === JSON.stringify(v)) continue;
        // A new or changed value goes to the copy this replica wrote, and leaves the others,
        // so a peer's undo of its own copy does not take this edit with it.
        const target: Y.Map<unknown> = this.#attrsOf(ownCopy) ?? attrs;
        target.set(k, v);
        if (view?.merge)
          for (const other of [key, ...view.companions])
            if (other !== ownCopy) this.#attrsOf(other)?.delete(k);
      }
    }
    const childKeys = node.children.map((c) => this.#syncNode(c, write));
    const childKeySet = new Set(childKeys);
    const children = rec.get(F.children);
    // A shown child from a copy stays in the copy's list; a hidden copy stays after its winner;
    // a new child goes to this replica's own copy.
    // A companion stays only in the list that holds it, so no record gets a second parent, and
    // no key is listed twice.
    const withCopies = (list: readonly string[], here: ReadonlySet<string>) => {
      const out: string[] = [];
      const seen = new Set<string>();
      const push = (k: string) => {
        if (seen.has(k)) return;
        seen.add(k);
        out.push(k);
      };
      for (const k of list) {
        push(k);
        for (const c of this.#views.get(k)?.companions ?? []) if (here.has(c)) push(c);
      }
      return out;
    };
    const listedHere = new Set(children instanceof Y.Array ? children.toArray() : []);
    // A node that the read shows only for a peer's content after an undo is not listed where
    // it shows, and a run that the move rescue shows here is not listed here.
    const undone = this.#unlisted.get(partName);
    const rescuedIn = this.#rescuedIn.get(partName)?.inParent;
    const unlisted = {
      has: (k: string) =>
        !write.materialize.has(k) && (undone?.get(k) === key || rescuedIn?.get(k) === key),
    };
    const added = view
      ? childKeys.filter((k) => !view.childFrom.has(k) && !listedHere.has(k) && !unlisted?.has(k))
      : [];
    const own = (
      view
        ? childKeys.filter((k) => !view.childFrom.has(k) && (listedHere.has(k) || ownCopy === key))
        : childKeys
    ).filter((k) => !unlisted?.has(k) || listedHere.has(k));
    if (children instanceof Y.Array) {
      const list = children as Y.Array<string>;
      const current = list.toArray();
      syncList(
        list,
        withDropped(
          current,
          inRoomOrder(node, current, withCopies(own, new Set(current)), (k) => this.#nameOf(k)),
          this.#dropped.get(partName)
        )
      );
    }
    // A merged container: a slot this node still fills keeps every copy's child for it (a
    // shadowed one shows when its winner goes), and a removed slot leaves every copy.
    const filled = new Set<string | undefined>();
    if (view?.merge)
      for (const c of node.children)
        if (c.type === 'element') filled.add(groupOf(node, c, (l) => attrOfElement(c, l))?.slot);
    // A slot that this write fills with a new child replaces what every copy shadowed there.
    const addedSet = new Set(added);
    if (view?.merge)
      node.children.forEach((c, i) => {
        if (c.type === 'element' && addedSet.has(childKeys[i]!))
          filled.delete(groupOf(node, c, (l) => attrOfElement(c, l))?.slot);
      });
    for (const copy of view?.merge ? view.companions : []) {
      const list = this.#childrenOf(copy);
      if (!list) continue;
      const current = list.toArray();
      const keep = current.filter((k) => {
        // A hidden copy stays while its winner does; it leaves with a removed or replaced one.
        if (this.#hidden.has(k)) return childKeySet.has(this.#winnerOf.get(k)!);
        if (view!.childFrom.get(k) === copy) return childKeySet.has(k);
        const slot = view!.shadowed.get(k);
        return slot !== undefined && filled.has(slot);
      });
      syncList(list, withCopies([...keep, ...(copy === ownCopy ? added : [])], new Set(current)));
    }
    return key;
  }

  /** The element name of a record, or undefined for a text record or a malformed one. */
  #nameOf(key: string): { ns: string; local: string } | undefined {
    const rec = this.#room.nodes.get(key);
    const shell = rec instanceof Y.Map ? rec.get(F.shell) : undefined;
    if (typeof shell !== 'object' || shell === null) return undefined;
    const { ns, local } = shell as Record<string, unknown>;
    return typeof ns === 'string' && typeof local === 'string' ? { ns, local } : undefined;
  }

  /** The attribute map of a record, or undefined. */
  #attrsOf(key: string): Y.Map<unknown> | undefined {
    const rec = this.#room.nodes.get(key);
    const attrs = rec instanceof Y.Map ? rec.get(F.attrs) : undefined;
    return attrs instanceof Y.Map ? attrs : undefined;
  }

  /** The child list of a record, or undefined. */
  #childrenOf(key: string): Y.Array<string> | undefined {
    const rec = this.#room.nodes.get(key);
    const list = rec instanceof Y.Map ? rec.get(F.children) : undefined;
    return list instanceof Y.Array ? (list as Y.Array<string>) : undefined;
  }

  /** The shared key of one of this replica's store nodes, or undefined. */
  keyOf(partName: string, id: NodeId): string | undefined {
    return this.#ids.get(partName)?.keyOf.get(id);
  }

  /** This replica's store id for a shared key, or undefined. */
  idOf(partName: string, key: string): NodeId | undefined {
    return this.#ids.get(partName)?.idOf.get(key);
  }

  /** The `Y.Text` of a text node record, or undefined. */
  textOf(key: string): Y.Text | undefined {
    const rec = this.#room.nodes.get(key);
    const text = rec instanceof Y.Map ? rec.get(F.text) : undefined;
    return text instanceof Y.Text ? text : undefined;
  }

  /** Locate the visible copy of a character after a run split, move, or restoration. */
  locateCharacter(part: string, id: Y.ID): { node: string; offset: number; text: Y.Text } | null {
    return locateCharacter(this.#room.doc, id, this.#ids.get(part)?.idOf ?? [], (key) =>
      this.textOf(key)
    );
  }

  /** The part of a record, or undefined. */
  partOf(key: string): string | undefined {
    const rec = this.#room.nodes.get(key);
    const part = rec instanceof Y.Map ? rec.get(F.part) : undefined;
    return typeof part === 'string' ? part : undefined;
  }

  /**
   * The shared key and part of the record that holds `text`, or undefined. The record states
   * its own key; the claim is checked against the nodes map, because a peer wrote it.
   */
  recordOfText(text: Y.Text): { readonly key: string; readonly part: string } | undefined {
    const rec = text.parent;
    const key = rec instanceof Y.Map ? rec.get(F.key) : undefined;
    if (typeof key !== 'string' || this.#room.nodes.get(key) !== rec) return undefined;
    const part = (rec as Y.Map<unknown>).get(F.part);
    return typeof part === 'string' ? { key, part } : undefined;
  }

  /** The part names a remote room transaction touched. */
  touchedParts(transaction: Y.Transaction): Set<string> {
    return touchedParts(this.#room, transaction, this.#partOfKey);
  }

  /** Every part the room holds, plus the local ones it does not (to remove). */
  allParts(local: PptxPackage): Set<string> {
    const out = new Set<string>(this.#room.parts.keys());
    for (const name of local.parts.keys()) out.add(name);
    return out;
  }

  /**
   * Read `partNames` back from the room and rebuild them against `local`. All or nothing: one
   * invalid part refuses the whole rebuild, and the mapping does not change.
   */
  rebuild(partNames: Iterable<string>, local: PptxPackage, fresh = false): RebuildResult {
    const parts: PackagePart[] = [];
    const removed: string[] = [];
    const nextIds = new Map<string, PartIds>();
    const nextViews = new Map<string, Map<string, MergeView>>();
    const nextDropped = new Map<string, Set<string>>();
    const nextHidden = new Map<string, Map<string, SkippedListings>>();
    const nextMerged = new Map<string, Set<string>>();
    const names = [...partNames];
    // The room changed under these parts: notes of an older read must not delete a listing
    // that the room now shows, or list a text body as it showed it (`merged-cells.ts`), also
    // when a refused part stops the rebuild. A rebuild that succeeds notes them again.
    for (const name of names)
      [this.#hiddenListings, this.#mergedBodies].forEach((m) => m.delete(name));
    const nextUnlisted = new Map<string, ReadonlyMap<string, string>>();
    const nextRescuedIn = new Map<string, RescuedIn>();
    // Which records list each key, built once for the pass on first use (`removed-content.ts`).
    const listingsOfPart = roomListings(this.#room);
    for (const name of names) {
      const entry = partEntryOf(this.#room.parts.get(name));
      if (!entry) {
        if (this.#room.parts.has(name)) return { ok: false, reason: `invalid part entry ${name}` };
        if (local.parts.has(name)) removed.push(name);
        continue;
      }
      if (entry.kind === 'binary') {
        const bytes = this.#room.binaries.get(name);
        if (!(bytes instanceof Uint8Array) || bytes.length > ROOM_LIMITS.maxBinaryBytes)
          return { ok: false, reason: `invalid binary part ${name}` };
        parts.push({ kind: 'binary', partName: name, bytes });
        continue;
      }
      this.#pendingLists = pendingChildren(
        this.#room,
        name,
        local.parts.get(name),
        this.#ids.get(name)?.keyOf
      );
      this.#pendingPlaced = new Map();
      const views = new Map<string, MergeView>();
      this.#reading = new Set();
      this.#readRefusal = undefined;
      this.#lineage = partLineage(this.#room, name, entry.root);
      this.#hiddenNow = new Map();
      this.#undone = UndoneListings.of(this.#room, name, entry.root);
      // Nodes a peer moved into a subtree another peer removed without seeing them go back.
      const readRecord = (key: string, seen: Set<string>, depth: number) =>
        this.#readElement(key, name, seen, depth, views);
      let [cameBack, stranded] = [new Set<string>(), new Set<string>()];
      const removedTablePlaces = new Map<string, string>();
      const listings = listingsOfPart(name);
      const readPart = (strandedBefore: ReadonlySet<string> = new Set()) => {
        [cameBack, stranded] = [new Set(), new Set()];
        this.#rescue = newParagraphRescue(strandedBefore, listings);
        let read = this.#readElement(entry.root, name, new Set(), 0, views);
        if (read && this.#undone) read = this.#undone.withStranded(read, readRecord);
        const moved =
          read && rescueMoved(this.#room, name, read, readRecord, cameBack, stranded, listings);
        removedTablePlaces.clear();
        return (
          moved && withRemovedTableText(this.#room, moved, readRecord, removedTablePlaces, cameBack)
        );
      };
      let rescued = readPart();
      // Split pieces inside a node the rescue put back are not lost: read once more with them.
      const lineage = this.#lineage;
      if (rescued && (lineage.lost.size > 0 || lineage.unseenWinners.size > 0)) {
        const again = partLineage(this.#room, name, entry.root, specKeys(rescued));
        if (!sameLineage(again, lineage)) {
          this.#lineage = again;
          this.#reading = new Set();
          this.#hiddenNow = new Map();
          views.clear();
          rescued = readPart();
        }
      }
      // A node moved into a removed paragraph that has no other place left: the paragraph
      // comes back with it. One more read, so the paragraph shows before the rescue looks.
      if (rescued && stranded.size > 0) {
        this.#reading = new Set();
        this.#hiddenNow = new Map();
        views.clear();
        rescued = readPart(stranded);
      }
      if (!rescued)
        return { ok: false, reason: `${this.#readRefusal ?? 'invalid node records'} in ${name}` };
      const placed = withoutEmptiedRescues(
        this.#room,
        placeEditedPieces(this.#room, name, rescued, readRecord, this.#lineage.lost),
        this.#rescue!
      );
      // A record that a rescue or a placement read refused refuses the part, as the main read.
      if (this.#readRefusal) return { ok: false, reason: `${this.#readRefusal} in ${name}` };
      // Rescued text, and removed paragraphs that came back, show in the drawn cell over them.
      const moved = new Set<string>();
      const back = new Set([
        ...cameBack,
        ...this.#rescue!.rescued,
        ...(this.#undone?.rehomed ?? []),
      ]);
      const spec = withMergedText(this.#room, placed, back, moved, this.#reading);
      nextMerged.set(name, moved);
      const known = fresh ? undefined : this.#ids.get(name);
      const current = local.parts.get(name);
      const nextId = current?.kind === 'xml' ? current.xml.nextId : 1;
      const built = buildReplicaPart(name, spec, {
        idOf: (key) => known?.idOf.get(key),
        nextId,
      });
      if (!built.ok) return { ok: false, reason: `${name}: ${built.reason}` };
      const ids: PartIds = { keyOf: new Map(), idOf: new Map() };
      for (const [key, id] of built.ids) {
        ids.idOf.set(key, id);
        ids.keyOf.set(id, key);
      }
      nextIds.set(name, ids);
      nextViews.set(name, views);
      nextDropped.set(name, this.#reading);
      nextHidden.set(name, shownParentsOnly(this.#hiddenNow, spec));
      nextUnlisted.set(
        name,
        new Map([...(this.#undone?.shownIn ?? []), ...this.#pendingPlaced, ...removedTablePlaces])
      );
      nextRescuedIn.set(name, rescuedParents(spec, cameBack));
      parts.push({ kind: 'xml', partName: name, xml: built.part as XmlPart });
    }
    for (const [name, ids] of nextIds) this.#setIds(name, ids);
    for (const name of removed) this.#setIds(name, undefined);
    for (const [name, views] of nextViews) this.#setViews(name, views);
    for (const name of removed) this.#setViews(name, new Map());
    for (const [name, dropped] of nextDropped) this.#dropped.set(name, dropped);
    for (const name of removed) this.#dropped.delete(name);
    for (const [name, hidden] of nextHidden)
      if (hidden.size > 0) this.#hiddenListings.set(name, hidden);
    for (const name of removed) this.#hiddenListings.delete(name);
    for (const [name, shown] of nextUnlisted) this.#unlisted.set(name, shown);
    for (const [name, inParent] of nextRescuedIn) this.#rescuedIn.set(name, inParent);
    for (const name of removed) this.#rescuedIn.delete(name);
    for (const name of removed) this.#unlisted.delete(name);
    for (const [name, bodies] of nextMerged) this.#mergedBodies.set(name, bodies);
    this.#undone = null;
    return { ok: true, parts, removed };
  }

  /**
   * Why the last `#readElement` refused, when a limit (not a malformed record) refused it. A
   * limit anywhere in the part refuses the part, so the reason always names the real cause.
   */
  #readRefusal: string | undefined;
  #pendingLists = new Map<string, string[]>();
  #pendingPlaced = new Map<string, string>();

  #readElement(
    key: string,
    partName: string,
    seen: Set<string>,
    depth: number,
    views: Map<string, MergeView>
  ): ReplicaElementSpec | null {
    if (depth > ROOM_LIMITS.maxReadDepth || seen.has(key)) return null;
    seen.add(key);
    const rec = this.#room.nodes.get(key);
    if (!(rec instanceof Y.Map) || rec.get(F.part) !== partName) return null;
    const shell = rec.get(F.shell);
    const attrs = rec.get(F.attrs);
    const children = rec.get(F.children);
    if (typeof shell !== 'object' || shell === null) return null;
    if (!(attrs instanceof Y.Map) || !(children instanceof Y.Array)) return null;
    const s = shell as Record<string, unknown>;
    const kids: (ReplicaElementSpec | ReplicaTextSpec)[] = [];
    let childKeys: readonly unknown[] = children.toArray();
    const pending = this.#pendingLists.get(key);
    if (pending) {
      const kept = new Set(
        pending.filter((k) => {
          const record = this.#room.nodes.get(k);
          return record instanceof Y.Map && !isRemoved(record);
        })
      );
      const listed = new Set(childKeys);
      childKeys = withDropped(pending, childKeys as string[], kept);
      for (const k of childKeys)
        if (typeof k === 'string' && !listed.has(k)) this.#pendingPlaced.set(k, key);
    }
    for (const [index, childKey] of childKeys.entries()) {
      if (typeof childKey !== 'string') continue;
      // A key listed twice (two peers moved one node) is kept once, at its first place. The
      // next local write deletes the other listing first (`dropHiddenListings`).
      if (seen.has(childKey)) {
        let skipped = this.#hiddenNow.get(key);
        if (!skipped)
          this.#hiddenNow.set(key, (skipped = { length: childKeys.length, at: new Map() }));
        skipped.at.set(index, childKey);
        continue;
      }
      const child = this.#room.nodes.get(childKey);
      if (child instanceof Y.Map && child.get(F.text) instanceof Y.Text) {
        if (child.get(F.part) !== partName) continue;
        seen.add(childKey);
        kids.push({ key: childKey, text: readText(child.get(F.text) as Y.Text) });
        continue;
      }
      const el = this.#readElement(childKey, partName, seen, depth + 1, views);
      if (el) kids.push(el);
      // A limit refuses the whole part (peers are untrusted); a malformed record is skipped.
      else if (this.#readRefusal) return null;
    }
    // Attribute order does not change the canonical form; sort for one order on every peer.
    const attrList = [...attrs.entries()]
      .sort(([a], [b]) => (a < b ? -1 : a > b ? 1 : 0))
      .map(([, v]) => v as { ns: string; local: string; prefix?: string; value: string });
    const ns = s.ns as string;
    const local = s.local as string;
    // A run or a paragraph that a peer removed while another typed into it comes back.
    const withRuns =
      (local === 'p' && ns === NS_A) || isTextBody({ ns, local })
        ? withRemovedRuns(
            this.#room,
            partName,
            rec as Y.Map<unknown>,
            kids,
            (k) => {
              const run = this.#readElement(k, partName, new Set(seen), depth + 1, views);
              return this.#readRefusal ? null : run;
            },
            this.#rescue,
            (k) => this.#undone?.showsBelow(k) ?? false,
            (spec) => {
              for (const key of specKeys(spec)) seen.add(key);
            }
          )
        : kids;
    // A node an undo unlisted shows again with what a peer wrote into it (`undo-unlisted.ts`).
    const listed = this.#undone
      ? this.#undone.withUndone(key, partName, withRuns, seen, (k) =>
          this.#readElement(k, partName, seen, depth + 1, views)
        )
      : withRuns;
    if (this.#readRefusal) return null;
    // Concurrent splits of one run keep one peer's pieces (`text-lineage.ts`).
    const kept = dropLosingSplits(
      this.#room,
      { ns, local },
      listed,
      this.#reading,
      this.#lineage.unseenWinners
    );
    const shown = withCompanions({ ns, local }, kept, views);
    if (!shown) {
      this.#readRefusal = 'too many concurrent copies of one element';
      return null;
    }
    return {
      key,
      ns,
      local,
      ...(typeof s.prefix === 'string' ? { prefix: s.prefix } : {}),
      nsDecls: Array.isArray(s.nsDecls) ? (s.nsDecls as { prefix: string; uri: string }[]) : [],
      attrs: attrList,
      // Concurrent inserts of two singletons at one place come back in schema order.
      children: [...inSchemaOrder({ ns, local }, shown, (k) => ('text' in k ? undefined : k))],
    };
  }
}
