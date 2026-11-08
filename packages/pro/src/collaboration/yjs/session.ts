/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// A collaboration session over a Y.Doc, after DOCX Editor's
// `packages/pro/src/collaboration/document-session.ts`. DOCX Editor commit adf0e178.
//
// The host owns the Y.Doc and its provider (Hocuspocus, y-webrtc, or a test link) and passes
// the document in, with an optional Yjs awareness for presence. On attach, an empty room is
// seeded from the opened deck; a room with content replaces the opened deck (the room is the
// truth). After that, local commits are written into the room, and remote room transactions
// rebuild the parts they touched. Yjs orders concurrent changes, so no server sequencer is
// needed.
//
// History is collaboration-owned, as in DOCX Editor: a `Y.UndoManager` tracks only this
// replica's own writes, so undo reverts this author's change and keeps a peer's concurrent
// edit. The undo itself is a room transaction, and it reaches the store through the same
// rebuild path as a peer's change. Each history step remembers the selection published when
// it was made, and `historySelection()` returns it.

import * as Y from 'yjs';
import {
  createCollaborationStatusTracker,
  type CollaborationDocumentPort,
  type CollaborationFailureCode,
  type CollaborationIdentity,
  type CollaborationLocalSelection,
  type CollaborationParticipant,
  type CollaborationRemoteSelection,
  type CollaborationSelectionMove,
  type CollaborationStatus,
  type CollaborationStatusSnapshot,
  type EditorCollaborationSession,
} from '@pptx-editor.dev/core/editor';
import type { PptxPackage } from '@pptx-editor.dev/core/store';
import { CollaborationSchemaError } from '../errors.ts';
import type { StepRanges } from './yjs-items.ts';
import { PackageMirror } from './mirror.ts';
import {
  AWARENESS_FIELD,
  decodeSelection,
  encodeSelection,
  identityError,
  localRecord,
  readParticipants,
  readRemoteSelections,
  type AwarenessLike,
  type EncodedSelection,
} from './presence.ts';
import {
  F,
  LOCAL_ORIGIN,
  REDO_REPAIR_ORIGIN,
  REPAIR_ORIGIN,
  SCHEMA_VERSION,
  SEED_ORIGIN,
  TEXT_REPAIR_ORIGIN,
  TIDY_ORIGIN,
  UNDONE_ORIGIN,
  roomOf,
} from './schema.ts';
import { RESTORED_KEY, followsIn } from './text-identity.ts';
import { repairCopies, repairFollows } from './text-repair.ts';
import { nodeRecordDeleteFilter } from './undo-delete-filter.ts';
import {
  carryFollowed,
  checkFollowed,
  type FollowedRun,
  tieFollowedRuns,
} from './undo-followed.ts';
import {
  dropRestoredTwins,
  noteRestored,
  repairRestored,
  restoredTextsOf,
  textsOfMarks,
} from './undo-restore.ts';
import { withoutLoserRestores, withoutTakenRedo } from './undo-splits.ts';
import { forgetUndoneListings, noteUndoneListings, unlistEmptied } from './undo-unlisted.ts';
import {
  attributeWrites,
  lostReferences,
  missingAttributes,
  unredoneAttributes,
  type AttributeWrite,
} from './undo-repair.ts';

/** Options of `createYjsCollaborationSession`. @public */
export interface YjsCollaborationSessionOptions {
  /** The shared document. The host connects it to a provider. */
  readonly doc: Y.Doc;
  /** The room's document identity. */
  readonly documentId: string;
  /** The local author. */
  readonly identity: CollaborationIdentity;
  /**
   * Yjs awareness for presence (`new Awareness(doc)` from `y-protocols/awareness`), usually
   * the provider's. Absent, the session lists only the local participant.
   */
  readonly awareness?: AwarenessLike;
}

/** A session over a Y.Doc. @public */
export interface YjsCollaborationSession extends EditorCollaborationSession {
  /**
   * Report the provider's connection. `false` sets `disconnected` with
   * `transport-disconnected`; edits continue and merge on reconnect. `true` returns to `ready`.
   * Call it from the provider's status event.
   */
  setConnected(connected: boolean): void;
}

let sessionCounter = 0;

/** Call every listener; one that throws does not stop the others or reach the caller. */
function notify<A extends unknown[]>(listeners: Set<(...args: A) => void>, ...args: A): void {
  for (const listener of [...listeners]) {
    try {
      listener(...args);
    } catch (error) {
      console.warn('[pptx-editor] collaboration listener failed', error);
    }
  }
}

/**
 * Create a session over a Y.Doc. Register it with `collaborationModule({ session })`. Throws
 * `CollaborationSchemaError` for an invalid document id or identity.
 *
 * @public
 */
export function createYjsCollaborationSession(
  options: YjsCollaborationSessionOptions
): YjsCollaborationSession {
  const { doc, documentId, identity, awareness } = options;
  if (typeof documentId !== 'string' || documentId.trim() === '' || documentId.length > 256)
    throw new CollaborationSchemaError('invalid-document-id');
  const bad = identityError(identity);
  if (bad) throw new CollaborationSchemaError(bad);

  const room = roomOf(doc);
  const tracker = createCollaborationStatusTracker('initializing');
  type StatusListener = (
    status: CollaborationStatus,
    reason?: CollaborationFailureCode,
    detail?: string
  ) => void;
  const statusListeners = new Set<StatusListener>();
  const participantListeners = new Set<(p: readonly CollaborationParticipant[]) => void>();
  const selectionListeners = new Set<(s: readonly CollaborationRemoteSelection[]) => void>();
  let detachCurrent: (() => void) | null = null;
  let mirror: PackageMirror | null = null;
  // One mint counter for every attach of this session, so keys never repeat across attaches.
  const mintCounter = { value: 0 };
  let history: Y.UndoManager | null = null;
  /** The port of the attached editor, for the package before and after an undo. */
  let attachedPort: CollaborationDocumentPort | null = null;
  /** The item the last undo or redo applied. */
  let popped: { meta: Map<string, unknown> } | null = null;
  let destroyed = false;
  let connected = true;
  let participants: readonly CollaborationParticipant[] = [{ ...identity, isLocal: true }];
  let remote: readonly CollaborationRemoteSelection[] = [];
  // This replica's published selection, as the editor gave it and as it travels.
  let local: { selection: CollaborationLocalSelection; encoded: EncodedSelection } | null = null;
  let move: CollaborationSelectionMove | null = null;
  let historyEncoded: EncodedSelection | null = null;
  // Set by a rebuild that failed inside an undo or redo, so that call reports false.
  let applyFailed = false;
  sessionCounter += 1;
  const sessionId = `${doc.clientID.toString(36)}-${sessionCounter}`;

  const setStatus = (
    next: CollaborationStatus,
    code?: CollaborationFailureCode,
    detail?: string,
    recovering = false
  ): void => {
    if (!tracker.set(next, code, detail, recovering)) return;
    const reason = tracker.snapshot().reason;
    notify(statusListeners, next, reason?.code, reason?.detail);
  };
  // The status a healthy replica holds: `ready`, or `disconnected` while the provider is down.
  const settled = (): void =>
    connected ? setStatus('ready') : setStatus('disconnected', 'transport-disconnected');

  const publishSelection = (encoded: EncodedSelection | undefined): void => {
    if (awareness && !destroyed)
      awareness.setLocalStateField(AWARENESS_FIELD, localRecord(identity, encoded));
  };

  // Presence: recompute on awareness changes and after each rebuild (ids can change).
  const refreshPresence = (): void => {
    if (!awareness) return;
    const nextParticipants = readParticipants(awareness);
    if (JSON.stringify(nextParticipants) !== JSON.stringify(participants)) {
      participants = nextParticipants;
      notify(participantListeners, participants);
    }
    const nextRemote = mirror ? readRemoteSelections(awareness, doc, mirror) : [];
    if (JSON.stringify(nextRemote) !== JSON.stringify(remote)) {
      remote = nextRemote;
      notify(selectionListeners, remote);
    }
  };
  publishSelection(undefined);
  awareness?.on('change', refreshPresence);
  refreshPresence();

  /**
   * Write a repair as part of the undo or redo that just ran: Yjs tracks it while it still
   * counts as undoing or redoing (it lands on the other stack). With `merge`, the item it adds
   * merges into the item the step pushed, so one redo or undo reverts both; without, it stays
   * its own item (a step Yjs could not redo). Returns whether the write added an item.
   */
  const stepRepair = (
    manager: Y.UndoManager,
    undo: boolean,
    write: () => void,
    merge = true
  ): boolean => {
    const stack = undo ? manager.redoStack : manager.undoStack;
    const before = stack.length;
    if (undo) manager.undoing = true;
    else manager.redoing = true;
    try {
      doc.transact(write, REDO_REPAIR_ORIGIN);
    } finally {
      manager.undoing = false;
      manager.redoing = false;
    }
    if (stack.length !== before + 1) return false;
    const item = stack[stack.length - 2];
    if (!merge || !item) return true;
    const added = stack.pop()!;
    item.deletions = Y.mergeDeleteSets([item.deletions, added.deletions]);
    item.insertions = Y.mergeDeleteSets([item.insertions, added.insertions]);
    return true;
  };

  /** Write a redo's values over an undo repair, as part of the redo (`stepRepair`). */
  const redoRepair = (
    manager: Y.UndoManager,
    values: [Y.Map<unknown>, string, unknown][],
    merge: boolean
  ): boolean =>
    values.length > 0 &&
    stepRepair(
      manager,
      false,
      () => {
        for (const [map, key, value] of values) map.set(key, value);
      },
      merge
    );

  const step = (manager: Y.UndoManager | null, undo: boolean): boolean => {
    const stack = undo ? manager?.undoStack : manager?.redoStack;
    if (!manager || !stack || stack.length === 0) return false;
    applyFailed = false;
    const before = attachedPort?.package();
    // The step's own record of its writes. Yjs gives the item that undo or redo pushes a new
    // meta, so the item it pushes takes the record of its own step (`currStackItem`).
    // Yjs skips each item it cannot apply and takes the next one in the same call, so the item
    // it applied (`popped`) can differ from the one on top before the call, and none may apply.
    const top = stack[stack.length - 1]!;
    // What the step must not take from a peer or put back twice (`undo-unlisted.ts`,
    // `undo-splits.ts`), set on the step before Yjs applies it.
    if (undo) {
      // Yjs can skip the top step and apply an older one in the same call: no step that it
      // can apply puts back the old entry of a run that a peer moved on (`undo-followed.ts`).
      for (const item of stack.slice(0, -1)) checkFollowed(item);
      noteUndoneListings(
        room,
        top,
        (write) => doc.transact(write, UNDONE_ORIGIN),
        checkFollowed(top)
      );
      top.deletions = withoutLoserRestores(room, top);
    } else {
      const kept = withoutTakenRedo(room, top);
      top.insertions = kept.insertions;
      top.deletions = kept.deletions;
    }
    const listed = [...stack];
    popped = null;
    if (undo) manager.undo();
    else manager.redo();
    if (!undo && popped === top)
      forgetUndoneListings(room, top, (write) => doc.transact(write, UNDONE_ORIGIN));
    const applied = popped as { meta: Map<string, unknown> } | null;
    const left = new Set(stack);
    // The items the call took off the stack without applying, in the order it took them.
    const skipped = listed.reverse().filter((item) => !left.has(item) && item !== applied);
    const writesOf = (item: { meta: Map<string, unknown> }) =>
      (item.meta.get('attrs') as AttributeWrite[] | undefined) ?? [];
    let repaired = false;
    let skippedTop = false;
    // The step that reverts the one Yjs applied in place of the top step: what it puts back
    // gets restore marks too (`noteRestored`), or those characters lose their identity.
    let compensated: { meta: Map<string, unknown> } | null = null;
    if (undo && applied && applied !== top) {
      // Yjs could not undo the top step (a peer took its values) and undid an older one in the
      // same call. Undo stays one step: the older step comes back, and the call reports that
      // it changed nothing, so the next undo takes the older step.
      popped = null;
      manager.redo();
      compensated = popped;
      skippedTop = true;
    } else if (!undo && applied === top) {
      repaired = redoRepair(manager, unredoneAttributes(writesOf(top)), true);
    } else if (!undo) {
      // Yjs could not redo the top step (a repair wrote over its values) and went on to later
      // steps. Redo stays one step: the step Yjs applied goes back with an undo, the other
      // skipped steps go back onto the redo stack, and the top step comes back by its repair.
      if (applied) {
        popped = null;
        manager.undo();
        compensated = popped;
      }
      for (const item of skipped.slice(1).reverse()) manager.redoStack.push(item);
      repaired = redoRepair(manager, unredoneAttributes(writesOf(top)), false);
    }
    // After the step, in one transaction, so a peer never sees its marks without its repair:
    // each character it put back stands for the character it restored (`noteRestored`; the
    // marks map is outside the undo scope, so no history tracks them); the characters a peer's
    // copy holds go again, and the others follow their moved neighbours (`undo-restore.ts`);
    // and a node an undone step made that a peer listed again leaves when it holds nothing
    // (`unlistEmptied`). The repairs are part of the step.
    if (compensated) {
      const ranges = compensated as unknown as StepRanges;
      doc.transact(() => noteRestored(room, ranges.deletions), UNDONE_ORIGIN);
    }
    const done = applied && !skippedTop && (undo || applied === top) ? applied : null;
    if (done && mirror) {
      const reading = mirror;
      const ranges = done as unknown as StepRanges;
      const shows = (part: string, key: string) => reading.idOf(part, key) !== undefined;
      let roots: ReadonlySet<string> = new Set();
      stepRepair(manager, undo, () => {
        const restored = noteRestored(room, ranges.deletions);
        roots = restored.roots;
        if (undo) unlistEmptied(room, ranges, shows);
        repairRestored(room, ranges.deletions, shows);
        // A peer that put back the same characters before shows them once.
        dropRestoredTwins(room, restored.texts, shows);
      });
      // This replica's delete of a character the step put back no longer reaches its copies.
      reading.forgetDeleted(roots);
    }
    // What the step took that a peer's write or content still needs comes back, read
    // after the redo bookkeeping above, so the repair writes over the current package.
    const after = attachedPort?.package();
    const values = undo && applied && !skippedTop ? missingAttributes(writesOf(applied)) : [];
    const lost = before && after && !skippedTop ? lostReferences(before, after) : [];
    if ((values.length > 0 || lost.length > 0) && after && mirror) {
      const repairing = mirror;
      doc.transact(() => {
        for (const [map, key, value] of values) map.set(key, value);
        if (lost.length === 0) return;
        const parts = new Map(after.parts);
        for (const p of lost) parts.set(p.partName, p);
        repairing.writeChange(
          after,
          { parts },
          lost.map((p) => p.partName)
        );
      }, REPAIR_ORIGIN);
    }
    // A refused rebuild leaves the store behind the room, and a step that Yjs could not apply
    // changed nothing: report both, never claim success.
    return !applyFailed && (undo ? applied !== null && !skippedTop : applied === top || repaired);
  };

  const session: YjsCollaborationSession = {
    documentId,
    sessionId,
    identity,
    get attached() {
      return detachCurrent !== null;
    },
    status: () => tracker.status(),
    statusSnapshot: (): CollaborationStatusSnapshot => tracker.snapshot(),
    subscribeStatus(listener) {
      statusListeners.add(listener);
      return () => statusListeners.delete(listener);
    },
    setConnected(next) {
      if (destroyed || next === connected) return;
      connected = next;
      const snapshot = tracker.snapshot();
      if (snapshot.status === 'ready' || snapshot.status === 'disconnected') settled();
    },
    participants: () => participants,
    subscribeParticipants(listener) {
      participantListeners.add(listener);
      return () => participantListeners.delete(listener);
    },
    remoteSelections: () => remote,
    subscribeRemoteSelections(listener) {
      selectionListeners.add(listener);
      return () => selectionListeners.delete(listener);
    },
    setLocalSelection(selection: CollaborationLocalSelection | null) {
      const encoded = selection && mirror ? encodeSelection(selection, mirror) : undefined;
      local = selection && encoded ? { selection, encoded } : null;
      publishSelection(encoded);
    },
    remoteSelectionMove: () => move,
    historySelection: () =>
      historyEncoded && mirror ? decodeSelection(historyEncoded, doc, mirror) : null,
    canUndo: () => (history ? history.undoStack.length > 0 : false),
    canRedo: () => (history ? history.redoStack.length > 0 : false),
    undo: () => step(history, true),
    redo: () => step(history, false),
    attach(port: CollaborationDocumentPort) {
      if (destroyed) return () => {};
      detachCurrent?.();
      setStatus('initializing');
      const schema = room.meta.get('schema');
      if (schema !== undefined && schema !== SCHEMA_VERSION) {
        setStatus('error', 'schema-version-mismatch', `room schema ${String(schema)}`);
        return () => {};
      }
      const attachedMirror = new PackageMirror(room, mintCounter);
      attachedPort = port;
      mirror = attachedMirror;

      const applyRemote = (names: Iterable<string>, fresh: boolean, peer = false): void => {
        const before = port.package();
        const rebuilt = attachedMirror.rebuild(names, before, fresh);
        if (!rebuilt.ok) {
          applyFailed = true;
          // A later update can complete the room state, so this heals by itself.
          return setStatus('error', 'remote-apply-failed', rebuilt.reason, true);
        }
        // Where the published selection lands after the change, before the store publishes
        // it, so the editor can follow it while it handles the change.
        const to = local ? decodeSelection(local.encoded, doc, attachedMirror) : null;
        move =
          local && to && JSON.stringify(to) !== JSON.stringify(local.selection)
            ? { from: local.selection, to }
            : null;
        if (move && local) local = { selection: move.to, encoded: local.encoded };
        if (rebuilt.parts.length > 0 || rebuilt.removed.length > 0) {
          const applied = port.applyRemoteParts(rebuilt.parts, rebuilt.removed, 'collaboration');
          move = null;
          if (!applied.ok) {
            applyFailed = true;
            const code = applied.rejection.code;
            return code === 'unsafe-part-name'
              ? setStatus('error', 'unsafe-part-name', code)
              : setStatus('error', 'remote-apply-failed', code, true);
          }
        }
        move = null;
        if (peer) repairAfterPeer(before, port.package());
        const snapshot = tracker.snapshot();
        if (snapshot.status === 'error' && snapshot.recovering) settled();
        refreshPresence();
      };

      if (attachedMirror.seeded) applyRemote(attachedMirror.allParts(port.package()), true);
      else
        doc.transact(() => {
          room.meta.set('schema', SCHEMA_VERSION);
          room.meta.set('documentId', documentId);
          attachedMirror.seed(port.package());
        }, SEED_ORIGIN);

      // Each store history entry is one undo step: a merged edit (typing a word) extends the
      // current step, any other edit starts a new one.
      const undoManager = new Y.UndoManager([room.parts, room.binaries, room.nodes], {
        trackedOrigins: new Set([LOCAL_ORIGIN, REDO_REPAIR_ORIGIN]),
        captureTimeout: Number.POSITIVE_INFINITY,
        // An undo keeps a node's record, so what a peer wrote into the node is never deleted.
        deleteFilter: nodeRecordDeleteFilter(room.nodes),
      });
      // A step remembers the selection published when it was made; undo and redo return there.
      type StackEvent = {
        stackItem: { meta: Map<string, unknown> };
        changedParentTypes?: ReadonlyMap<unknown, readonly unknown[]>;
      };
      // Each step keeps the attribute values from before it, for `missingAttributes`.
      const isAttrs = (map: Y.Map<unknown>): boolean => {
        const rec = map.parent;
        return rec instanceof Y.Map && rec.get(F.attrs) === map;
      };
      const noteWrites = (event: StackEvent): void => {
        // An item that undo or redo pushes gets the record of the step it reverts.
        const current = undoManager.currStackItem as { meta: Map<string, unknown> } | null;
        if ((undoManager.undoing || undoManager.redoing) && current) {
          event.stackItem.meta.set('attrs', current.meta.get('attrs') ?? []);
          carryFollowed(current, event.stackItem);
          return;
        }
        const writes = (event.stackItem.meta.get('attrs') as AttributeWrite[] | undefined) ?? [];
        // `changedParentTypes` lists an event under each changed ancestor: read each event once.
        const events = new Set<Y.YEvent<Y.AbstractType<unknown>>>();
        for (const list of event.changedParentTypes?.values() ?? [])
          for (const e of list as Y.YEvent<Y.AbstractType<unknown>>[]) events.add(e);
        attributeWrites(events, isAttrs, writes);
        event.stackItem.meta.set('attrs', writes);
      };
      undoManager.on('stack-item-added', (event: StackEvent) => {
        if (local) event.stackItem.meta.set('selection', local.encoded);
        noteWrites(event);
      });
      undoManager.on('stack-item-updated', noteWrites);
      undoManager.on('stack-item-popped', (event: StackEvent) => {
        historyEncoded = (event.stackItem.meta.get('selection') as EncodedSelection) ?? null;
        popped = event.stackItem;
      });
      history = undoManager;
      const offLocal = port.subscribeLocal((tx) => {
        if (!tx.merged) undoManager.stopCapturing();
        // Listings no read shows go first, in their own step that no undo reverts (mirror.ts).
        if (attachedMirror.hasHiddenListings(tx.parts))
          doc.transact(
            () => attachedMirror.dropHiddenListings(tx.parts, tx.before, tx.after),
            TIDY_ORIGIN
          );
        doc.transact(() => attachedMirror.writeChange(tx.before, tx.after, tx.parts), LOCAL_ORIGIN);
        // A node shown only for a peer's content leaves when this edit took that content.
        const again = attachedMirror.showingUnlisted(tx.parts);
        if (again.length > 0) applyRemote(again, false);
        // A local edit moves peers' text positions and can mint new ids: map them again.
        refreshPresence();
      });
      // A peer's change can take what this replica's content still uses: a peer's undo removes
      // a picture's relationship and media part while this replica copied the picture, or the
      // `Default` that this replica's new picture relied on. The peer could not see that use, so
      // the replica that holds it writes the repair. Two replicas that repair at once
      // write the same entries, and the "OPC entries" rule shows each one once.
      const repairAfterPeer = (before: PptxPackage, after: PptxPackage): void => {
        const lost = lostReferences(before, after);
        if (lost.length === 0) return;
        const parts = new Map(after.parts);
        for (const p of lost) parts.set(p.partName, p);
        doc.transact(
          () =>
            attachedMirror.writeChange(
              after,
              { parts },
              lost.map((p) => p.partName)
            ),
          REPAIR_ORIGIN
        );
      };
      // A peer's split that moved text this replica edited at the same time: the replica that
      // typed or deleted that text writes the repair (`text-repair.ts`). Text it moves keeps its
      // place in this replica's history, so its undo still removes it.
      const repairText = (transaction: Y.Transaction, fresh: Map<string, string[]>): void => {
        const deleted = attachedMirror.deletedText;
        const follows = deleted.size > 0 ? followsIn(room, transaction) : new Set<string>();
        if (fresh.size === 0 && follows.size === 0) return;
        const view = {
          keyOf: attachedMirror.keyOf.bind(attachedMirror),
          package: port.package(),
        };
        const moved: [Y.ID, Y.ID, number][] = [];
        const runs: FollowedRun[] = [];
        doc.transact(() => {
          for (const [part, keys] of fresh)
            repairCopies(
              room,
              part,
              keys,
              deleted,
              view,
              (from, to, length) => moved.push([from, to, length]),
              (run) => runs.push(run)
            );
          repairFollows(room, follows, deleted);
        }, TEXT_REPAIR_ORIGIN);
        // A peer's run that followed this replica's Enter goes back with an undo of the Enter.
        tieFollowedRuns([...undoManager.undoStack, ...undoManager.redoStack], runs);
        for (const [from, to, length] of moved)
          for (const item of [...undoManager.undoStack, ...undoManager.redoStack]) {
            if (!Y.isDeleted(item.insertions, from)) continue;
            const added = Y.createDeleteSet();
            added.clients.set(to.client, [{ clock: to.clock, len: length }] as never);
            item.insertions = Y.mergeDeleteSets([item.insertions, added]);
          }
      };
      const restoredMarks = doc.getMap(RESTORED_KEY) as unknown as Y.AbstractType<
        Y.YEvent<Y.Map<unknown>>
      >;
      const onTransaction = (transaction: Y.Transaction): void => {
        // A room format that arrives after attach (a late provider sync, a peer that seeded
        // the room at the same time) refuses the room as attach does: this replica stops
        // applying the room and keeps its own presentation.
        const schema = room.meta.get('schema');
        if (schema !== SCHEMA_VERSION && transaction.changed.has(room.meta as never)) {
          detach();
          setStatus('error', 'schema-version-mismatch', `room schema ${String(schema)}`);
          return;
        }
        const origin: unknown = transaction.origin;
        // This replica's own undone listings change nothing until its undo, which rebuilds.
        if (
          origin === LOCAL_ORIGIN ||
          origin === SEED_ORIGIN ||
          origin === UNDONE_ORIGIN ||
          origin === TIDY_ORIGIN
        )
          return;
        const touched = attachedMirror.touchedParts(transaction);
        // Only a peer's change needs the repair here: this replica's own undo and redo repair in
        // `step`, and its own repair writes nothing that a repair would restore.
        if (touched.size > 0) applyRemote(touched, false, !transaction.local);
        if (transaction.local) return;
        const fresh = attachedMirror.newCopies(transaction);
        repairText(transaction, fresh);
        // A character put back twice, by two peers' undos or beside a peer's copy of it, shows
        // once (`dropRestoredTwins`): when the marks or the copy arrive.
        const marks = transaction.changed.get(restoredMarks);
        const texts = marks ? textsOfMarks(room, marks) : new Set<Y.Text>();
        if (fresh.size > 0 && room.doc.getMap(RESTORED_KEY).size > 0)
          for (const t of restoredTextsOf(room, new Set(fresh.keys()))) texts.add(t);
        if (texts.size > 0) {
          const shows = (part: string, key: string) => attachedMirror.idOf(part, key) !== undefined;
          doc.transact(() => dropRestoredTwins(room, texts, shows), TEXT_REPAIR_ORIGIN);
        }
      };
      doc.on('afterTransaction', onTransaction);
      if (tracker.status() === 'initializing') settled();
      refreshPresence();

      const detach = (): void => {
        offLocal();
        doc.off('afterTransaction', onTransaction);
        undoManager.destroy();
        if (history === undoManager) history = null;
        if (mirror === attachedMirror) mirror = null;
        if (attachedPort === port) attachedPort = null;
        // A detached editor shows no caret: peers must not see a stale one.
        local = null;
        historyEncoded = null;
        publishSelection(undefined);
        refreshPresence();
        if (detachCurrent === detach) detachCurrent = null;
      };
      detachCurrent = detach;
      return detach;
    },
    destroy() {
      if (destroyed) return;
      detachCurrent?.();
      awareness?.off('change', refreshPresence);
      awareness?.setLocalStateField(AWARENESS_FIELD, null);
      destroyed = true;
      setStatus('destroyed', 'collaboration-session-destroyed');
      statusListeners.clear();
      participantListeners.clear();
      selectionListeners.clear();
    },
  };
  return session;
}
