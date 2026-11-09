// The `EditorModule` seam, after DOCX Editor's `packages/core/src/contracts/modules.ts`: how an
// optional capability (the review module and the collaboration module in `@pptx-editor.dev/pro`)
// contributes to the engine without the free tier importing it. DOCX Editor commit adf0e178.
//
// NOT a plugin system. The shape is closed: every contribution point is named here, the editor
// reads the registry at its own dispatch points, and core never imports a capability package.
// The free engine behaves the same with an empty registry: decks round-trip, comments are kept,
// and no replica attaches. It lives in the `editor` lane, not `contracts`, because it names
// store types (`PptxOp`, `PptxPackage`) and `contracts` imports no lane.

import type {
  NodeId,
  OpRejection,
  PackagePart,
  PptxDocumentStore,
  PptxOp,
  PptxPackage,
  StoreChange,
} from '../store/index.ts';
import type { ReviewPaneOptions } from '../contracts/index.ts';
import { bindReplicaIds } from './replica-ids.ts';

// --- Collaboration --------------------------------------------------------------------------

/**
 * The state of a collaboration replica. The same values as DOCX Editor's `CollaborationStatus`.
 *
 * @public
 */
export type CollaborationStatus = 'initializing' | 'ready' | 'disconnected' | 'error' | 'destroyed';

/**
 * Why a replica refused work, left `ready`, or failed a check. The subset of DOCX Editor's
 * `CollaborationFailureCode` that PPTX Editor emits, with the same meanings. Branch on the code;
 * log the detail.
 *
 * @public
 */
export type CollaborationFailureCode =
  /** The session was destroyed. Terminal. */
  | 'collaboration-session-destroyed'
  /** Factory input rejected before joining. Fix the value the host passed. */
  | 'invalid-document-id'
  | 'invalid-identity'
  | 'invalid-identity-color'
  /** This client and the room have different formats. Terminal. Save local work. */
  | 'schema-version-mismatch'
  /** A remote update is not applied yet. Heals when a later update arrives. Wait. */
  | 'remote-apply-failed'
  /** The store refused a rebuilt part name. Terminal for this room. */
  | 'unsafe-part-name'
  /** The host reported that the provider lost its connection. Recovers on its own. */
  | 'transport-disconnected'
  /** The Y.Doc holds no room yet. Wait for the provider to sync, or seed the room. */
  | 'not-initialized';

/** One collaboration failure: a typed code plus optional free-form detail. @public */
export interface CollaborationFailure {
  /** Stable code for application logic. */
  readonly code: CollaborationFailureCode;
  /** Bounded diagnostic text for logs. Not a parsing contract. */
  readonly detail?: string;
}

/** Cached status read, as in DOCX Editor. Same reference until one of its fields changes. @public */
export interface CollaborationStatusSnapshot {
  /** Current lifecycle state. */
  readonly status: CollaborationStatus;
  /** Why the replica holds `status`, or `undefined` when nothing is wrong. */
  readonly reason: CollaborationFailure | undefined;
  /** The most recent `error` reason. Kept after the replica recovers. */
  readonly lastFailure: CollaborationFailure | undefined;
  /** How many failures this session has recorded. Absent from a session that does not count. */
  readonly failureCount?: number;
  /** In `error`, but the session heals by itself on the next clean update. Absent means false. */
  readonly recovering?: boolean;
}

/** Human or automation identity attached to authored collaboration transactions. @public */
export interface CollaborationIdentity {
  /** Stable author id. Recorded on authored changes. */
  readonly actorId: string;
  /** Display name. */
  readonly name: string;
  /** Presence color, a CSS color string. */
  readonly color?: string;
  /** Whether a person or an automation made the change. */
  readonly role?: 'human' | 'agent';
}

/** One participant in a collaboration room. @public */
export interface CollaborationParticipant extends CollaborationIdentity {
  /** Whether this participant is the local editor. */
  readonly isLocal: boolean;
}

/**
 * A selection as the local editor reports it to a session. Ids are this replica's store node
 * ids; the session sends shared keys, never these ids. A text point is an offset (UTF-16) into
 * a text node, or offset 0 of a paragraph element that holds no text yet.
 *
 * @public
 */
export type CollaborationLocalSelection =
  | {
      readonly kind: 'shapes';
      /** The slide part. */
      readonly part: string;
      /** Selected shape element ids. */
      readonly shapes: readonly NodeId[];
    }
  | {
      readonly kind: 'text';
      /** The slide part. */
      readonly part: string;
      /** Where the selection starts: a text node id and an offset. */
      readonly anchor: { readonly node: NodeId; readonly offset: number };
      /** Where it ends. Equal to `anchor` for a caret. */
      readonly head: { readonly node: NodeId; readonly offset: number };
    };

/**
 * How a remote change moved this replica's published selection: `to` is `from` re-read
 * through the room's relative positions, after DOCX Editor's `CollaborationSelectionMove`.
 *
 * @public
 */
export interface CollaborationSelectionMove {
  /** The selection as it was published, on the ids before the change. */
  readonly from: CollaborationLocalSelection;
  /** The same selection after the change, on the current ids. */
  readonly to: CollaborationLocalSelection;
}

/** A peer's selection, mapped onto this replica's node ids. @public */
export interface CollaborationRemoteSelection {
  /** The peer's author id. */
  readonly actorId: string;
  /** The peer's display name. */
  readonly name: string;
  /** The peer's presence color. */
  readonly color?: string;
  /** What the peer selected. */
  readonly selection: CollaborationLocalSelection;
}

/**
 * One committed local change, as the replica sends it to its peers: an edit, an undo, or a
 * redo. A state-based replica (the pro Yjs binding) diffs `before` and `after` over `parts`; an
 * op-based one sends `ops`, which only edits carry.
 *
 * @public
 */
export interface CollaborationLocalTransaction {
  /** What made the change. */
  readonly origin: 'user' | 'undo' | 'redo';
  /** The package before the change. */
  readonly before: PptxPackage;
  /** The package after the change. */
  readonly after: PptxPackage;
  /** Part names whose content changed. */
  readonly parts: readonly string[];
  /** The store revision before the commit. */
  readonly fromRevision: number;
  /** The store revision after the commit. */
  readonly toRevision: number;
  /** The committed ops, in order. Absent for undo and redo, which restore snapshots. */
  readonly ops?: readonly PptxOp[];
  /** The history label (`typing`, `move`), when the command set one. */
  readonly label?: string;
  /** True when the store merged this edit into the previous history entry (typing a word). */
  readonly merged?: boolean;
}

/** The outcome of applying a peer's ops. @public */
export type CollaborationApplyResult =
  | { readonly ok: true; readonly revision: number }
  | { readonly ok: false; readonly rejection: OpRejection };

/**
 * What the editor gives a collaboration session: the document as ops in and ops out. The session
 * owns transport, ordering, and conflict handling; the port never reads the DOM or the clock.
 *
 * @public
 */
export interface CollaborationDocumentPort {
  /** The store revision. */
  revision(): number;
  /** The current package, for a session that bootstraps a room from the opened deck. */
  package(): PptxPackage;
  /** Local transactions with their ops, in commit order. Returns the unsubscribe function. */
  subscribeLocal(listener: (transaction: CollaborationLocalTransaction) => void): () => void;
  /**
   * Apply a peer's ops as one transaction. It makes no local history entry and clears local
   * undo and redo (undo restores snapshots and would revert the peer). A refusal is returned,
   * never dropped, so the session can resynchronize.
   */
  applyRemote(ops: readonly PptxOp[], label?: string): CollaborationApplyResult;
  /**
   * Publish parts that a state-based replica rebuilt (`buildReplicaPart`), and drop `removed`
   * parts, as one remote revision. Same history rule as `applyRemote`.
   */
  applyRemoteParts(
    parts: readonly PackagePart[],
    removed: readonly string[],
    label?: string
  ): CollaborationApplyResult;
}

/**
 * A ready collaboration session, created by the host before mount. The member names follow
 * DOCX Editor's `EditorCollaborationSession`; selection, history, and journal members arrive
 * with the collaboration change.
 *
 * @public
 */
export interface EditorCollaborationSession {
  /** The room's document identity, chosen by the host before mount. */
  readonly documentId: string;
  /**
   * This replica's session identity. It must be unique per replica (per tab, not per user):
   * the editor mints new relationship ids, part names, slide ids, and shape ids from it, so
   * two replicas with one session id can mint one id twice.
   */
  readonly sessionId: string;
  /** The local author. */
  readonly identity: CollaborationIdentity;
  /** The current replica state. */
  status(): CollaborationStatus;
  /** Status, reason, and last failure. Same reference until a value changes. */
  statusSnapshot(): CollaborationStatusSnapshot;
  /** Listen to state changes, with the reason. Returns the unsubscribe function. */
  subscribeStatus(
    listener: (
      status: CollaborationStatus,
      reason?: CollaborationFailureCode,
      detail?: string
    ) => void
  ): () => void;
  /** Attach to an opened document. Returns the detach function. */
  attach(port: CollaborationDocumentPort): () => void;
  /** Whether a document is attached. */
  readonly attached: boolean;
  /** Everyone in the room, the local participant included. */
  participants(): readonly CollaborationParticipant[];
  /** Listen to participant changes. Returns the unsubscribe function. */
  subscribeParticipants(
    listener: (participants: readonly CollaborationParticipant[]) => void
  ): () => void;
  /**
   * Collaboration-owned history. When present, the editor's undo and redo go to the session,
   * which undoes only this replica's own changes. Absent, local store undo stays the authority.
   */
  canUndo?(): boolean;
  /** Whether the session can redo. */
  canRedo?(): boolean;
  /** Undo this replica's last change. Returns false when there is nothing to undo, or when the step could not apply. */
  undo?(): boolean;
  /** Redo this replica's last undone change. Returns false when there is nothing to redo, or when the step could not apply. */
  redo?(): boolean;
  /**
   * The selection the last `undo` or `redo` returns to: the selection published when that
   * change was made, re-read against the current room. `null` when there is none.
   */
  historySelection?(): CollaborationLocalSelection | null;
  /**
   * How the last remote change moved this replica's published selection, or `null` when it
   * did not. The editor follows it, so a caret stays on its character while a peer types
   * before it.
   */
  remoteSelectionMove?(): CollaborationSelectionMove | null;
  /** Publish this replica's selection to the room. `null` clears it. */
  setLocalSelection?(selection: CollaborationLocalSelection | null): void;
  /** Peers' selections, mapped onto this replica's node ids. */
  remoteSelections?(): readonly CollaborationRemoteSelection[];
  /** Listen to remote selection changes. Returns the unsubscribe function. */
  subscribeRemoteSelections?(
    listener: (selections: readonly CollaborationRemoteSelection[]) => void
  ): () => void;
  /** Leave the room and release the transport. */
  destroy(): void;
}

/** What a collaboration module contributes: the session the editor attaches. @public */
export interface CollaborationModuleContribution {
  /** A ready session. The host creates the room, then wraps it with `collaborationModule`. */
  readonly session: EditorCollaborationSession;
}

// --- Review ---------------------------------------------------------------------------------

/**
 * How the editor shows review content. `final` is the free engine's view: the deck as saved.
 * `markup` shows comment markers and suggestions.
 *
 * @public
 */
export type ReviewDisplayMode = 'final' | 'markup';

/**
 * The kinds of review item. PresentationML has comments but no tracked-change markup, so a
 * `suggestion` is a proposed set of ops that the review module holds outside the deck until a
 * reviewer accepts or rejects it.
 *
 * @public
 */
export type ReviewItemKind = 'comment' | 'suggestion';

/** One entry of the review queue. @public */
export interface ReviewItem {
  /** Stable within one package revision. */
  readonly id: string;
  /** A comment from the deck, or a suggestion held by the review module. */
  readonly kind: ReviewItemKind;
  /** The slide part the item belongs to. */
  readonly slidePart: string;
  /** The author's display name. */
  readonly author?: string;
  /** ISO 8601 date, as stored. */
  readonly date?: string;
  /** The comment text, or a summary of the suggestion. */
  readonly text: string;
  /** For a reply, the id of the comment it answers. */
  readonly parentId?: string;
}

/** What a review module reads to build its queue. @public */
export interface ReviewModelInput {
  /** The package of the current store revision. */
  readonly package: PptxPackage;
}

/** Derives the review queue. The free engine has no implementation. @public */
export type CollectReviewItems = (input: ReviewModelInput) => readonly ReviewItem[];

/** What a review module contributes. @public */
export interface ReviewModuleContribution {
  /** Display modes this module unlocks. The free engine shows `final` only. */
  readonly displayModes: readonly ReviewDisplayMode[];
  /** The review queue derivation. */
  readonly collectReviewItems: CollectReviewItems;
  /** The comment view settings: how a comment marker looks and what a press on it opens. */
  readonly pane?: ReviewPaneOptions;
}

// --- Registration ---------------------------------------------------------------------------

/**
 * One registered capability module. Registration is construction-time
 * (`createPptxEditor({ modules })`) and fixed for the instance's lifetime.
 *
 * @public
 */
export interface EditorModule {
  /** Diagnostic identity (`'review'`, `'collaboration'`); not a dispatch key. */
  readonly id: string;
  /** Review capability: the queue derivation and display modes. */
  readonly review?: ReviewModuleContribution;
  /** Absent, the editor stays single-user and local undo is the authority. */
  readonly collaboration?: CollaborationModuleContribution;
}

/** The resolved modules an editor instance dispatches over. @public */
export interface EditorModuleRegistry {
  /** The first registered review contribution, or null. */
  readonly review: ReviewModuleContribution | null;
  /** The first registered collaboration contribution, or null. */
  readonly collaboration: CollaborationModuleContribution | null;
}

const EMPTY_REGISTRY: EditorModuleRegistry = Object.freeze({ review: null, collaboration: null });

/**
 * Resolve construction-time modules. The first module of each kind wins: two review or two
 * collaboration modules is a configuration mistake, and merging them would leave neither author
 * able to say which one runs. The second is ignored, not thrown, as in DOCX Editor.
 *
 * @public
 */
export function resolveEditorModules(
  modules: readonly EditorModule[] | undefined
): EditorModuleRegistry {
  if (!modules || modules.length === 0) return EMPTY_REGISTRY;
  let review: ReviewModuleContribution | null = null;
  let collaboration: CollaborationModuleContribution | null = null;
  for (const module of modules) {
    if (module.review && review === null) review = module.review;
    if (module.collaboration && collaboration === null) collaboration = module.collaboration;
  }
  return Object.freeze({ review, collaboration });
}

/** The port over one store. Exported for the pro package's tests and for headless hosts. @public */
export function collaborationPort(store: PptxDocumentStore): CollaborationDocumentPort {
  return {
    revision: () => store.revision,
    package: () => store.package,
    subscribeLocal: (listener) => {
      // The package before each change: the store publishes changes, not snapshots.
      let last = store.package;
      return store.subscribe((change: StoreChange) => {
        const before = last;
        last = store.package;
        if (change.origin !== 'user' && change.origin !== 'undo' && change.origin !== 'redo')
          return;
        listener({
          origin: change.origin,
          before,
          after: last,
          parts: change.parts,
          fromRevision: change.fromRevision,
          toRevision: change.toRevision,
          ...(change.ops ? { ops: change.ops } : {}),
          ...(change.label ? { label: change.label } : {}),
          ...(change.merged ? { merged: true } : {}),
        });
      });
    },
    applyRemote: (ops, label) =>
      store.transact(ops, { origin: 'remote', ...(label ? { label } : {}) }),
    applyRemoteParts: (parts, removed, label) =>
      store.replaceParts(parts, removed, label ? { label } : {}),
  };
}

/**
 * Attach the registry's collaboration session to an opened store. Returns the detach function.
 * Internal to the editor lane: hosts register modules through `createPptxEditor({ modules })`.
 * While the session is attached, the store mints new ids in the session's stripe
 * (`bindReplicaIds`): bound before the attach, so a command that a listener runs during the
 * attach already mints in it, and bound again after it, so the record takes the room's numbers.
 */
export function attachEditorModules(
  registry: EditorModuleRegistry,
  store: PptxDocumentStore
): () => void {
  const session = registry.collaboration?.session;
  if (!session) return () => {};
  bindReplicaIds(store, session);
  const detach = session.attach(collaborationPort(store));
  bindReplicaIds(store, session);
  return () => {
    detach();
    bindReplicaIds(store, undefined);
  };
}

/**
 * Run undo or redo on the attached session's history, when it has one. `none`: no session
 * history, so the caller uses the store. `conflict`: the step could not apply (a peer's change
 * took its values) and the history still has steps. Internal to the editor lane.
 */
export function runSessionHistory(
  registry: EditorModuleRegistry,
  undo: boolean
): 'none' | 'ok' | 'empty' | 'conflict' {
  const session = registry.collaboration?.session;
  if (!session?.attached || !session.undo || !session.redo) return 'none';
  if (undo ? session.undo() : session.redo()) return 'ok';
  const more = undo ? session.canUndo?.() : session.canRedo?.();
  return more ? 'conflict' : 'empty';
}

/** Whether the attached session can undo or redo, or `null` without session history. */
export function canSessionHistory(registry: EditorModuleRegistry, undo: boolean): boolean | null {
  const session = registry.collaboration?.session;
  if (!session?.attached || !session.canUndo || !session.canRedo) return null;
  return undo ? session.canUndo() : session.canRedo();
}

/** The editors that hold a collaboration module. Module-private; read by the automation host. */
const COLLABORATING = new WeakSet<object>();

/**
 * Record that `editor` holds a collaboration module, then return it. `createPptxEditor` calls
 * this once, so the automation host can merge cells the way the editor does in a room.
 *
 * @internal
 */
export function withCollaborationFlag<T extends object>(
  editor: T,
  modules: EditorModuleRegistry
): T {
  if (modules.collaboration) COLLABORATING.add(editor);
  return editor;
}

/** True when `editor` holds a collaboration module. @internal */
export function isCollaborating(editor: object): boolean {
  return COLLABORATING.has(editor);
}
