// The document store: the single source of truth and the only write path.
//
// One transaction = one atomic publication = at most one history entry (DOCX Editor's
// TreeDocumentStore contract, store/store/tree-store.ts). `transact` stages every op against a
// working package; nothing is visible until all ops apply, so a refused op mid-batch leaves the
// revision, the package, and every subscriber exactly as they were.
//
// History is scoped by intent, never by a timer: consecutive transactions with the same
// `historyGroup` token are one entry, and a `projection` or `awareness` commit is none
// (`history-group.ts`). Each entry keeps the selection before and after it.
//
// The commit checks the parts a transaction touched against the PresentationML rules of
// `part-checks.ts` (DOCX Editor's `commitTransaction` validates its touched parts the same way):
// a transaction that leaves a part broken is refused like a refused op, before anything is
// published. Remote transactions are checked too; the refusal goes back to the caller.
//
// History entries are package snapshots. The package is persistent and structurally shared,
// so an entry costs one Map of part references, and undo is a pointer swap.
//
// Listeners are isolated: one that throws does not stop delivery to the others, and never makes
// a committed transaction look failed. A transaction made inside a listener is delivered after
// the current change has reached every listener, so every listener sees changes in order.

import { normalizePartName } from '@eigenpal/ooxml-core';
import {
  changedPartNames,
  CONTENT_TYPES_PART,
  type PackagePart,
  type PptxPackage,
} from '../package/package.ts';
import { carryFileTimings } from '../package/timing.ts';
import { applyOps, type OpRejection, type PptxOp } from './ops.ts';
import type { HistoryGroup, TransactOrigin } from './history-group.ts';
import { introducedViolation } from './part-checks.ts';

/** What a committed transaction changed. */
export interface StoreChange {
  readonly fromRevision: number;
  readonly toRevision: number;
  readonly origin: 'undo' | 'redo' | 'load' | TransactOrigin;
  /** Part names whose content changed. Layout re-does only what these touch. */
  readonly parts: readonly string[];
  readonly label?: string;
  /**
   * The ops of a `user` or `remote` transaction, as committed. Undo and redo restore snapshots
   * and carry none. A collaboration replica sends the `user` ops to its peers.
   */
  readonly ops?: readonly PptxOp[];
  /**
   * True when a recorded transaction continued the open history group (`historyGroup`). It is
   * also true for the first frame after a `remote` commit cleared the stacks in the middle of
   * a gesture: the local entry is new, but the gesture goes on, so a collaboration session
   * that owns undo keeps it one step.
   */
  readonly merged?: boolean;
}

/** The outcome of a transaction. */
export type TransactResult =
  | { readonly ok: true; readonly revision: number }
  | { readonly ok: false; readonly rejection: OpRejection };

/** Options for one transaction. */
export interface TransactOptions {
  /** A short label for the history entry (`move`, `typing`). */
  readonly label?: string;
  /**
   * The gesture this transaction belongs to: consecutive transactions that carry the SAME
   * token extend one history entry (the package from before the first, the selection after
   * the latest). Another token, none, undo, or redo closes the group. See {@link HistoryGroup}.
   */
  readonly historyGroup?: HistoryGroup;
  /** Opaque selection state restored when this entry is undone. */
  readonly selectionBefore?: unknown;
  /**
   * Opaque selection state restored when this entry is redone. A later frame of an open
   * group replaces it; a frame that omits it keeps the previous one.
   */
  readonly selectionAfter?: unknown;
  /** What made the transaction; `user` when omitted. See {@link TransactOrigin}. */
  readonly origin?: TransactOrigin;
  /**
   * Whether this transaction enters the undo stack. Omitted: `true` for `user`, `false` for
   * `projection` and `awareness`. A `remote` transaction never records history and always
   * clears it.
   */
  readonly recordsHistory?: boolean;
}

interface HistoryEntry {
  readonly before: PptxPackage;
  readonly after: PptxPackage;
  readonly parts: readonly string[];
  readonly label?: string;
  readonly selectionBefore?: unknown;
  readonly selectionAfter?: unknown;
  /**
   * The gesture this entry still collects, while it is the newest entry and nothing closed
   * it. Only the top of the undo stack carries one: `closeHistoryGroup` strips it, and undo
   * and redo never copy it onto the entries they push.
   */
  readonly group?: HistoryGroup;
}

/** The selection fields of an entry, without `undefined` values. */
function selections(before: unknown, after: unknown) {
  return {
    ...(before !== undefined ? { selectionBefore: before } : {}),
    ...(after !== undefined ? { selectionAfter: after } : {}),
  };
}

const HISTORY_LIMIT = 200;

/** Options for a store. */
export interface DocumentStoreOptions {
  /**
   * Called when a listener throws. Delivery continues to the other listeners either way. The
   * default reports the error to the host (`reportError` where it exists, otherwise an uncaught
   * error on a microtask), so it is never swallowed.
   */
  readonly onListenerError?: (error: unknown, change: StoreChange) => void;
}

function reportListenerError(error: unknown): void {
  const host = globalThis as {
    reportError?: (e: unknown) => void;
    queueMicrotask?: (task: () => void) => void;
  };
  if (typeof host.reportError === 'function') host.reportError(error);
  else if (typeof host.queueMicrotask === 'function')
    host.queueMicrotask(() => {
      throw error;
    });
  else void Promise.reject(error);
}

/** Part names whose part object differs between two packages. */
export function changedParts(a: PptxPackage, b: PptxPackage): string[] {
  return changedPartNames(a, b, Infinity) ?? [];
}

/** The transactional document store. */
export class PptxDocumentStore {
  #pkg: PptxPackage;
  #revision = 0;
  #undo: HistoryEntry[] = [];
  #redo: HistoryEntry[] = [];
  #listeners = new Set<(change: StoreChange) => void>();
  #pending: StoreChange[] = [];
  #delivering = false;
  /** Per part, the highest id counter any revision reached. Undo never lowers it. */
  #highWater = new Map<string, number>();
  /**
   * The group that was open when a `remote` commit cleared the stacks. The next frame of that
   * group reports `merged`, as it would have without the peer's edit.
   */
  #cutGroup: HistoryGroup | undefined;
  readonly #onListenerError: (error: unknown, change: StoreChange) => void;
  /** Why the store refuses every commit, once `block` ran. */
  #blocked: string | null = null;

  constructor(pkg: PptxPackage, options: DocumentStoreOptions = {}) {
    this.#pkg = pkg;
    this.#onListenerError = options.onListenerError ?? reportListenerError;
  }

  /**
   * Refuse every later commit, undo, and redo with `unsupported` and `detail`: a slide of a
   * progressive open did not parse, so the package holds a stand-in that must never become an
   * edit (`deferred-parts.ts`). There is no way back. @internal
   */
  block(detail: string): void {
    this.#blocked ??= detail;
  }

  /** Why the store refuses every commit, or null. @internal */
  get blocked(): string | null {
    return this.#blocked;
  }

  /** The current package. Reference-stable until the next commit. */
  get package(): PptxPackage {
    return this.#pkg;
  }

  /** Monotonic revision, bumped by every commit (including undo and redo). */
  get revision(): number {
    return this.#revision;
  }

  get canUndo(): boolean {
    return this.#blocked === null && this.#undo.length > 0;
  }

  get canRedo(): boolean {
    return this.#blocked === null && this.#redo.length > 0;
  }

  /** Whether any commit landed since the store opened (undo history aside). @internal */
  get edited(): boolean {
    return this.#revision > 0;
  }

  /** Validate ops, and the parts they leave, without committing. */
  dryRun(ops: readonly PptxOp[]): TransactResult {
    const r = this.#stage(ops);
    return r.ok ? { ok: true, revision: this.#revision } : r;
  }

  /**
   * Apply ops to the current package and check the parts they touched. A refused op and a
   * broken part are refused alike: the caller gets the rejection, and nothing is published.
   */
  #stage(
    ops: readonly PptxOp[]
  ):
    | { readonly ok: true; readonly pkg: PptxPackage; readonly parts: string[] }
    | { readonly ok: false; readonly rejection: OpRejection } {
    const r = applyOps(this.#pkg, ops);
    if (!r.ok) return r;
    const parts = changedParts(this.#pkg, r.pkg);
    const v = introducedViolation(this.#pkg, r.pkg, parts);
    if (!v) return { ok: true, pkg: r.pkg, parts };
    const { part, detail } = v;
    const rejection: OpRejection =
      v.kind === 'tree'
        ? { code: 'tree-invariant', part, rule: v.rule, detail }
        : { code: 'package-invariant', part, rule: v.rule, detail };
    return { ok: false, rejection };
  }

  /** Apply ops atomically and publish one change. */
  transact(ops: readonly PptxOp[], options: TransactOptions = {}): TransactResult {
    const origin = options.origin ?? 'user';
    const records = origin !== 'remote' && (options.recordsHistory ?? origin === 'user');
    // An intent that names another group (or none) closes the open one HERE, whether it goes
    // on to land, to change nothing, or to be refused. A frame of the same group keeps it open.
    if (records && this.#openGroup() !== options.historyGroup) this.closeHistoryGroup();
    if (ops.length === 0) return { ok: true, revision: this.#revision };
    if (this.#blocked !== null)
      return { ok: false, rejection: { code: 'unsupported', detail: this.#blocked } };
    const r = this.#stage(ops);
    if (!r.ok) return r;
    const before = this.#pkg;
    const parts = r.parts;
    if (origin === 'remote') {
      this.#clearForRemote();
      return this.#publish(r.pkg, 'remote', parts, options.label, ops);
    }
    if (!records) return this.#publish(r.pkg, origin, parts, options.label, ops);
    const group = options.historyGroup;
    const top = this.#undo[this.#undo.length - 1];
    const extendsTop = group !== undefined && top?.group === group;
    // A remote commit emptied the stacks during this gesture: a new local entry, same gesture.
    const merged = extendsTop || (group !== undefined && this.#cutGroup === group);
    this.#cutGroup = undefined;
    if (extendsTop && top) {
      // Same gesture, still open: the entry keeps its `before` and `selectionBefore` (the
      // state before the FIRST frame) and takes this frame's package and selection as its end.
      this.#undo[this.#undo.length - 1] = {
        ...top,
        after: r.pkg,
        parts: [...new Set([...top.parts, ...parts])],
        ...selections(top.selectionBefore, options.selectionAfter ?? top.selectionAfter),
      };
    } else {
      this.#undo.push({
        before,
        after: r.pkg,
        parts,
        ...(options.label ? { label: options.label } : {}),
        ...selections(options.selectionBefore, options.selectionAfter),
        ...(group !== undefined ? { group } : {}),
      });
      if (this.#undo.length > HISTORY_LIMIT) this.#undo.shift();
    }
    this.#redo = [];
    return this.#publish(r.pkg, origin, parts, options.label, ops, merged);
  }

  /**
   * Publish parts rebuilt by a collaboration replica as one `remote` revision: `parts` replace
   * or add parts, `removed` names parts to drop. Like a remote transaction, it makes no history
   * entry and clears local undo and redo. The replica validates XML parts with
   * `buildReplicaPart`; this method checks part names and that the content types part stays.
   */
  replaceParts(
    parts: readonly PackagePart[],
    removed: readonly string[] = [],
    options: { readonly label?: string } = {}
  ): TransactResult {
    if (parts.length === 0 && removed.length === 0) return { ok: true, revision: this.#revision };
    if (this.#blocked !== null)
      return { ok: false, rejection: { code: 'unsupported', detail: this.#blocked } };
    const next = new Map(this.#pkg.parts);
    for (const part of parts) {
      const name = normalizePartName(part.partName);
      if (!name.ok || name.partName !== part.partName)
        return { ok: false, rejection: { code: 'unsafe-part-name', part: part.partName } };
      if (part.kind === 'xml' && part.xml.partName !== part.partName)
        return {
          ok: false,
          rejection: { code: 'invalid-part-content', part: part.partName, detail: 'name mismatch' },
        };
      // A slide's `p:timing` that arrives as the file wrote it saves as the file wrote it here
      // too, as on the replica that sent it.
      const prior = this.#pkg.parts.get(part.partName);
      if (prior?.kind === 'xml' && part.kind === 'xml')
        carryFileTimings(prior.xml.root, part.xml.root);
      next.set(part.partName, part);
    }
    for (const name of removed) {
      if (name === CONTENT_TYPES_PART)
        return { ok: false, rejection: { code: 'unsafe-part-name', part: name } };
      next.delete(name);
    }
    const before = this.#pkg;
    const after: PptxPackage = { parts: next };
    const changed = changedParts(before, after);
    this.#clearForRemote();
    return this.#publish(after, 'remote', changed, options.label);
  }

  /** The group a recorded transaction would continue: the top entry's, or one a peer cut. */
  #openGroup(): HistoryGroup | undefined {
    const top = this.#undo[this.#undo.length - 1];
    return top ? top.group : this.#cutGroup;
  }

  /** A peer's edit: no local undo can apply over it. An open gesture stays open. */
  #clearForRemote(): void {
    this.#cutGroup = this.#openGroup();
    this.#undo = [];
    this.#redo = [];
  }

  /**
   * Close the open history group, if any: the next transaction records its own entry. The
   * store does this itself for another token, none, undo, and redo. A caller does it for a
   * boundary the store cannot see, such as a selection change.
   */
  closeHistoryGroup(): void {
    this.#cutGroup = undefined;
    const top = this.#undo[this.#undo.length - 1];
    if (top?.group === undefined) return;
    const { group: _closed, ...rest } = top;
    this.#undo[this.#undo.length - 1] = rest;
  }

  /**
   * Set the selection after the newest history entry, for a caller that knows it only after
   * the commit published (a command that adds a slide). Returns false, and changes nothing,
   * when the newest entry did not produce the current package.
   */
  setSelectionAfter(selection: unknown): boolean {
    const top = this.#undo[this.#undo.length - 1];
    if (!top || top.after !== this.#pkg) return false;
    this.#undo[this.#undo.length - 1] = { ...top, ...selections(top.selectionBefore, selection) };
    return true;
  }

  /** The selection that `undo()` restores next: the newest entry's `selectionBefore`. */
  selectionForUndo(): unknown {
    return this.#undo[this.#undo.length - 1]?.selectionBefore;
  }

  /** The selection that `redo()` restores next: the newest redo entry's `selectionAfter`. */
  selectionForRedo(): unknown {
    return this.#redo[this.#redo.length - 1]?.selectionAfter;
  }

  /** Undo one entry. Returns the selection captured before it, if any. */
  undo(): { readonly ok: boolean; readonly selection?: unknown } {
    if (this.#blocked !== null) return { ok: false };
    const entry = this.#undo.pop();
    if (!entry) return { ok: false };
    // Undo is a boundary: the entry now on top must not collect a gesture that resumes.
    this.closeHistoryGroup();
    const { group: _closed, ...rest } = entry;
    // Redo returns to the package as it is now, so it keeps a history-free commit made since.
    const current = this.#pkg;
    this.#redo.push({ ...rest, after: current, parts: changedParts(entry.before, current) });
    this.#publish(entry.before, 'undo', changedParts(current, entry.before), entry.label);
    return { ok: true, selection: entry.selectionBefore };
  }

  /** Redo one entry. Returns the selection captured after it, if any. */
  redo(): { readonly ok: boolean; readonly selection?: unknown } {
    if (this.#blocked !== null) return { ok: false };
    const entry = this.#redo.pop();
    if (!entry) return { ok: false };
    this.closeHistoryGroup();
    const current = this.#pkg;
    const parts = changedParts(current, entry.after);
    this.#undo.push({ ...entry, before: current, parts });
    this.#publish(entry.after, 'redo', parts, entry.label);
    return { ok: true, selection: entry.selectionAfter };
  }

  /** Subscribe to commits. Returns the unsubscribe function. */
  subscribe(listener: (change: StoreChange) => void): () => void {
    this.#listeners.add(listener);
    return () => this.#listeners.delete(listener);
  }

  /**
   * Keep node ids unique across history. Undo restores a revision whose id counter is lower
   * than one already used; an insert after the undo would mint those ids again, and a selection
   * or a peer that still holds an id from the undone revision would address the new node.
   * Restored parts get the highest counter their part ever reached.
   */
  #withIdFloor(pkg: PptxPackage, parts: readonly string[]): PptxPackage {
    let out: Map<string, PackagePart> | null = null;
    for (const name of parts) {
      const part = pkg.parts.get(name);
      if (!part || part.kind !== 'xml') continue;
      const floor = this.#highWater.get(name) ?? 0;
      if (part.xml.nextId >= floor) continue;
      out ??= new Map(pkg.parts);
      out.set(name, { ...part, xml: { ...part.xml, nextId: floor } });
    }
    return out ? { parts: out } : pkg;
  }

  #recordIds(pkg: PptxPackage, parts: readonly string[]): void {
    for (const name of parts) {
      const part = pkg.parts.get(name);
      if (part?.kind === 'xml' && part.xml.nextId > (this.#highWater.get(name) ?? 0))
        this.#highWater.set(name, part.xml.nextId);
    }
  }

  #publish(
    next: PptxPackage,
    origin: StoreChange['origin'],
    parts: readonly string[],
    label?: string,
    ops?: readonly PptxOp[],
    merged = false
  ): TransactResult {
    // Record the counters of the revision being left, then raise the restored one to them.
    this.#recordIds(this.#pkg, parts);
    // `insertNode` takes ids from the part counter. Transactions keep the counter as applied,
    // so two replicas that apply the same ops in the same order mint the same ids. Undo and
    // redo raise it (the id floor), which a collaboration session must account for.
    const pkg = origin === 'undo' || origin === 'redo' ? this.#withIdFloor(next, parts) : next;
    this.#recordIds(pkg, parts);
    const from = this.#revision;
    this.#pkg = pkg;
    this.#revision += 1;
    this.#pending.push({
      fromRevision: from,
      toRevision: this.#revision,
      origin,
      parts,
      ...(label ? { label } : {}),
      ...(ops ? { ops } : {}),
      ...(merged ? { merged: true } : {}),
    });
    const revision = this.#revision;
    this.#deliver();
    return { ok: true, revision };
  }

  /** Deliver queued changes in order. Re-entrant calls only queue; the outer loop delivers. */
  #deliver(): void {
    if (this.#delivering) return;
    this.#delivering = true;
    try {
      for (let change = this.#pending.shift(); change; change = this.#pending.shift()) {
        for (const listener of [...this.#listeners]) {
          try {
            listener(change);
          } catch (error) {
            this.#onListenerError(error, change);
          }
        }
      }
    } finally {
      this.#delivering = false;
    }
  }
}
