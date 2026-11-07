// Build one transaction op by op against a working package, so later ops can address nodes
// that earlier ops created. Ids minted by `insertNode` are deterministic (`n<nextId>` of the
// part), so the store re-applying the same ops produces the same ids.

import {
  applyOps,
  nodeIndex,
  openTransaction,
  xmlPartOf,
  type Element,
  type OpenTransaction,
  type OpRejection,
  type PptxOp,
  type PptxPackage,
} from '../../store/index.ts';

/** Thrown inside a command builder to abort the whole transaction with a reason. */
export class CommandRefusal extends Error {
  constructor(
    readonly reason: string,
    readonly detail?: string
  ) {
    super(reason);
  }
}

/**
 * Reasons that say a command had nothing to do, not that an edit was refused: nothing to undo
 * or redo, nothing selected. A rejection with one of them carries `noop`.
 */
const NOTHING_TO_DO: ReadonlySet<string> = new Set([
  'disabled.nothingToUndo',
  'disabled.nothingToRedo',
  'disabled.noSelection',
  'disabled.noVisibilityChange',
  'disabled.noEffectChange',
  'disabled.noLockChange',
  'disabled.noArrangeChange',
  'disabled.noNameChange',
  'disabled.noSizeChange',
  'disabled.noAdjustChange',
  'disabled.noAnimationChange',
]);

/** Whether a refusal reason means the command had nothing to do. */
export function isNothingToDo(reason: string): boolean {
  return NOTHING_TO_DO.has(reason);
}

/** The kinds of number a replica mints in its stripe, each with its own record. @internal */
export type MintedIdKind =
  | 'rId'
  | 'part'
  | 'slideId'
  | 'shapeId'
  | 'creationId'
  | 'commentIdx'
  | 'timeNodeId'
  | 'buildGroupId';

/** Numbers a replica minted in its stripe, by id kind. @internal */
export type MintedIds = Map<MintedIdKind, Set<number>>;

/** Options for an `OpBatch`. */
export interface OpBatchOptions {
  /**
   * Case-folded part names (`partNameKey`) that new parts must not take: every name the
   * session has held. A name then always means one part, also across delete and undo.
   */
  readonly reservedPartNames?: ReadonlySet<string>;
  /**
   * The actor while a collaboration session is attached (its session id): new relationship
   * ids, part names, slide ids, and shape ids are minted in its stripe (`replica-ids.ts`).
   * @internal
   */
  readonly allocationActor?: string;
  /**
   * Numbers this replica minted in its stripe during the session, one set per id kind
   * (`replica-ids.ts`). @internal
   */
  readonly minted?: MintedIds;
  /**
   * A probe (the `can` check): numbers it mints go into copies of the record, so a probe uses
   * up no number. @internal
   */
  readonly probe?: boolean;
}

const NO_NAMES: ReadonlySet<string> = new Set();

export class OpBatch {
  #pkg: PptxPackage;
  /**
   * The transaction the ops since the last read of `pkg` apply to, and those ops. One
   * transaction for a run of pushes: a format check pushes an op for each paragraph, and a
   * transaction per op copied the part's id index each time.
   */
  #tx: OpenTransaction | null = null;
  #open: PptxOp[] = [];
  readonly ops: PptxOp[] = [];
  /** Case-folded part names that new parts must not take. */
  readonly reservedPartNames: ReadonlySet<string>;
  /**
   * Set when a command formatted an empty run at a caret: a format armed for the text typed
   * next. The editor keeps such a batch as session state and commits it with that text.
   */
  armedCaret = false;
  /** The actor that new ids are minted for, or undefined with no collaboration session. @internal */
  readonly allocationActor: string | undefined;
  /** Numbers this replica minted in its stripe during the session, by id kind. */
  readonly #minted: MintedIds;
  /** A probe's copies of the record, made on first use. */
  #probeCopies: MintedIds | undefined;

  constructor(pkg: PptxPackage, options: OpBatchOptions = {}) {
    this.#pkg = pkg;
    this.reservedPartNames = options.reservedPartNames ?? NO_NAMES;
    this.allocationActor = options.allocationActor;
    this.#minted = options.minted ?? new Map();
    this.#probeCopies = options.probe ? new Map() : undefined;
  }

  /**
   * The numbers this replica minted for one id kind, so one kind never uses up another's
   * numbers. A probe gets a copy. @internal
   */
  mintedOf(kind: MintedIdKind): Set<number> {
    if (this.#probeCopies) {
      let copy = this.#probeCopies.get(kind);
      if (!copy) this.#probeCopies.set(kind, (copy = new Set(this.#minted.get(kind))));
      return copy;
    }
    let set = this.#minted.get(kind);
    if (!set) this.#minted.set(kind, (set = new Set()));
    return set;
  }

  /** The working package after every op so far. */
  get pkg(): PptxPackage {
    this.#settle();
    return this.#pkg;
  }

  /** Close the open transaction into `#pkg`. */
  #settle(): void {
    if (!this.#tx) return;
    this.#pkg = this.#tx.finish();
    this.#tx = null;
    this.#open = [];
  }

  /** Apply an op to the working package, or refuse the transaction. */
  push(op: PptxOp): void {
    this.#tx ??= openTransaction(this.#pkg);
    const rejection = this.#tx.apply(op);
    if (rejection) {
      // A refused op leaves the open transaction unusable: rebuild the package from the ops
      // that applied, so the batch stays what its `ops` say.
      const kept = applyOps(this.#pkg, this.#open);
      this.#tx = null;
      this.#open = [];
      if (kept.ok) this.#pkg = kept.pkg;
      throw new CommandRefusal('errors.unsupported', describe(rejection));
    }
    this.#open.push(op);
    this.ops.push(op);
  }

  /**
   * Apply many ops in one pass over the working package, or refuse the transaction. Each op
   * still lands in `ops`, so the store re-applies the same ops; a bulk copy then costs one
   * package update, not one for each part.
   */
  pushAll(ops: readonly PptxOp[]): void {
    if (ops.length === 0) return;
    this.#settle();
    const r = applyOps(this.#pkg, ops);
    if (!r.ok) throw new CommandRefusal('errors.unsupported', describe(r.rejection));
    this.#pkg = r.pkg;
    for (const op of ops) this.ops.push(op);
  }

  /** Insert and return the new element's id. */
  insert(op: Extract<PptxOp, { op: 'insertNode' }>): string {
    // From the open transaction: reading `pkg` would close it before each insert.
    const before =
      (this.#tx ? this.#tx.nextId(op.part) : xmlPartOf(this.#pkg, op.part)?.nextId) ?? 0;
    this.push(op);
    return `n${before}`;
  }

  /** The parent id of node `id` in a part, or undefined for the root or a missing node. */
  parentId(part: string, id: string): string | undefined {
    // From the open transaction: reading `pkg` would close it before each lookup.
    if (this.#tx) return this.#tx.parentId(part, id) ?? undefined;
    const xml = xmlPartOf(this.#pkg, part);
    return (xml ? nodeIndex(xml).get(id)?.parent : undefined) ?? undefined;
  }

  /**
   * The kind of the parent of node `id` (`grpSp` for a group), read without holding the parent:
   * an edit after it writes in place. Undefined for the root or a missing node.
   */
  parentKind(part: string, id: string): { id: string; kind: string } | undefined {
    const parentId = this.parentId(part, id);
    if (!parentId) return undefined;
    const parent = this.#tx ? this.#tx.peek(part, parentId) : undefined;
    const kind = (parent ?? this.element(part, parentId)).kind;
    return { id: parentId, kind };
  }

  /** The current element for an id in a part. */
  element(part: string, id: string): Element {
    // From the open transaction: reading `pkg` would close it before each op.
    let node: Element | undefined;
    if (this.#tx) node = this.#tx.element(part, id);
    else {
      const xml = xmlPartOf(this.#pkg, part);
      const found = xml ? nodeIndex(xml).get(id)?.node : undefined;
      node = found?.type === 'element' ? found : undefined;
    }
    if (!node || node.type !== 'element')
      throw new CommandRefusal('errors.unsupported', `missing node ${id}`);
    return node;
  }
}

/**
 * The `detail` of a command the store refused at commit: the rejection code, and for a broken
 * part also the rule, the part, and what is wrong, for example
 * `tree-invariant table-grid /ppt/slides/slide1.xml: row 0 has 1 a:tc for 2 a:gridCol`.
 */
export function commitRefusalDetail(r: OpRejection): string {
  if (r.code === 'tree-invariant' || r.code === 'package-invariant')
    return `${r.code} ${r.rule} ${r.part}: ${r.detail}`;
  return r.code;
}

function describe(r: OpRejection): string {
  return Object.entries(r)
    .map(([k, v]) => `${k}=${String(v)}`)
    .join(' ');
}
