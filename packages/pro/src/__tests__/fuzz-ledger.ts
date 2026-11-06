/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// The text conservation oracle of the seeded collaboration fuzzers (`*-fuzz.test.ts`). PPTX-only.
//
// Every character in a fuzz run is unique, so the ledger can follow each one: the edits that
// made it show (typing, an undo of a delete, a redo of a typing) and the edits that removed it,
// each with its Yjs clock. After a flush it checks the shown text exactly:
// - a character shows at most once;
// - a character shows when every edit that removed it was followed by an edit that brought it
//   back (`mustShow`), unless an exemption names the concurrent edit that may take it;
// - a character does not show when an edit removed it after every edit that brought it back
//   (`mustHide`).
// A removal and a revival that are concurrent leave the result open: either is a valid merge.
// So do a removal and a concurrent copy of the character to another place (`move`): a table
// merge copies the text of the cells it covers into the merged cell.
//
// An exemption is narrow: a character may be lost only when an edit that placed it is
// concurrent with a risky edit (a peer's join, or a row or column delete) whose region holds
// the place the character went to. The place is the text around the character, as its writer
// saw it, that the risky edit's writer also saw. When the two writers share no character of
// that place (an empty paragraph or cell), the place cannot be told, and any concurrent risky
// edit exempts the character; `stats.blind` counts those.

import * as Y from 'yjs';

/** What a replica had before an edit: its state vector and its delete set. */
export interface ReplicaState {
  readonly seen: ReadonlyMap<number, number>;
  readonly ds: ReturnType<typeof Y.createDeleteSetFromStructStore>;
}

/**
 * When an edit ran: its writer, what the writer had before it, the clock range of the items it
 * inserted, and the items it deleted. A Yjs delete takes no clock, so a state vector alone
 * cannot tell whether a peer had a delete: the delete set can.
 */
export interface EditClock extends ReplicaState {
  readonly client: number;
  /** The writer's own clock before the edit: the first item it inserted. */
  readonly start: number;
  /** The writer's own clock after the edit. */
  readonly clock: number;
  /** The items the edit deleted, as `[client, clock]`. */
  readonly deleted: readonly (readonly [number, number])[];
}

/** What `doc` has now: take it before an edit, and pass it to `editClock`. */
export const stateOf = (doc: Y.Doc): ReplicaState => ({
  seen: Y.decodeStateVector(Y.encodeStateVector(doc)),
  ds: Y.createDeleteSetFromStructStore(doc.store),
});

/** The clock of an edit on `doc`, after it ran, with the state taken before it. */
export function editClock(doc: Y.Doc, before: ReplicaState): EditClock {
  const deleted: [number, number][] = [];
  for (const [client, items] of Y.createDeleteSetFromStructStore(doc.store).clients)
    for (const { clock, len } of items)
      for (let at = clock; at < clock + len; at += 1)
        if (!Y.isDeleted(before.ds, Y.createID(client, at))) deleted.push([client, at]);
  return {
    ...before,
    client: doc.clientID,
    start: before.seen.get(doc.clientID) ?? 0,
    clock: Y.getState(doc.store, doc.clientID),
    deleted,
  };
}

/** Whether the edit wrote to the room: inserted or deleted an item. */
export const wrote = (edit: EditClock): boolean =>
  edit.clock > edit.start || edit.deleted.length > 0;

/** Whether the writer of `viewer` had all of `edit` when it ran: its inserts and its deletes. */
export const saw = (viewer: EditClock, edit: EditClock): boolean =>
  (edit.clock <= edit.start || (viewer.seen.get(edit.client) ?? 0) >= edit.clock) &&
  edit.deleted.every(([client, clock]) => Y.isDeleted(viewer.ds, Y.createID(client, clock)));

/** Whether both edits wrote, and neither writer had the other's edit when it ran. */
export const concurrent = (x: EditClock, y: EditClock): boolean =>
  wrote(x) && wrote(y) && !saw(x, y) && !saw(y, x);

/** An edit that may take text that a peer places at the same time. */
export interface RiskyEdit {
  readonly clock: EditClock;
  /** The characters of the paragraphs or cells it removes or moves. */
  readonly region: ReadonlySet<string>;
  /** Every character its writer showed before it. */
  readonly view: ReadonlySet<string>;
}

/** An edit that made a character show, and the characters around it then. */
interface Placement {
  readonly clock: EditClock;
  readonly place: ReadonlySet<string>;
}

const quote = (chars: Iterable<string>): string => JSON.stringify([...chars].join(''));

/** Characters of `chars` that `set` lacks. */
export const outside = (chars: Iterable<string>, set: ReadonlySet<string>): string[] =>
  [...chars].filter((ch) => !set.has(ch));

/** The text conservation ledger of one run. */
export class CharacterLedger {
  private readonly deck: ReadonlySet<string>;
  private readonly placements = new Map<string, Placement[]>();
  private readonly removals = new Map<string, EditClock[]>();
  private readonly risky: RiskyEdit[] = [];
  /** The edits that copied a character to another place. */
  private readonly moves = new Map<string, EditClock[]>();
  /** Checks run, exemptions that applied, and exemptions that could not tell the place. */
  readonly stats = { characterChecks: 0, exempt: 0, blind: 0 };

  constructor(deckText: string) {
    this.deck = new Set(deckText);
  }

  /** Record what one peer's own edit changed on its replica. */
  record(
    clock: EditClock,
    gone: Iterable<string>,
    came: Iterable<string>,
    placeOf: (ch: string) => ReadonlySet<string>
  ): void {
    for (const ch of gone) push(this.removals, ch, clock);
    for (const ch of came) push(this.placements, ch, { clock, place: placeOf(ch) });
  }

  /** Record an edit that copies `chars` to another place, so a concurrent removal may lose. */
  move(clock: EditClock, chars: Iterable<string>): void {
    if (wrote(clock)) for (const ch of chars) push(this.moves, ch, clock);
  }

  /** Record an edit that may take text a peer places at the same time. */
  risk(edit: RiskyEdit): void {
    if (wrote(edit.clock)) this.risky.push(edit);
  }

  /**
   * What an undo or a redo by `client` may bring back: what `client` removed itself, and what
   * no edit removed for good (text that a concurrent join or delete took).
   */
  mayBringBack(client: number): Set<string> {
    const own = keysWhere(this.removals, (list) => list.some((c) => c.client === client));
    for (const ch of [...this.deck, ...this.placements.keys()]) if (!this.mustHide(ch)) own.add(ch);
    return own;
  }

  /** The characters `client` placed in an edit of its own: what its undo or redo may remove. */
  placedBy(client: number): Set<string> {
    return keysWhere(this.placements, (list) => list.some((p) => p.clock.client === client));
  }

  /** Every edit that removed `ch` was followed, on its writer's replica, by one that placed it. */
  private mustShow(ch: string): boolean {
    const placed = this.placements.get(ch) ?? [];
    return (this.removals.get(ch) ?? []).every((r) => placed.some((p) => saw(p.clock, r)));
  }

  /** Some edit removed `ch` after every edit that placed it, and no edit copied it meanwhile. */
  private mustHide(ch: string): boolean {
    const placed = this.placements.get(ch) ?? [];
    const moved = this.moves.get(ch) ?? [];
    return (this.removals.get(ch) ?? []).some(
      (r) => placed.every((p) => saw(r, p.clock)) && !moved.some((m) => concurrent(m, r))
    );
  }

  /** Whether an edit that placed `ch` is concurrent with a risky edit whose region holds it. */
  private exempt(ch: string): boolean {
    for (const p of this.placements.get(ch) ?? [])
      for (const risky of this.risky) {
        if (!concurrent(p.clock, risky.clock)) continue;
        const shared = [...p.place].filter((c) => risky.view.has(c));
        if (shared.length === 0) {
          this.stats.blind += 1;
          this.stats.exempt += 1;
          return true;
        }
        if (shared.some((c) => risky.region.has(c))) {
          this.stats.exempt += 1;
          return true;
        }
      }
    return false;
  }

  /** The conservation problems of the shown text. */
  problems(text: string): string[] {
    const problems: string[] = [];
    const counts = new Map<string, number>();
    for (const ch of text) counts.set(ch, (counts.get(ch) ?? 0) + 1);
    const twice = [...counts].filter(([, n]) => n > 1).map(([ch]) => ch);
    if (twice.length > 0) problems.push(`characters show more than once: ${quote(twice)}`);
    const known = new Set([...this.deck, ...this.placements.keys()]);
    this.stats.characterChecks += known.size;
    const strangers = [...counts.keys()].filter((ch) => !known.has(ch));
    if (strangers.length > 0)
      problems.push(`characters show that no edit typed: ${quote(strangers)}`);
    const lost = [...known].filter(
      (ch) => !counts.has(ch) && this.mustShow(ch) && !this.exempt(ch)
    );
    if (lost.length > 0) problems.push(`characters that no peer removed are lost: ${quote(lost)}`);
    const back = [...counts.keys()].filter((ch) => known.has(ch) && this.mustHide(ch));
    if (back.length > 0) problems.push(`removed characters show again: ${quote(back)}`);
    return problems;
  }
}

function push<T>(map: Map<string, T[]>, key: string, value: T): void {
  const list = map.get(key);
  if (list) list.push(value);
  else map.set(key, [value]);
}

function keysWhere<T>(map: ReadonlyMap<string, T[]>, test: (list: T[]) => boolean): Set<string> {
  return new Set([...map].filter(([, list]) => test(list)).map(([key]) => key));
}

/** What an edit may change on its own replica. */
export interface LocalRule {
  /** The characters it may remove. */
  readonly mayRemove: ReadonlySet<string>;
  /** The characters it must remove (a delete that the editor accepted). */
  readonly mustRemove?: ReadonlySet<string>;
  /** The characters it may add. */
  readonly mayAdd: ReadonlySet<string>;
  /** The characters it must add (a typing that the editor accepted). */
  readonly mustAdd?: ReadonlySet<string>;
}

/**
 * The problems of one edit on its own replica: it removed or added only what its selection and
 * its kind allow. A remote change cannot arrive during a local edit, so the diff is the edit's.
 */
export function localProblems(
  what: string,
  gone: ReadonlySet<string>,
  came: ReadonlySet<string>,
  rule: LocalRule
): string[] {
  const problems: string[] = [];
  const extraGone = outside(gone, rule.mayRemove);
  if (extraGone.length > 0)
    problems.push(`${what} removed characters it may not remove: ${quote(extraGone)}`);
  const kept = outside(rule.mustRemove ?? [], gone);
  if (kept.length > 0) problems.push(`${what} kept characters it removes: ${quote(kept)}`);
  const extraCame = outside(came, rule.mayAdd);
  if (extraCame.length > 0)
    problems.push(`${what} added characters it may not add: ${quote(extraCame)}`);
  const missing = outside(rule.mustAdd ?? [], came);
  if (missing.length > 0) problems.push(`${what} did not add characters: ${quote(missing)}`);
  return problems;
}

export const NONE: ReadonlySet<string> = new Set();

/**
 * `rule` without the characters that show more than once in `text`, the replica's text before
 * the edit, in what the edit must remove: a removed copy leaves the other. The round check
 * reports the repeated character.
 */
export function withoutRepeated(rule: LocalRule, text: string): LocalRule {
  if (!rule.mustRemove) return rule;
  const seen = new Set<string>();
  const twice = new Set<string>();
  for (const ch of text) (seen.has(ch) ? twice : seen).add(ch);
  return { ...rule, mustRemove: new Set(outside(rule.mustRemove, twice)) };
}
