/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Where merged text shows when peers change one table's structure at the same time. A cell
// merge moves the paragraphs of the cells it covers into the cell it keeps (`mergeRange` in
// core's `editor/table-model.ts`), and each cell it left keeps one empty paragraph. The rule:
//
//   Moved text shows in the drawn cell that covers the cell it sits in.
//
// A covered cell is not drawn (ECMA-376 Part 1, 21.1.3.16 tc), so text that a concurrent edit
// leaves in one would be lost. Three cases leave text there:
//
// - A peer deletes the row or the column of the merged cell without seeing the merge. The
//   rescue (`move-rescue.ts`) puts each moved paragraph back into the cell it came from, and
//   that cell can be covered (a span that lost its first cell).
// - A peer's merge covers the cell that another merge moved text into (two overlapping merges).
// - A peer types into a cell that a merge covers while it was empty. The merge replaces the
//   cell's empty paragraph with a copy, and the removed paragraph comes back with the peer's
//   text (`withRemovedRuns` in `text-identity.ts`).
//
// Such a paragraph goes to the end of the text body of the drawn cell over it, in reading order
// of the cells it sat in. The grid is the one the editor draws (`tableGridOf`), so the canvas,
// find, and the save agree. A covered cell keeps at least one paragraph (21.1.2.1.1): when every
// paragraph would go, its last one stays.
//
// The read must give every replica's store what the room shows, also the store of a replica
// that wrote the state itself, which does not read its own write. So the rule acts only on
// states that no single writer makes:
//
// - A paragraph the rescue put back in this read: a rescue needs a removal that did not see it.
// - A paragraph with text that no writer saw hidden in that cell, in a table that a structure
//   command marked (`table-marks.ts` in core) or that a move (`mf`) or a removal (`rw`) touched.
//   A write that leaves a table frame's cell covered records the paragraphs the cell holds
//   (`cv.<client>`), so a merge that covers a cell, a move into a covered cell, and a span that
//   grows over a cell each count as seen, and that text stays hidden as a file would hide it.
//   A table that no structure command changed keeps hidden text that its file brought.
//
// The empty paragraph that a merge left in a cell (`mh`) leaves the read when text the rescue
// put back shows in that cell, so the cell shows its text without an extra empty line.
//
// Every replica reads the same room, so all of them show the same. The next local write of the
// part lists the children of each text body that the rule changed as the store holds them, so a
// later structure edit that uncovers the old cell does not show the text there on other
// replicas while the writer, which does not read its own write, shows it in the merged cell.
// That write moves the text in the room for good: a peer who removes the old cell at the same
// time, and saw the text there, does not take it with the cell (a known gap). A refused read
// notes no text body, and drops what an older read noted, so a write never lists a body as an
// older read showed it.
//
// Everything read here comes from the room and is untrusted: a field of a wrong type is
// ignored, the walk stops at the depth the reader caps, and a table larger than the editor draws
// keeps its paragraphs where they are.

import * as Y from 'yjs';
import {
  children,
  elements,
  isStructureEdited,
  path,
  TABLE_MAX_COLUMNS,
  TABLE_MAX_ROWS,
  tableGridOf,
  type Element,
  type Node,
  type NodeId,
  type ReplicaElementSpec,
} from '@pptx-editor.dev/core/store';
import type { Room } from './schema.ts';
import { isRemoved, LINEAGE, type Spec } from './text-lineage.ts';

/** The record fields this module writes. */
export const MERGED = Object.freeze({
  /** On an empty paragraph a merge left in a cell: true. */
  holder: 'mh',
  /** On a covered cell, per writer: the keys of the paragraphs it saw hidden there. */
  cover: 'cv.',
});

const NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
const NS_P = 'http://schemas.openxmlformats.org/presentationml/2006/main';
/** The depth the walk for table frames stops at (the reader caps a part at this depth too). */
const MAX_DEPTH = 256;
/** The most paragraphs a cover record keeps; a larger one (only a hostile room) is all seen. */
const MAX_COVER = 4_096;

const isElement = (n: Node | undefined): n is Element => n?.type === 'element';
const isA = (s: Spec | undefined, local: string): s is ReplicaElementSpec =>
  !!s && !('text' in s) && s.ns === NS_A && s.local === local;

/** The text of a run or a field (`a:r`, `a:fld`): its `a:t`, or '' when it has none. */
function runText(run: Element): string {
  const t = elements(run).find((c) => c.local === 't');
  return (t?.children ?? []).map((c) => (c.type === 'text' ? c.value : '')).join('');
}

/**
 * Whether a paragraph (`a:p`) holds text: a run or a field with text, or a line break, as
 * core's `tableCellHasText` reads a cell.
 */
export function holdsText(p: Element): boolean {
  return elements(p).some(
    (c) => c.local === 'br' || ((c.local === 'r' || c.local === 'fld') && runText(c).length > 0)
  );
}

/** Whether a paragraph spec (`a:p`) holds text, as `holdsText` reads an element. */
function specHoldsText(p: Spec): boolean {
  if ('text' in p) return false;
  return (p.children ?? []).some((c) => {
    if ('text' in c || !(c.local === 'br' || c.local === 'r' || c.local === 'fld')) return false;
    if (c.local === 'br') return true;
    const t = c.children?.find((x) => !('text' in x) && x.local === 't');
    return !!t && !('text' in t) && (t.children ?? []).some((x) => 'text' in x && x.text !== '');
  });
}

/** Every table frame (`p:graphicFrame`) under `root`, in groups and alternate content too. */
function framesOf<T extends { ns: string; local: string }>(
  root: T,
  kids: (t: T) => readonly T[]
): T[] {
  const out: T[] = [];
  const walk = (s: T, depth: number): void => {
    if (s.ns === NS_P && s.local === 'graphicFrame') {
      out.push(s);
      return;
    }
    if (depth < MAX_DEPTH) for (const c of kids(s)) walk(c, depth + 1);
  };
  walk(root, 0);
  return out;
}

/**
 * Each covered cell of the table of a frame, with the drawn cell over it, by the grid the editor
 * draws: as the cascade reads it (`resolveFrameTable` in core's `cascade/table.ts`), up to the
 * rows and columns it draws, a short row included.
 */
function coveredCells(frame: Element): { tc: Element; over: Element }[] {
  const tbl = path(frame, 'a:graphic', 'a:graphicData', 'a:tbl');
  const nC = Math.min(children(path(tbl, 'a:tblGrid'), 'a:gridCol').length, TABLE_MAX_COLUMNS);
  const cells = children(tbl, 'a:tr')
    .slice(0, TABLE_MAX_ROWS)
    .map((tr) => children(tr, 'a:tc'));
  if (nC === 0 || cells.length === 0) return [];
  const grid = tableGridOf(frame, cells, nC);
  const out: { tc: Element; over: Element }[] = [];
  cells.forEach((row, r) =>
    row.slice(0, nC).forEach((tc, c) => {
      const block = grid.blocks[grid.owner[r * nC + c]];
      const over = block && cells[block.row][block.col];
      if (over && (block.row !== r || block.col !== c)) out.push({ tc, over });
    })
  );
  return out;
}

/**
 * The text bodies of a cell. A cell holds one or none (21.1.3.16 tc), but two peers that type
 * into a cell with none each give it one, and core joins them (`cellAsSaved`).
 */
const bodiesOf = (tc: Element): Element[] =>
  elements(tc).filter((c) => c.ns === NS_A && c.local === 'txBody');

/** The paragraph keys that any writer saw hidden in a covered cell, or 'all'. */
function coverSeen(rec: unknown): ReadonlySet<string> | 'all' {
  const out = new Set<string>();
  if (!(rec instanceof Y.Map)) return out;
  for (const [field, list] of (rec as Y.Map<unknown>).entries()) {
    if (!field.startsWith(MERGED.cover) || !Array.isArray(list)) continue;
    if (list.length > MAX_COVER) return 'all';
    for (const k of list as unknown[]) if (typeof k === 'string') out.add(k);
  }
  return out;
}

/**
 * After a local write: each empty paragraph a merge left in a cell gets `mh`, and each covered
 * cell of a table frame the write changed records the paragraphs it holds (`cv.<client>`).
 */
export function markMergedCells(
  room: Room,
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
  const present = new Set<NodeId>();
  const emptied: { body: Element; old: Element }[] = [];
  const walk = (el: Element, depth: number): void => {
    const old = before.get(el.id);
    if (old === el) return;
    for (const c of el.children) present.add(c.id);
    if (el.ns === NS_A && el.local === 'txBody' && isElement(old)) {
      const kept = new Set(el.children.map((c) => c.id));
      const paragraphs = old.children.filter((c) => isElement(c) && c.local === 'p');
      if (paragraphs.length > 0 && paragraphs.every((c) => !kept.has(c.id)))
        emptied.push({ body: el, old });
    }
    if (depth < MAX_DEPTH) for (const c of el.children) if (isElement(c)) walk(c, depth + 1);
  };
  walk(root, 0);
  // A text body whose paragraphs this write all moved elsewhere (a cell merge) and that got
  // new empty paragraphs, because a text body holds at least one: each new one holds the place
  // of the moved ones.
  for (const { body, old } of emptied) {
    const moved = old.children.filter((c) => isElement(c) && c.local === 'p');
    if (!moved.every((c) => present.has(c.id))) continue;
    for (const c of body.children)
      if (isElement(c) && c.local === 'p' && !before.has(c.id) && !holdsText(c))
        record(c.id)?.set(MERGED.holder, true);
  }
  // What this writer leaves hidden in each covered cell of a table it changed.
  const field = `${MERGED.cover}${room.doc.clientID.toString(36)}`;
  // Only the subtrees this write changed: a changed frame has changed ancestors.
  const changed = (el: Element): Element[] => elements(el).filter((c) => before.get(c.id) !== c);
  for (const frame of framesOf(root, changed)) {
    for (const { tc } of coveredCells(frame)) {
      const rec = record(tc.id);
      const bodies = bodiesOf(tc);
      if (!rec || bodies.length === 0) continue;
      const seen = coverSeen(rec);
      if (seen === 'all') continue;
      const keys = bodies
        .flatMap((body) => elements(body).filter((p) => p.local === 'p'))
        .map((p) => keyOf.get(p.id))
        .filter((k): k is string => k !== undefined && !seen.has(k));
      if (keys.length === 0) continue;
      const own = rec.get(field);
      const mine = Array.isArray(own)
        ? (own as unknown[]).filter((k) => typeof k === 'string')
        : [];
      rec.set(field, [...mine, ...keys].slice(0, MAX_COVER + 1));
    }
  }
}

/** `map` without `key`: a write's `before` that makes the write list one node again. */
export function withoutKey<K, V>(map: ReadonlyMap<K, V>, key: K): Map<K, V> {
  const out = new Map(map);
  out.delete(key);
  return out;
}

/** The elements of `root` whose keys `keys` holds, in document order. */
export function elementsByKey(
  root: Element,
  keys: ReadonlySet<string> | undefined,
  keyOf: ReadonlyMap<NodeId, string>
): Element[] {
  const out: Element[] = [];
  if (!keys || keys.size === 0) return out;
  const walk = (el: Element, depth: number): void => {
    const key = keyOf.get(el.id);
    if (key !== undefined && keys.has(key)) out.push(el);
    if (depth < MAX_DEPTH) for (const c of elements(el)) walk(c, depth + 1);
  };
  walk(root, 0);
  return out;
}

/** An element view of a spec, for the grid of core, which reads elements. */
function viewOf(spec: ReplicaElementSpec, byKey: Map<string, ReplicaElementSpec>): Element {
  byKey.set(spec.key, spec);
  const kids: Node[] = (spec.children ?? []).map((c) =>
    'text' in c
      ? { type: 'text', id: c.key as NodeId, value: c.text }
      : viewOf(c as ReplicaElementSpec, byKey)
  );
  return {
    type: 'element',
    id: spec.key as NodeId,
    kind: 'generic',
    ns: spec.ns,
    local: spec.local,
    ...(spec.prefix ? { prefix: spec.prefix } : {}),
    nsDecls: [],
    attrs: spec.attrs ?? [],
    children: kids,
  } as Element;
}

/**
 * The paragraphs that go from the covered cells of one table frame to the drawn cells over
 * them, by the text body they go to, in reading order; and the text bodies they leave, with the
 * children that stay.
 */
function planFrame(
  room: Room,
  cameBack: ReadonlySet<string>,
  frame: ReplicaElementSpec,
  into: Map<string, Spec[]>,
  left: Map<string, Spec[]>
): void {
  const byKey = new Map<string, ReplicaElementSpec>();
  const view = viewOf(frame, byKey);
  // A structure command in a room marks the table, and its write records what it hid: any
  // other text in a covered cell of a marked table came there by a concurrent change.
  const marked = isStructureEdited(view);
  for (const { tc, over } of coveredCells(view)) {
    const toView = bodiesOf(over)[0];
    const to = toView && byKey.get(toView.id);
    if (!to) continue;
    const seen = coverSeen(room.nodes.get(tc.id));
    const unseen = (k: string): boolean => {
      if (seen === 'all' || seen.has(k)) return false;
      const rec = room.nodes.get(k);
      if (!(rec instanceof Y.Map)) return false;
      if (marked) return true;
      return typeof rec.get(LINEAGE.movedFrom) === 'string' || isRemoved(rec as Y.Map<unknown>);
    };
    const plans = bodiesOf(tc).flatMap((view) => {
      const from = byKey.get(view.id);
      if (!from) return [];
      const kids = from.children ?? [];
      const going = kids.filter(
        (k) => isA(k, 'p') && (cameBack.has(k.key) || (specHoldsText(k) && unseen(k.key)))
      );
      return [{ from, kids, going, paragraphs: kids.filter((k) => isA(k, 'p')).length }];
    });
    // A text body holds at least one paragraph: when every paragraph of the cell would go, the
    // last stays. A second body (two peers typed into a cell with no body) that keeps none
    // leaves the read while another body of the cell keeps one.
    if (plans.every((p) => p.going.length === p.paragraphs))
      [...plans]
        .reverse()
        .find((p) => p.going.length > 0)
        ?.going.pop();
    const emptied = new Set(
      plans.filter((p) => p.going.length === p.paragraphs).map((p) => p.from)
    );
    const cell = byKey.get(tc.id);
    if (cell && emptied.size > 0 && emptied.size < plans.length)
      left.set(
        cell.key,
        (cell.children ?? []).filter((k) => !emptied.has(k as ReplicaElementSpec))
      );
    for (const { from, kids, going } of plans) {
      if (going.length === 0) continue;
      const gone = new Set(going);
      left.set(
        from.key,
        kids.filter((k) => !gone.has(k))
      );
      const list = into.get(to.key);
      if (list) for (const k of going) list.push(k);
      else into.set(to.key, going);
    }
  }
}

/**
 * A part's spec with merged text shown by the rule above. `cameBack` holds the keys that the
 * rescue put back in this read. The key of each text body that the rule changed is added to
 * `changed`: the next local write of the part lists those bodies in the room as the store holds
 * them (`PackageMirror`), so the room and this store agree.
 */
export function withMergedText(
  room: Room,
  spec: ReplicaElementSpec,
  cameBack: ReadonlySet<string>,
  changed?: Set<string>
): ReplicaElementSpec {
  const into = new Map<string, Spec[]>();
  const left = new Map<string, Spec[]>();
  const frames = framesOf<ReplicaElementSpec>(spec, (s) =>
    (s.children ?? []).filter((c): c is ReplicaElementSpec => !('text' in c))
  );
  for (const frame of frames) planFrame(room, cameBack, frame, into, left);
  if (into.size === 0 && cameBack.size === 0) return spec;
  const placeholder = (k: Spec): boolean => {
    if (!isA(k, 'p') || specHoldsText(k)) return false;
    const rec = room.nodes.get(k.key);
    return rec instanceof Y.Map && rec.get(MERGED.holder) === true;
  };
  const walk = (s: ReplicaElementSpec, depth: number): ReplicaElementSpec => {
    let kids = left.get(s.key) ?? s.children ?? [];
    const extra = into.get(s.key);
    if (extra) kids = [...kids, ...extra];
    // Text the rescue put back shows in this text body: the empty paragraph that held its
    // place goes.
    if (s.local === 'txBody' && kids.some((k) => isA(k, 'p') && cameBack.has(k.key)))
      kids = kids.filter((k) => !placeholder(k));
    const next =
      depth < MAX_DEPTH
        ? kids.map((c) => ('text' in c ? c : walk(c as ReplicaElementSpec, depth + 1)))
        : kids;
    const same =
      next.length === (s.children ?? []).length && next.every((c, i) => c === s.children![i]);
    if (!same && s.local === 'txBody') changed?.add(s.key);
    return same ? s : { ...s, children: next };
  };
  return walk(spec, 0);
}
