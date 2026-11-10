/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Where a peer's text goes when an undo takes out the table cell or the table row that holds
// it. An undo of a row insert, a column insert, or a cell split unlists the cells that the step
// made (`undo-unlisted.ts`). A peer who typed into such a cell at the same time did not remove
// that text, so the text must stay. A cell or a row cannot show again with a part of its
// content: a row needs a cell for each grid column, and a cell a text body with the undoing
// author's properties (ECMA-376 Part 1, 21.1.3.18 tr and 21.1.3.16 tc). So each paragraph that
// holds a peer's content goes to the end of the text body of a cell that stays:
//
// - From a cell of a row: the cell that it followed in the row, else the cell before which it
//   was listed, else the first cell of the row with a text body.
// - From a cell of a row that the undo took: the cell in the same column of the row that the
//   row followed, else of the row before which it was listed, else of the first row.
//
// The paragraph is not listed there: it shows while it holds a peer's content, as any node that
// an undo unlisted (`UndoneListings.shown`). A redo lists the cell again, and the paragraph
// shows in it. A cell that a merge covers shows the paragraph in the drawn cell over it
// (`merged-cells.ts`), so the caller passes the keys it placed (`placed`) to that rule.
//
// Everything read here comes from the room and is untrusted: each walk stops at `MAX_KEYS`
// keys, and a key that is read twice shows once.

import type { ReplicaElementSpec, ReplicaTextSpec } from '@pptx-editor.dev/core/store';

type Spec = ReplicaElementSpec | ReplicaTextSpec;

const NS_A = 'http://schemas.openxmlformats.org/drawingml/2006/main';
/** The most child keys one walk looks at; only a hostile room lists more. */
const MAX_KEYS = 4_096;

/** An undone listing, as `undo-unlisted.ts` reads it. */
export interface RehomeEntry {
  readonly key: string;
  readonly after: string;
  readonly before: string;
}

/** What the rule reads from the room, through the undone listings of one part read. */
export interface RehomeSource<E extends RehomeEntry> {
  /** The local name of the DrawingML element that `key` records, or undefined. */
  localOf(key: string): string | undefined;
  /** The keys a list holds under `key`, then the undone keys that may show there. */
  childrenOf(key: string): readonly string[];
  /** Whether `key` holds content that the undo of `e` did not take. */
  holdsForeign(key: string, e: E): boolean;
  /** The paragraph `key` as the read shows it, or null when it does not show. */
  readParagraph(key: string): Spec | null;
}

const isA = (s: Spec | undefined, local: string): s is ReplicaElementSpec =>
  !!s && !('text' in s) && s.ns === NS_A && s.local === local;

/** The paragraphs under the cell `cell` that hold content the undo of `e` did not take. */
function foreignParagraphs<E extends RehomeEntry>(
  src: RehomeSource<E>,
  cell: string,
  e: E
): Spec[] {
  const out: Spec[] = [];
  let looked = 0;
  for (const body of src.childrenOf(cell)) {
    if (src.localOf(body) !== 'txBody') continue;
    for (const p of src.childrenOf(body)) {
      if ((looked += 1) > MAX_KEYS) return out;
      if (src.localOf(p) !== 'p' || !src.holdsForeign(p, e)) continue;
      const spec = src.readParagraph(p);
      if (spec) out.push(spec);
    }
  }
  return out;
}

/**
 * The index in `out` of the sibling that `e` followed (through `after`, the anchors of the
 * undone siblings), else the one before which it was listed, else -1.
 */
function anchorIndex(out: readonly Spec[], e: RehomeEntry, after: ReadonlyMap<string, string>) {
  const at = (k: string) => out.findIndex((c) => c.key === k);
  const visited = new Set([e.key]);
  for (let a = e.after; a !== '' && !visited.has(a); a = after.get(a) ?? '') {
    visited.add(a);
    const i = at(a);
    if (i >= 0) return i;
  }
  return e.before === '' ? -1 : at(e.before);
}

/**
 * `cell` with `paragraphs` at the end of its first text body, or null when it has none. Notes
 * the text body of each paragraph in `placed`.
 */
function withParagraphs(
  cell: ReplicaElementSpec,
  paragraphs: readonly Spec[],
  placed: Map<string, string>
) {
  const kids = cell.children ?? [];
  const i = kids.findIndex((k) => isA(k, 'txBody'));
  if (i < 0) return null;
  const body = kids[i] as ReplicaElementSpec;
  const next = [...kids];
  next[i] = { ...body, children: [...(body.children ?? []), ...paragraphs] };
  for (const p of paragraphs) placed.set(p.key, body.key);
  return { ...cell, children: next };
}

/** `row` with `paragraphs` in its cell at `column`, or the first cell after or before it. */
function intoRow(
  row: ReplicaElementSpec,
  column: number,
  paragraphs: readonly Spec[],
  placed: Map<string, string>
) {
  const kids = row.children ?? [];
  const cells = kids.flatMap((k, i) => (isA(k, 'tc') ? [i] : []));
  if (cells.length === 0) return null;
  const from = Math.max(0, Math.min(column, cells.length - 1));
  const order = [...cells.slice(from), ...cells.slice(0, from).reverse()];
  for (const i of order) {
    const cell = withParagraphs(kids[i] as ReplicaElementSpec, paragraphs, placed);
    if (!cell) continue;
    const next = [...kids];
    next[i] = cell;
    return { ...row, children: next };
  }
  return null;
}

/**
 * `out` (the children of a table row or a table, as read) with the paragraphs of each undone
 * cell or row in `orphans` that hold a peer's content, placed by the rule above. `after` maps
 * each undone sibling to the key it followed. Maps each placed paragraph to its text body in
 * `placed`.
 */
export function rehomeUndone<E extends RehomeEntry>(
  out: readonly Spec[],
  orphans: readonly E[],
  after: ReadonlyMap<string, string>,
  src: RehomeSource<E>,
  placed: Map<string, string>
): Spec[] {
  const next = [...out];
  for (const e of orphans) {
    const local = src.localOf(e.key);
    if (local === 'tc') {
      const ps = foreignParagraphs(src, e.key, e);
      if (ps.length === 0) continue;
      const cells = next.flatMap((k, i) => (isA(k, 'tc') ? [i] : []));
      const anchor = anchorIndex(next, e, after);
      const order = anchor >= 0 ? [anchor, ...cells.filter((i) => i !== anchor)] : cells;
      for (const i of order) {
        const cell = withParagraphs(next[i] as ReplicaElementSpec, ps, placed);
        if (!cell) continue;
        next[i] = cell;
        break;
      }
    } else if (local === 'tr') {
      const rows = next.flatMap((k, i) => (isA(k, 'tr') ? [i] : []));
      const anchor = anchorIndex(next, e, after);
      const target = anchor >= 0 && isA(next[anchor], 'tr') ? anchor : rows[0];
      if (target === undefined) continue;
      const cells = src.childrenOf(e.key).filter((k) => src.localOf(k) === 'tc');
      cells.forEach((cell, column) => {
        const ps = foreignParagraphs(src, cell, e);
        if (ps.length === 0) return;
        const row = intoRow(next[target] as ReplicaElementSpec, column, ps, placed);
        if (row) next[target] = row;
      });
    }
  }
  return next;
}
