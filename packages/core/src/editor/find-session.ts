// The find state of one editor: which match is shown, and stepping, replacing, and replacing
// all from there. Matches come from `find.ts`; edits go through the editor's command path, so
// a replacement is refused, undone, and redone like any other command.

import type {
  CommandResult,
  EditorCommand,
  EditorSelection,
  FindMatch,
  FindOptions,
} from '../contracts/index.ts';
import type { DocumentSession } from './document-session.ts';
import { cleanReplacement, findMatches, textBodies } from '../automation/commands/find.ts';
import { findItem, parentGroupOf } from '../automation/commands/placement.ts';
import { addressOf, cellLayoutAt, tableItem } from './table-cells.ts';

/** Where a find step landed. `index` is -1 and `match` null when nothing matches. @public */
export type FindStep =
  | {
      readonly ok: true;
      readonly match: FindMatch | null;
      readonly index: number;
      readonly total: number;
    }
  | { readonly ok: false; readonly reason: string; readonly detail?: string };

/** The outcome of replace all: how many matches were replaced. @public */
export type ReplaceAllResult =
  | { readonly ok: true; readonly count: number }
  | { readonly ok: false; readonly reason: string; readonly detail?: string };

/**
 * Find and replace on an editor. Each call takes the query and options, so the find bar owns
 * them. A step selects its match: text on a slide or in a table cell as a text selection, and
 * speaker notes by going to the slide. @public
 */
export interface PptxFind {
  /** Every match of `query`, ordered by slide and position (notes after the slide's shapes). */
  matches(query: string, options?: FindOptions): readonly FindMatch[];
  /** The index of the match the editor shows now, or -1. */
  position(query: string, options?: FindOptions): number;
  /** Select the next match after the selection, wrapping at the end of the deck. */
  next(query: string, options?: FindOptions): FindStep;
  /** Select the match before the selection, wrapping at the start of the deck. */
  previous(query: string, options?: FindOptions): FindStep;
  /**
   * Replace the match the editor shows and select the next one. When the editor shows no
   * match, select the next match and replace nothing.
   */
  replace(query: string, replacement: string, options?: FindOptions): FindStep;
  /** Replace every match as one undo step. */
  replaceAll(query: string, replacement: string, options?: FindOptions): ReplaceAllResult;
}

/** What find needs from the editor. */
export interface FindHost {
  session(): DocumentSession | null;
  slide(): number;
  selection(): EditorSelection;
  /** Go to a slide and select there. */
  show(slide: number, selection: EditorSelection): void;
  exec(command: EditorCommand): CommandResult;
}

/** A deck position: text body index, paragraph, offset. */
type Position = readonly [number, number, number];

const compare = (a: Position, b: Position): number => a[0] - b[0] || a[1] - b[1] || a[2] - b[2];
const same = (a: FindMatch, b: FindMatch): boolean =>
  a.part === b.part &&
  a.txBodyId === b.txBodyId &&
  a.paragraph === b.paragraph &&
  a.start === b.start &&
  a.end === b.end;

/** The selection that shows a match. */
function selectionFor(m: FindMatch, session: DocumentSession): EditorSelection {
  if (m.where === 'notes') return { kind: 'none' };
  const items = session.layout(m.slide)?.items ?? [];
  const item = findItem(items, m.nodeId);
  if (item?.kind === 'shape' && item.element.text)
    return {
      kind: 'text',
      nodeId: m.nodeId,
      anchor: { paragraph: m.paragraph, offset: m.start },
      focus: { paragraph: m.paragraph, offset: m.end },
    };
  // A match in a table cell selects it in the cell.
  const cell =
    item?.kind === 'frame'
      ? item.table?.cells.find((c) => c.cell.text.nodeId === m.txBodyId)
      : undefined;
  if (cell)
    return {
      kind: 'text',
      nodeId: m.nodeId,
      cell: addressOf(cell),
      anchor: { paragraph: m.paragraph, offset: m.start },
      focus: { paragraph: m.paragraph, offset: m.end },
    };
  const group = parentGroupOf(items, m.nodeId);
  return { kind: 'shapes', nodeIds: [m.nodeId], ...(group ? { group } : {}) };
}

/** The listeners and the shown notes match of each find state, for the adapters. */
const watchers = new WeakMap<
  PptxFind,
  { readonly listeners: Set<() => void>; readonly notesMatch: () => FindMatch | null }
>();

/**
 * Run `listener` each time `find` shows a match or forgets one. A step inside the speaker notes
 * changes neither the slide nor the selection, so the find bar and the notes pane listen here
 * to draw again. Returns the function that stops. @internal
 */
export function onFindShown(find: PptxFind, listener: () => void): () => void {
  const w = watchers.get(find);
  if (!w) return () => {};
  w.listeners.add(listener);
  return () => w.listeners.delete(listener);
}

/**
 * The speaker notes match that `find` shows, while the editor still shows it: the same deck, the
 * same slide, and no selection on the slide. Null otherwise. @internal
 */
export function findNotesMatch(find: PptxFind): FindMatch | null {
  return watchers.get(find)?.notesMatch() ?? null;
}

/** Create the find state for one editor. */
export function createFind(host: FindHost): PptxFind {
  /** The match shown last, with the deck, slide, and selection that showed it. */
  let shown: {
    match: FindMatch;
    pkg: unknown;
    slide: number;
    selection: EditorSelection;
  } | null = null;
  const listeners = new Set<() => void>();
  const changed = (): void => {
    for (const listener of [...listeners]) listener();
  };
  let cache: { pkg: unknown; key: string; list: readonly FindMatch[] } | null = null;

  /** The text body of the cell a text selection is in. */
  const cellBodyId = (sel: Extract<EditorSelection, { kind: 'text' }>): string | undefined => {
    const frame = tableItem(host.session()?.layout(host.slide()), sel.nodeId);
    return frame && sel.cell ? cellLayoutAt(frame.table, sel.cell)?.cell.text.nodeId : undefined;
  };

  const matches = (query: string, options: FindOptions = {}): readonly FindMatch[] => {
    const session = host.session();
    if (!session) return [];
    const pkg = session.store.package;
    const key = JSON.stringify([query, !!options.matchCase, !!options.wholeWord]);
    if (cache?.pkg !== pkg || cache.key !== key)
      cache = { pkg, key, list: findMatches(pkg, query, options) };
    return cache.list;
  };

  /** The index of the shown match in `list`, while the editor still shows it. */
  const shownIndex = (list: readonly FindMatch[]): number => {
    const sel = host.selection();
    if (
      shown &&
      host.slide() === shown.slide &&
      JSON.stringify(sel) === JSON.stringify(shown.selection)
    ) {
      const at = list.findIndex((m) => same(m, shown!.match));
      if (at >= 0) return at;
    }
    if (sel.kind !== 'text') return -1;
    const [lo, hi] = [sel.anchor, sel.focus].sort(
      (a, b) => a.paragraph - b.paragraph || a.offset - b.offset
    );
    return list.findIndex(
      (m) =>
        m.slide === host.slide() &&
        m.where === 'slide' &&
        m.nodeId === sel.nodeId &&
        (!sel.cell || m.txBodyId === cellBodyId(sel)) &&
        m.paragraph === lo.paragraph &&
        m.start === lo.offset &&
        hi.paragraph === lo.paragraph &&
        m.end === hi.offset
    );
  };

  /** Where the selection starts and ends, as deck positions. */
  const range = (session: DocumentSession): [Position, Position] => {
    const bodies = textBodies(session.store.package);
    const slide = host.slide();
    const part = session.info.slides[slide]?.partName;
    const sel = host.selection();
    const id = sel.kind === 'text' ? sel.nodeId : sel.kind === 'shapes' ? sel.nodeIds[0] : null;
    const body = sel.kind === 'text' && sel.cell ? cellBodyId(sel) : undefined;
    const at = id
      ? bodies.findIndex(
          (b) => b.part === part && b.nodeId === id && (!body || b.txBodyId === body)
        )
      : -1;
    if (sel.kind === 'text' && at >= 0) {
      const a: Position = [at, sel.anchor.paragraph, sel.anchor.offset];
      const b: Position = [at, sel.focus.paragraph, sel.focus.offset];
      return compare(a, b) <= 0 ? [a, b] : [b, a];
    }
    let first = at >= 0 ? at : bodies.findIndex((b) => b.slide >= slide);
    if (first < 0) first = bodies.length;
    return [
      [first, 0, 0],
      [first, 0, 0],
    ];
  };

  const bodyIndex = (session: DocumentSession): ((m: FindMatch) => number) => {
    const index = new Map(
      textBodies(session.store.package).map((b, i) => [`${b.part}\u0000${b.txBodyId}`, i])
    );
    return (m) => index.get(`${m.part}\u0000${m.txBodyId}`) ?? -1;
  };

  const show = (list: readonly FindMatch[], index: number): FindStep => {
    const session = host.session();
    const m = list[index];
    if (!session || !m) return { ok: true, match: null, index: -1, total: list.length };
    host.show(m.slide, selectionFor(m, session));
    shown = {
      match: m,
      pkg: session.store.package,
      slide: host.slide(),
      selection: host.selection(),
    };
    changed();
    return { ok: true, match: m, index, total: list.length };
  };

  const step = (query: string, options: FindOptions | undefined, forward: boolean): FindStep => {
    const session = host.session();
    if (!session) return { ok: false, reason: 'disabled.noEditor' };
    const list = matches(query, options);
    if (!list.length) return { ok: true, match: null, index: -1, total: 0 };
    const current = shownIndex(list);
    if (current >= 0) return show(list, (current + (forward ? 1 : list.length - 1)) % list.length);
    const [lo, hi] = range(session);
    const body = bodyIndex(session);
    const key = (m: FindMatch): Position => [body(m), m.paragraph, m.start];
    if (forward) {
      const at = list.findIndex((m) => compare(key(m), hi) >= 0);
      return show(list, at < 0 ? 0 : at);
    }
    let at = -1;
    list.forEach((m, i) => {
      if (compare(key(m), lo) < 0) at = i;
    });
    return show(list, at < 0 ? list.length - 1 : at);
  };

  const find: PptxFind = {
    matches,
    position: (query, options) => shownIndex(matches(query, options)),
    next: (query, options) => step(query, options, true),
    previous: (query, options) => step(query, options, false),
    replace(query, replacement, options) {
      const session = host.session();
      if (!session) return { ok: false, reason: 'disabled.noEditor' };
      const list = matches(query, options);
      const current = shownIndex(list);
      if (current < 0) return step(query, options, true);
      const m = list[current];
      const r = host.exec({ type: 'replace', match: m, replacement });
      if (!r.ok) return r;
      shown = null;
      // Continue after the new text, so a replacement that holds the query is not found again.
      const after = matches(query, options);
      const body = bodyIndex(session);
      const from: Position = [body(m), m.paragraph, m.start + cleanReplacement(replacement).length];
      const at = after.findIndex((x) => compare([body(x), x.paragraph, x.start], from) >= 0);
      return show(after, at < 0 ? 0 : at);
    },
    replaceAll(query, replacement, options) {
      const count = matches(query, options).length;
      if (!count) return { ok: true, count: 0 };
      const command: EditorCommand = { type: 'replaceAll', query, replacement };
      const r = host.exec(options ? { ...command, options } : command);
      shown = null;
      changed();
      return r.ok ? { ok: true, count } : r;
    },
  };
  watchers.set(find, {
    listeners,
    notesMatch: () =>
      shown?.match.where === 'notes' &&
      shown.pkg === host.session()?.store.package &&
      shown.slide === host.slide() &&
      host.selection().kind === 'none'
        ? shown.match
        : null,
  });
  return find;
}
