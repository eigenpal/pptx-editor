// History groups, selection marks, and history-free origins of the document store.

import { describe, expect, test } from 'bun:test';
import { buildPptx, para, run, shape } from '../../../../../scripts/fixtures/builder.ts';
import {
  find,
  PptxDocumentStore,
  readPptx,
  textOf,
  xmlPartOf,
  type StoreChange,
  type TransactOptions,
} from '../index.ts';

const SLIDE = '/ppt/slides/slide1.xml';

function store(): PptxDocumentStore {
  const r = readPptx(
    buildPptx({
      slides: [{ shapes: shape({ id: 2, x: 0, y: 0, cx: 100, cy: 100, text: para(run('')) }) }],
    })
  );
  if (!r.ok) throw new Error('fixture');
  return new PptxDocumentStore(r.pkg);
}

const textNow = (s: PptxDocumentStore): string =>
  textOf(find(xmlPartOf(s.package, SLIDE)!.root, 'p:txBody'));

/** Set the slide's text in one transaction. */
function type(s: PptxDocumentStore, text: string, options: TransactOptions = {}) {
  const t = find(xmlPartOf(s.package, SLIDE)!.root, 'a:t')!;
  const r = s.transact([{ op: 'setText', part: SLIDE, node: t.id, text }], options);
  expect(r.ok).toBe(true);
  return r;
}

/** Undo until nothing is left; returns the number of steps. */
function undoAll(s: PptxDocumentStore): number {
  let steps = 0;
  while (s.undo().ok) steps += 1;
  return steps;
}

describe('history groups', () => {
  test('a typing burst with one group is one undo step', () => {
    const s = store();
    const historyGroup = Symbol('typing');
    for (const text of ['H', 'He', 'Hel', 'Hell', 'Hello']) type(s, text, { historyGroup });
    expect(undoAll(s)).toBe(1);
    expect(textNow(s)).toBe('');
  });

  test('a drag with one group is one undo step, and its change says merged', () => {
    const s = store();
    const changes: StoreChange[] = [];
    s.subscribe((c) => changes.push(c));
    const drag = Symbol('drag');
    type(s, 'a', { historyGroup: drag, label: 'transform' });
    type(s, 'b', { historyGroup: drag, label: 'transform' });
    type(s, 'c', { historyGroup: drag, label: 'transform' });
    expect(changes.map((c) => c.merged ?? false)).toEqual([false, true, true]);
    expect(undoAll(s)).toBe(1);
  });

  test('another group, or none, closes the open group', () => {
    const s = store();
    const first = Symbol('first');
    const second = Symbol('second');
    type(s, 'a', { historyGroup: first });
    type(s, 'ab', { historyGroup: first });
    type(s, 'abc', { historyGroup: second });
    type(s, 'abcd');
    // The first group does not reopen after another intent closed it.
    type(s, 'abcde', { historyGroup: first });
    expect(undoAll(s)).toBe(4);
  });

  test('a refused transaction of another intent still closes the group', () => {
    const s = store();
    const g = Symbol('g');
    type(s, 'a', { historyGroup: g });
    const bad = s.transact([{ op: 'setText', part: SLIDE, node: 'missing', text: 'x' }]);
    expect(bad.ok).toBe(false);
    type(s, 'ab', { historyGroup: g });
    expect(undoAll(s)).toBe(2);
  });

  test('undo and redo are boundaries; closeHistoryGroup is one too', () => {
    const s = store();
    const g = Symbol('g');
    type(s, 'a', { historyGroup: g });
    expect(s.undo().ok).toBe(true);
    expect(s.redo().ok).toBe(true);
    type(s, 'ab', { historyGroup: g });
    s.closeHistoryGroup();
    type(s, 'abc', { historyGroup: g });
    expect(undoAll(s)).toBe(3);
  });
});

describe('selection before and after', () => {
  test('undo restores selectionBefore and redo restores selectionAfter', () => {
    const s = store();
    const g = Symbol('typing');
    type(s, 'a', { historyGroup: g, selectionBefore: 'caret 0', selectionAfter: 'caret 1' });
    type(s, 'ab', { historyGroup: g, selectionBefore: 'ignored', selectionAfter: 'caret 2' });
    // A frame without a selection keeps the one before it.
    type(s, 'abc', { historyGroup: g });
    expect(s.selectionForUndo()).toBe('caret 0');
    const undone = s.undo();
    expect(undone).toEqual({ ok: true, selection: 'caret 0' });
    expect(s.selectionForRedo()).toBe('caret 2');
    expect(s.redo()).toEqual({ ok: true, selection: 'caret 2' });
    expect(textNow(s)).toBe('abc');
  });

  test('setSelectionAfter sets the newest entry only while it made the current package', () => {
    const s = store();
    type(s, 'a', { selectionBefore: 'before' });
    expect(s.setSelectionAfter('after')).toBe(true);
    expect(s.undo().selection).toBe('before');
    expect(s.setSelectionAfter('stale')).toBe(false);
    expect(s.redo().selection).toBe('after');
  });
});

describe('origins that record no history', () => {
  test('a projection commit publishes a revision but adds no undo step', () => {
    const s = store();
    type(s, 'a');
    const origins: string[] = [];
    s.subscribe((c) => origins.push(c.origin));
    const revision = s.revision;
    type(s, 'ab', { origin: 'projection' });
    expect(s.revision).toBe(revision + 1);
    expect(origins).toEqual(['projection']);
    expect(undoAll(s)).toBe(1);
  });

  test('awareness records nothing and keeps the redo stack', () => {
    const s = store();
    type(s, 'a');
    s.undo();
    type(s, 'b', { origin: 'awareness' });
    expect(s.canRedo).toBe(true);
    expect(s.canUndo).toBe(false);
  });

  test('a projection commit does not close the open group', () => {
    const s = store();
    const g = Symbol('g');
    type(s, 'a', { historyGroup: g });
    type(s, 'ab', { origin: 'projection' });
    type(s, 'abc', { historyGroup: g });
    expect(undoAll(s)).toBe(1);
  });

  test('recordsHistory overrides the origin default', () => {
    const s = store();
    type(s, 'a', { recordsHistory: false });
    expect(s.canUndo).toBe(false);
    type(s, 'ab', { origin: 'projection', recordsHistory: true });
    expect(undoAll(s)).toBe(1);
  });

  test('a remote commit still records nothing and clears both stacks', () => {
    const s = store();
    type(s, 'a');
    type(s, 'ab');
    s.undo();
    type(s, 'x', { origin: 'remote', recordsHistory: true });
    expect(s.canUndo).toBe(false);
    expect(s.canRedo).toBe(false);
  });

  test('a remote commit in the middle of a gesture keeps the next frame merged', () => {
    const s = store();
    const changes: StoreChange[] = [];
    s.subscribe((c) => changes.push(c));
    const drag = Symbol('drag');
    type(s, 'a', { historyGroup: drag });
    type(s, 'ab', { origin: 'remote' });
    type(s, 'abc', { historyGroup: drag });
    expect(changes.map((c) => c.merged ?? false)).toEqual([false, false, true]);
    // Another intent after the peer's edit closes the cut gesture, as it closes an open one.
    type(s, 'x', { origin: 'remote' });
    type(s, 'xy');
    type(s, 'xyz', { historyGroup: drag });
    expect(changes.slice(3).map((c) => c.merged ?? false)).toEqual([false, false, false]);
  });

  test('closeHistoryGroup also ends a gesture that a remote commit cut', () => {
    const s = store();
    const changes: StoreChange[] = [];
    s.subscribe((c) => changes.push(c));
    const drag = Symbol('drag');
    type(s, 'a', { historyGroup: drag });
    type(s, 'ab', { origin: 'remote' });
    s.closeHistoryGroup();
    type(s, 'abc', { historyGroup: drag });
    expect(changes[2].merged).toBeUndefined();
  });

  test('redo after a history-free commit keeps that commit', () => {
    const s = store();
    type(s, 'a');
    s.undo();
    const before = s.package;
    type(s, 'p', { origin: 'projection' });
    expect(s.package).not.toBe(before);
    expect(s.redo().ok).toBe(true);
    expect(textNow(s)).toBe('a');
    // Undo returns to the package the redo started from, with the projection in it.
    expect(s.undo().ok).toBe(true);
    expect(textNow(s)).toBe('p');
  });
});
