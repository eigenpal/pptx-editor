/*
Copyright (c) 2026 EigenPal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// Concurrent typing into a table cell with no text body (AGENTS.md,
// "Collaboration verification"). A cell may have no `a:txBody` (ECMA-376 Part 1, 21.1.3.16 tc),
// and each peer that types into it gives it one. Both bodies stay in the room and in every
// store; the canvas and the save join them into one body (`cellAsSaved` in core). Every history
// runs in every client order, with undo and redo. On every replica the cell shows every peer's
// text in one order, the replicas hold the same package, and each save reopens to the same
// valid package with one text body in the cell.

import { describe, expect, test } from 'bun:test';
import type { PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import {
  canonicalPackageFingerprint,
  readPptx,
  textOf,
  type Element,
} from '@pptx-editor.dev/core/store';
import { packageViolations } from '../../../core/src/store/store/part-checks.ts';
import { buildPptx, cell, table } from '../../../../scripts/fixtures/builder.ts';
import { IN, all, kids, peers, root, sessionsOf } from './lineage-support.ts';

/** A 2 x 2 table whose cell (0, 1) has no text body. */
const deck = (): Uint8Array =>
  buildPptx({
    slides: [
      {
        layout: 'blank',
        shapes: table({
          id: 2,
          x: IN,
          y: IN,
          cx: 2 * IN,
          cy: 2 * 370_840,
          columns: [IN, IN],
          tblPr: '<a:tblPr firstRow="1" bandRow="1"/>',
          rows:
            `<a:tr h="370840">${cell('A')}<a:tc><a:tcPr/></a:tc></a:tr>` +
            `<a:tr h="370840">${cell('C')}${cell('D')}</a:tr>`,
        }),
      },
    ],
  });

/** The slide's graphic frame (PresentationML, so `all` does not find it). */
function frameOf(e: PptxEditorInstance): Element {
  const find = (el: Element): Element | undefined => {
    for (const c of kids(el)) {
      const hit = c.kind === 'graphicFrame' ? c : find(c);
      if (hit) return hit;
    }
    return undefined;
  };
  return find(root(e))!;
}

/** Put a caret in cell (`row`, `column`) of `e`'s table. */
function caret(e: PptxEditorInstance, row = 0, column = 1, paragraph = 0, offset = 0): void {
  e.setSelection({
    kind: 'text',
    nodeId: frameOf(e).id,
    cell: { row, column },
    anchor: { paragraph, offset },
    focus: { paragraph, offset },
  });
}

/** Type `text` into the cell with no text body. */
function type(e: PptxEditorInstance, text: string): void {
  caret(e);
  expect(e.input('text', text).ok).toBe(true);
}

/**
 * Split cell (`row`, `column`) of `e`'s table into two rows. The command gives the cell with no
 * text body an empty one.
 */
function split(e: PptxEditorInstance, row: number, column: number): void {
  caret(e, row, column);
  expect(e.exec({ type: 'splitCell', rows: 2, cols: 1 }).ok).toBe(true);
  expect(bodiesIn(root(e)).map(textOf)).toEqual(['']);
}

/** The paragraphs that `e` draws in cell (0, `column`). */
function drawnCell(e: PptxEditorInstance, column = 1): string[] {
  const item = e.session!.layout(0)!.items.find((i) => i.element.kind === 'graphicFrame')!;
  const t = item.element.kind === 'graphicFrame' ? item.element.table! : undefined;
  const c = t!.cells.find((x) => x.row === 0 && x.col === column)!;
  return c.text.paragraphs.map((p) => p.runs.map((r) => r.text).join(''));
}

/** The text bodies of cell (0, `column`) of a slide root. */
const bodiesIn = (slide: Element, column = 1): Element[] =>
  kids(kids(all(slide, 'tr')[0]!).filter((c) => c.local === 'tc')[column]!).filter(
    (c) => c.local === 'txBody'
  );

/**
 * Every session is `ready`, the replicas hold the same package and draw the same cell, and each
 * replica's save reopens to the same valid package, with one text body in the cell that holds
 * the paragraphs the canvas drew. Returns those paragraphs.
 */
function settled(editors: readonly PptxEditorInstance[], column = 1): string[] {
  for (const s of sessionsOf(editors)) expect(s.statusSnapshot().status).toBe('ready');
  const print = canonicalPackageFingerprint(editors[0]!.session!.store.package);
  const shown = drawnCell(editors[0]!, column);
  const saves = editors.map((e) => {
    expect(canonicalPackageFingerprint(e.session!.store.package)).toBe(print);
    expect(drawnCell(e, column)).toEqual(shown);
    const r = readPptx(e.save()!);
    if (!r.ok) throw new Error('the save does not open');
    return r.pkg;
  });
  const saved = canonicalPackageFingerprint(saves[0]!);
  for (const pkg of saves) expect(canonicalPackageFingerprint(pkg)).toBe(saved);
  expect(packageViolations(saves[0]!)).toEqual([]);
  const part = saves[0]!.parts.get(editors[0]!.session!.info.slides[0]!.partName)!;
  if (part.kind !== 'xml') throw new Error('the slide is not XML');
  const bodies = bodiesIn(part.xml.root, column);
  expect(bodies.length).toBeLessThanOrEqual(1);
  expect(
    bodies.flatMap((b) =>
      kids(b)
        .filter((p) => p.local === 'p')
        .map(textOf)
    )
  ).toEqual(bodies.length === 0 ? [] : shown);
  return shown;
}

/** Every order of `ids`. */
function orders(ids: readonly number[]): number[][] {
  if (ids.length <= 1) return [[...ids]];
  return ids.flatMap((id, i) =>
    orders([...ids.slice(0, i), ...ids.slice(i + 1)]).map((rest) => [id, ...rest])
  );
}

/** The texts of peers by ascending client id: the order Yjs gives concurrent inserts. */
const byClient = (ids: readonly number[], texts: readonly string[]): string[] =>
  ids
    .map((id, i) => [id, texts[i]!] as const)
    .sort(([a], [b]) => a - b)
    .map(([, t]) => t);

describe('concurrent typing into a table cell with no text body keeps every peer text', () => {
  for (const ids of orders([1_001, 2_002])) {
    const name = `client ids ${ids.join(', ')}`;

    test(`both texts show in one body, in client order (${name})`, async () => {
      const { editors, sync } = await peers(deck(), ids);
      const [a, b] = editors as [PptxEditorInstance, PptxEditorInstance];
      sync.pause();
      type(a, 'xx');
      type(b, 'yy');
      sync.flush();
      expect(settled(editors)).toEqual(byClient(ids, ['xx', 'yy']));
    });

    test(`each undo takes only its own text, and each redo brings it back (${name})`, async () => {
      const { editors, sync } = await peers(deck(), ids);
      const [a, b] = editors as [PptxEditorInstance, PptxEditorInstance];
      const both = byClient(ids, ['xx', 'yy']);
      sync.pause();
      type(a, 'xx');
      type(b, 'yy');
      sync.flush();
      for (const [e, other] of [
        [a, 'yy'],
        [b, 'xx'],
      ] as const) {
        expect(e.exec({ type: 'undo' }).ok).toBe(true);
        sync.flush();
        expect(settled(editors)).toEqual([other]);
        expect(e.exec({ type: 'redo' }).ok).toBe(true);
        sync.flush();
        expect(settled(editors)).toEqual(both);
      }
      // Both undone at once: the cell has no text body again.
      expect(a.exec({ type: 'undo' }).ok).toBe(true);
      expect(b.exec({ type: 'undo' }).ok).toBe(true);
      sync.flush();
      expect(settled(editors)).toEqual([]);
      expect(bodiesIn(root(a))).toEqual([]);
    });

    test(`typing after the flush joins the bodies on every replica (${name})`, async () => {
      const { editors, sync } = await peers(deck(), ids);
      const [a, b] = editors as [PptxEditorInstance, PptxEditorInstance];
      const both = byClient(ids, ['xx', 'yy']);
      sync.pause();
      type(a, 'xx');
      type(b, 'yy');
      sync.flush();
      // Both peers type at the end of the cell at once: each edit joins the bodies first.
      sync.pause();
      caret(a, 0, 1, 1, 2);
      expect(a.input('text', 'p').ok).toBe(true);
      caret(b, 0, 1, 0, 0);
      expect(b.input('text', 'q').ok).toBe(true);
      sync.flush();
      expect(settled(editors)).toEqual([`q${both[0]}`, `${both[1]}p`]);
      expect(bodiesIn(root(a))).toHaveLength(1);
    });

    test(`a blank body that a peer's table command adds gives way to the typed text (${name})`, async () => {
      const { editors, sync } = await peers(deck(), ids);
      const [a, b] = editors as [PptxEditorInstance, PptxEditorInstance];
      sync.pause();
      type(a, 'xx');
      // A split of another cell gives every cell of the table with no text body an empty one.
      split(b, 1, 0);
      sync.flush();
      expect(bodiesIn(root(a))).toHaveLength(2);
      expect(settled(editors)).toEqual(['xx']);
      expect(b.exec({ type: 'undo' }).ok).toBe(true);
      sync.flush();
      expect(settled(editors)).toEqual(['xx']);
    });

    test(`text typed into the cell while a peer merges it into its neighbor shows in the merged cell (${name})`, async () => {
      const { editors, sync } = await peers(deck(), ids);
      const [a, b] = editors as [PptxEditorInstance, PptxEditorInstance];
      sync.pause();
      a.setSelection({
        kind: 'cells',
        nodeId: frameOf(a).id,
        anchor: { row: 0, column: 0 },
        focus: { row: 0, column: 1 },
      });
      expect(a.exec({ type: 'mergeCells' }).ok).toBe(true);
      type(b, 'yy');
      sync.flush();
      expect(settled(editors, 0)).toEqual(['A', 'yy']);
      expect(a.exec({ type: 'undo' }).ok).toBe(true);
      sync.flush();
      expect(settled(editors, 0)).toEqual(['A']);
      expect(settled(editors)).toEqual(['yy']);
    });

    test(`edits after the merge keep the typed text when the merge is undone (${name})`, async () => {
      const { editors, sync } = await peers(deck(), ids);
      const [a, b] = editors as [PptxEditorInstance, PptxEditorInstance];
      sync.pause();
      a.setSelection({
        kind: 'cells',
        nodeId: frameOf(a).id,
        anchor: { row: 0, column: 0 },
        focus: { row: 0, column: 1 },
      });
      expect(a.exec({ type: 'mergeCells' }).ok).toBe(true);
      type(b, 'yy');
      sync.flush();
      // Each peer writes the part once more: the merged cell lists the typed text as its stores
      // hold it.
      caret(a, 0, 0, 0, 1);
      expect(a.input('text', 'p').ok).toBe(true);
      sync.flush();
      caret(b, 1, 1, 0, 1);
      expect(b.input('text', 'q').ok).toBe(true);
      sync.flush();
      expect(settled(editors, 0)).toEqual(['Ap', 'yy']);
      expect(a.exec({ type: 'undo' }).ok).toBe(true);
      expect(a.exec({ type: 'undo' }).ok).toBe(true);
      sync.flush();
      const text = [...settled(editors, 0), ...settled(editors)].join('');
      expect(text).toContain('A');
      expect(text).toContain('yy');
      expect(text).not.toContain('p');
    });

    test(`two peers' table commands leave one empty paragraph in the cell (${name})`, async () => {
      const { editors, sync } = await peers(deck(), ids);
      const [a, b] = editors as [PptxEditorInstance, PptxEditorInstance];
      sync.pause();
      split(a, 1, 0);
      split(b, 1, 1);
      sync.flush();
      expect(bodiesIn(root(a))).toHaveLength(2);
      expect(settled(editors)).toEqual(['']);
    });
  }

  for (const ids of orders([1_001, 2_002, 3_003])) {
    const name = `client ids ${ids.join(', ')}`;

    test(`three peers' texts show in client order, in every delivery order (${name})`, async () => {
      for (const reversed of [false, true]) {
        const { editors, sync, docs } = await peers(deck(), ids);
        sync.pause();
        const typed = ['xx', 'yy', 'zz'];
        const turn = reversed ? [2, 1, 0] : [0, 1, 2];
        for (const i of turn) type(editors[i]!, typed[i]!);
        // One peer gets the others' updates before the third peer does.
        sync.deliverTo(docs[turn[0]!]!);
        sync.flush();
        expect(settled(editors)).toEqual(byClient(ids, typed));
      }
    });

    test(`undo and redo of one of three peers keeps the others' text (${name})`, async () => {
      const { editors, sync } = await peers(deck(), ids);
      sync.pause();
      const typed = ['xx', 'yy', 'zz'];
      editors.forEach((e, i) => type(e, typed[i]!));
      sync.flush();
      const all3 = byClient(ids, typed);
      for (const [i, e] of editors.entries()) {
        expect(e.exec({ type: 'undo' }).ok).toBe(true);
        sync.flush();
        expect(settled(editors)).toEqual(all3.filter((t) => t !== typed[i]));
        expect(e.exec({ type: 'redo' }).ok).toBe(true);
        sync.flush();
        expect(settled(editors)).toEqual(all3);
      }
    });
  }
});
