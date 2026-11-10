/*
Copyright (c) 2026 Eigenpal, Inc. All rights reserved.
Licensed under the EigenPal Pro Evaluation License 1.0. See packages/pro/LICENSE.md.
Production use requires a commercial agreement: licensing@eigenpal.com
*/
// An undo keeps what peers wrote (`undo-delete-filter.ts`, `undo-unlisted.ts`, `undo-splits.ts`):
// two editors on two linked Y.Docs, in both client-id orders. Each case compares the text or
// the XML, not only the fingerprints, because the replicas converge on a wrong result too.

import { describe, expect, test } from 'bun:test';
import * as Y from 'yjs';
import type { PptxEditorInstance } from '@pptx-editor.dev/core/editor';
import {
  canonicalPackageFingerprint,
  readPptx,
  serializeXmlPart,
  xmlPartOf,
} from '@pptx-editor.dev/core/store';
import { buildPptx, cell, table } from '../../../../scripts/fixtures/builder.ts';
import { ATTR_SEP, F, NODES_KEY } from '../collaboration/yjs/schema.ts';
import { UNDONE_KEY } from '../collaboration/yjs/undo-unlisted.ts';
import {
  converged,
  formatOf,
  IN,
  inlineOf,
  oneRun,
  paragraphs,
  peers,
  replicas,
  selectIn,
  sessionsOf,
  slide1,
  twoParagraphs,
} from './lineage-support.ts';

const ORDERS = [false, true] as const;
const order = (swap: boolean) => `client order ${swap ? 'B, A' : 'A, B'}`;

const xml = (e: PptxEditorInstance) =>
  serializeXmlPart(xmlPartOf(e.session!.store.package, slide1(e))!);
const spPr = (e: PptxEditorInstance) => /<p:spPr>[\s\S]*?<\/p:spPr>/.exec(xml(e))![0];
const pPr = (e: PptxEditorInstance) =>
  /<a:pPr\b[^>]*?(\/>|>[\s\S]*?<\/a:pPr>)/.exec(xml(e))?.[0] ?? '';
const selectShape = (e: PptxEditorInstance) =>
  e.setSelection({ kind: 'shapes', nodeIds: [e.session!.layout(0)!.items[0]!.element.nodeId] });
const ok = (result: { ok: boolean }) => expect(result.ok).toBe(true);

/** Both saved decks reopen to one package, and a second save keeps it. */
function savedAlike(a: PptxEditorInstance, b: PptxEditorInstance): void {
  const reopened = [a, b].map((e) => {
    const r = readPptx(e.save()!);
    if (!r.ok) throw new Error('reopen failed');
    return r.pkg;
  });
  expect(canonicalPackageFingerprint(reopened[1]!)).toBe(canonicalPackageFingerprint(reopened[0]!));
  const again = readPptx(a.save()!);
  if (!again.ok) throw new Error('reopen failed');
  expect(canonicalPackageFingerprint(again.pkg)).toBe(canonicalPackageFingerprint(reopened[0]!));
}

/** Converged text, and saves that reopen alike. */
function settled(a: PptxEditorInstance, b: PptxEditorInstance): string[] {
  const text = converged(a, b);
  savedAlike(a, b);
  return text;
}

describe('an undo of a format keeps the text a peer typed into its run', () => {
  for (const swap of ORDERS) {
    test(`typed before the undo (${order(swap)})`, async () => {
      const { a, b } = await replicas(oneRun('Hello world'), swap);
      selectIn(b, 0, 6, 11);
      ok(b.exec({ type: 'toggleItalic' }));
      selectIn(a, 0, 11, 11);
      ok(a.input('text', '!'));
      ok(b.exec({ type: 'undo' }));
      expect(settled(a, b)).toEqual(['Hello world!']);
      expect(formatOf(a, 'world')).toEqual({});
      // B's redo brings the italic back over the word, and the typed text stays once.
      ok(b.exec({ type: 'redo' }));
      expect(settled(a, b)).toEqual(['Hello world!']);
      expect(formatOf(a, 'world!')).toEqual({ i: '1' });
      // A's undo takes back only what A typed.
      ok(a.exec({ type: 'undo' }));
      expect(settled(a, b)).toEqual(['Hello world']);
      ok(a.exec({ type: 'redo' }));
      expect(settled(a, b)).toEqual(['Hello world!']);
    });

    test(`a write of the paragraph does not keep the run after its text goes (${order(swap)})`, async () => {
      const { a, b } = await replicas(oneRun('Hello world'), swap);
      selectIn(b, 0, 6, 11);
      ok(b.exec({ type: 'toggleItalic' }));
      selectIn(a, 0, 11, 11);
      ok(a.input('text', '!'));
      ok(b.exec({ type: 'undo' }));
      // Both peers write the paragraph while the typed text shows in its own run.
      selectIn(b, 0, 0, 0);
      ok(b.input('text', '<'));
      selectIn(a, 0, 1, 1);
      ok(a.input('text', '>'));
      expect(settled(a, b)).toEqual(['<>Hello world!']);
      // The peer that typed it deletes it: no empty run is left.
      selectIn(a, 0, 13, 14);
      ok(a.input('deleteBackward'));
      expect(settled(a, b)).toEqual(['<>Hello world']);
      expect(inlineOf(a, 'Hello')).toEqual(['<>Hello world']);
      ok(a.exec({ type: 'undo' }));
      expect(settled(a, b)).toEqual(['<>Hello world!']);
      // A's undo of its first typing takes the text back, and with it the run.
      ok(a.exec({ type: 'undo' }));
      ok(a.exec({ type: 'undo' }));
      expect(settled(a, b)).toEqual(['<Hello world']);
      expect(xml(a)).not.toMatch(/<a:r\/>|<a:r>(?!<a:rPr|<a:t)/);
    });

    test(`text the undoing peer types into the run later stays (${order(swap)})`, async () => {
      const { a, b } = await replicas(oneRun('Hello world'), swap);
      selectIn(b, 0, 6, 11);
      ok(b.exec({ type: 'toggleItalic' }));
      selectIn(a, 0, 11, 11);
      ok(a.input('text', '!'));
      ok(b.exec({ type: 'undo' }));
      selectIn(b, 0, 12, 12);
      ok(b.input('text', '?'));
      expect(settled(a, b)).toEqual(['Hello world!?']);
      ok(a.exec({ type: 'undo' }));
      expect(settled(a, b)).toEqual(['Hello world?']);
      ok(b.exec({ type: 'undo' }));
      expect(settled(a, b)).toEqual(['Hello world']);
    });

    test(`typed while the undo is in flight (${order(swap)})`, async () => {
      const { a, b, sync } = await replicas(oneRun('Hello world'), swap);
      selectIn(b, 0, 6, 11);
      ok(b.exec({ type: 'toggleItalic' }));
      sync.pause();
      selectIn(a, 0, 8, 8);
      ok(a.input('text', '-'));
      ok(b.exec({ type: 'undo' }));
      sync.flush();
      const [text] = settled(a, b);
      // The typed text stays; it shows after the word, in a run of its own.
      expect(text!.replace('-', '')).toBe('Hello world');
      expect(text).toContain('-');
      ok(b.exec({ type: 'redo' }));
      expect(settled(a, b)[0]!.replace('-', '')).toBe('Hello world');
    });

    test(`a delete of the shape while the undo is in flight (${order(swap)})`, async () => {
      const { a, b, sync } = await replicas(oneRun('Hello world'), swap);
      selectIn(b, 0, 6, 11);
      ok(b.exec({ type: 'toggleItalic' }));
      selectIn(a, 0, 11, 11);
      ok(a.input('text', '!'));
      sync.pause();
      selectShape(a);
      ok(a.exec({ type: 'delete' }));
      ok(b.exec({ type: 'undo' }));
      sync.flush();
      expect(settled(a, b)).toEqual([]);
      ok(a.exec({ type: 'undo' }));
      expect(settled(a, b)).toEqual(['Hello world!']);
    });
  }
});

/** A 2 x 2 table: A, B above, C, D below. */
const twoByTwo = () =>
  buildPptx({
    slides: [
      {
        layout: 'blank',
        shapes: table({
          id: 2,
          x: IN,
          y: IN,
          cx: 4 * IN,
          cy: 2 * IN,
          columns: [2 * IN, 2 * IN],
          tblPr: '<a:tblPr firstRow="1"/>',
          rows: `<a:tr h="${IN}">${cell('A')}${cell('B')}</a:tr><a:tr h="${IN}">${cell('C')}${cell('D')}</a:tr>`,
        }),
      },
    ],
  });
const caretInCell = (e: PptxEditorInstance, row: number, column: number) =>
  e.setSelection({
    kind: 'text',
    nodeId: e.session!.layout(0)!.items[0]!.element.nodeId,
    cell: { row, column },
    anchor: { paragraph: 0, offset: 0 },
    focus: { paragraph: 0, offset: 0 },
  });
/** The cells of each saved row: one per grid column, so the table opens without repair. */
const savedRows = (e: PptxEditorInstance): number[] => {
  const r = readPptx(e.save()!);
  if (!r.ok) throw new Error('reopen failed');
  const saved = serializeXmlPart(xmlPartOf(r.pkg, slide1(e))!);
  return [...saved.matchAll(/<a:tr\b[\s\S]*?<\/a:tr>/g)].map(
    (m) => [...m[0].matchAll(/<a:tc[ >]/g)].length
  );
};

describe("an undo of an element's creation keeps a peer's later edit in it", () => {
  for (const swap of ORDERS) {
    test(`text typed into a new paragraph (${order(swap)})`, async () => {
      const { a, b } = await replicas(oneRun('Hello'), swap);
      selectIn(a, 0, 5, 5);
      ok(a.input('paragraph'));
      selectIn(b, 1, 0, 0);
      ok(b.input('text', 'Z'));
      ok(a.exec({ type: 'undo' }));
      // The paragraph stays with what B typed in it.
      expect(settled(a, b)).toEqual(['Hello', 'Z']);
      ok(b.exec({ type: 'undo' }));
      expect(settled(a, b)).toEqual(['Hello']);
      ok(b.exec({ type: 'redo' }));
      ok(a.exec({ type: 'redo' }));
      expect(settled(a, b)).toEqual(['Hello', 'Z']);
    });

    test(`a new table row that a peer typed into stays valid (${order(swap)})`, async () => {
      const { a, b } = await replicas(twoByTwo(), swap);
      caretInCell(a, 0, 0);
      ok(a.exec({ type: 'insertRow', where: 'below' }));
      caretInCell(b, 1, 1);
      ok(b.input('text', 'Z'));
      ok(a.exec({ type: 'undo' }));
      // A row is not valid without what the undo took (its height, each cell's body), so it
      // does not show again. B's paragraph shows in the cell of the same column in the row
      // above (`undo-rehome.ts`); A's redo brings the row back with B's text.
      expect(settled(a, b)).toEqual(['A', 'B', 'Z', 'C', 'D']);
      expect(savedRows(a)).toEqual([2, 2]);
      ok(a.exec({ type: 'redo' }));
      expect(settled(a, b)).toEqual(['A', 'B', '', 'Z', 'C', 'D']);
      expect(savedRows(a)).toEqual([2, 2, 2]);
    });

    test(`a line width in a new a:ln (${order(swap)})`, async () => {
      const { a, b } = await replicas(oneRun('Hello'), swap);
      selectShape(a);
      selectShape(b);
      ok(a.exec({ type: 'setLineColor', color: '#C00000' }));
      ok(b.exec({ type: 'setLineWidth', points: 6 }));
      ok(a.exec({ type: 'undo' }));
      settled(a, b);
      expect(spPr(a)).toMatch(/<a:ln\b[^>]*w="76200"/);
      expect(spPr(a)).not.toContain('C00000');
      // A's redo brings the color back beside B's width.
      ok(a.exec({ type: 'redo' }));
      settled(a, b);
      expect(spPr(a)).toMatch(/<a:ln\b[^>]*w="76200"/);
      expect(spPr(a)).toContain('C00000');
      ok(a.exec({ type: 'undo' }));
      // B's undo takes its width back too: no line is left.
      ok(b.exec({ type: 'undo' }));
      settled(a, b);
      expect(spPr(a)).not.toContain('<a:ln');
      ok(b.exec({ type: 'redo' }));
      settled(a, b);
      expect(spPr(a)).toMatch(/<a:ln\b[^>]*w="76200"/);
    });

    test(`a width written while the undo is in flight (${order(swap)})`, async () => {
      const { a, b, sync } = await replicas(oneRun('Hello'), swap);
      selectShape(a);
      selectShape(b);
      ok(a.exec({ type: 'setLineColor', color: '#C00000' }));
      sync.pause();
      ok(b.exec({ type: 'setLineWidth', points: 6 }));
      ok(a.exec({ type: 'undo' }));
      sync.flush();
      settled(a, b);
      expect(spPr(a)).toMatch(/<a:ln\b[^>]*w="76200"/);
      expect(spPr(a)).not.toContain('C00000');
    });

    test(`bullets in a new a:pPr (${order(swap)})`, async () => {
      const { a, b } = await replicas(oneRun('Hello'), swap);
      selectIn(a, 0, 0, 5);
      selectIn(b, 0, 0, 5);
      ok(a.exec({ type: 'setParagraphFormat', alignment: 'center' }));
      ok(b.exec({ type: 'toggleBullets' }));
      ok(a.exec({ type: 'undo' }));
      settled(a, b);
      expect(pPr(a)).toContain('<a:buChar');
      expect(pPr(a)).not.toContain('algn=');
      ok(a.exec({ type: 'redo' }));
      settled(a, b);
      expect(pPr(a)).toContain('<a:buChar');
      expect(pPr(a)).toContain('algn="ctr"');
    });

    test(`a delete of the line while the undo is in flight (${order(swap)})`, async () => {
      const { a, b, sync } = await replicas(oneRun('Hello'), swap);
      selectShape(a);
      selectShape(b);
      ok(a.exec({ type: 'setLineColor', color: '#C00000' }));
      ok(b.exec({ type: 'setLineWidth', points: 6 }));
      sync.pause();
      ok(b.exec({ type: 'delete' }));
      ok(a.exec({ type: 'undo' }));
      sync.flush();
      expect(settled(a, b)).toEqual([]);
      expect(xml(a)).not.toContain('<a:ln');
    });
  }
});

/** Two peers format one run at once; returns the replicas after the merge. */
async function concurrentFormats(
  text: string,
  [aFrom, aTo]: readonly [number, number],
  [bFrom, bTo]: readonly [number, number],
  swap: boolean
) {
  const r = await replicas(oneRun(text), swap);
  selectIn(r.a, 0, aFrom, aTo);
  selectIn(r.b, 0, bFrom, bTo);
  r.sync.pause();
  ok(r.a.exec({ type: 'toggleBold' }));
  ok(r.b.exec({ type: 'toggleItalic' }));
  r.sync.flush();
  return r;
}

describe('an undo of concurrent formats of one run shows the text once', () => {
  // [name, text, a's range, b's range]. Peer a cuts first or at the same place, and the tie
  // goes to the smaller client id, so the winner depends on the order only in the last case.
  const CASES = [
    ['overlapping ranges', 'Hello', [1, 3], [2, 4]],
    ['disjoint ranges', 'Hello big world', [6, 9], [10, 15]],
    ['one word', 'Hello world', [6, 11], [6, 11]],
  ] as const;
  for (const swap of ORDERS)
    for (const [name, text, ra, rb] of CASES) {
      // The loser is b, except for "one word" with b's smaller client id.
      const loserIsB = !(name === 'one word' && swap);
      test(`${name}: the loser undoes and redoes (${order(swap)})`, async () => {
        const { a, b } = await concurrentFormats(text, ra, rb, swap);
        const loser = loserIsB ? b : a;
        const winner = loserIsB ? a : b;
        const merged = settled(a, b);
        expect(merged).toEqual([text]);
        const shown = inlineOf(a, text.slice(0, 1));
        ok(loser.exec({ type: 'undo' }));
        expect(settled(a, b)).toEqual([text]);
        expect(inlineOf(a, text.slice(0, 1))).toEqual(shown);
        ok(loser.exec({ type: 'redo' }));
        expect(settled(a, b)).toEqual([text]);
        ok(loser.exec({ type: 'undo' }));
        ok(winner.exec({ type: 'undo' }));
        expect(settled(a, b)).toEqual([text]);
        expect(inlineOf(a, text.slice(0, 1))).toEqual([text]);
        ok(winner.exec({ type: 'redo' }));
        expect(settled(a, b)).toEqual([text]);
      });

      for (const winnerFirst of [true, false])
        test(`${name}: both undo, the ${winnerFirst ? 'winner' : 'loser'} first (${order(swap)})`, async () => {
          const { a, b } = await concurrentFormats(text, ra, rb, swap);
          const [first, second] = winnerFirst === loserIsB ? [a, b] : [b, a];
          ok(first.exec({ type: 'undo' }));
          expect(settled(a, b)).toEqual([text]);
          ok(second.exec({ type: 'undo' }));
          expect(settled(a, b)).toEqual([text]);
          expect(inlineOf(a, text.slice(0, 1))).toEqual([text]);
          ok(first.exec({ type: 'redo' }));
          ok(second.exec({ type: 'redo' }));
          expect(settled(a, b)).toEqual([text]);
        });

      test(`${name}: both undo at once (${order(swap)})`, async () => {
        const { a, b, sync } = await concurrentFormats(text, ra, rb, swap);
        sync.pause();
        ok(a.exec({ type: 'undo' }));
        ok(b.exec({ type: 'undo' }));
        sync.flush();
        expect(settled(a, b)).toEqual([text]);
        expect(inlineOf(a, text.slice(0, 1))).toEqual([text]);
      });
    }

  for (const swap of ORDERS)
    for (const joiner of ['a', 'b'] as const)
      test(`an undo after a join that took the winner's pieces keeps the word (${joiner} joins, ${order(swap)})`, async () => {
        const { a, b, sync } = await replicas(twoParagraphs(), swap);
        selectIn(a, 1, 6, 11);
        selectIn(b, 1, 6, 11);
        sync.pause();
        ok(a.exec({ type: 'toggleBold' }));
        ok(b.exec({ type: 'toggleItalic' }));
        const [j, other] = joiner === 'a' ? [a, b] : [b, a];
        selectIn(j, 1, 0, 0);
        ok(j.input('deleteBackward'));
        sync.flush();
        expect(settled(a, b)).toEqual(['FirstHello world']);
        // The peer that did not join undoes its format, then the joiner undoes its join and its
        // format: the word stays once at every step.
        ok(other.exec({ type: 'undo' }));
        expect(settled(a, b)).toEqual(['FirstHello world']);
        ok(j.exec({ type: 'undo' }));
        expect(settled(a, b)).toEqual(['First', 'Hello world']);
        ok(j.exec({ type: 'undo' }));
        expect(settled(a, b)).toEqual(['First', 'Hello world']);
      });

  for (const ids of [
    [1_001, 2_002, 3_003],
    [3_003, 2_002, 1_001],
    [2_002, 3_003, 1_001],
  ])
    test(`both formatting peers undo after a third peer joined (client ids ${ids.join(', ')})`, async () => {
      const { editors, sync } = await peers(twoParagraphs(), ids);
      const [a, b, c] = editors as [PptxEditorInstance, PptxEditorInstance, PptxEditorInstance];
      selectIn(a, 1, 6, 11);
      selectIn(b, 1, 6, 11);
      sync.pause();
      ok(a.exec({ type: 'toggleBold' }));
      ok(b.exec({ type: 'toggleItalic' }));
      selectIn(c, 1, 0);
      ok(c.input('deleteBackward'));
      sync.flush();
      expect(converged(a, b, c)).toEqual(['FirstHello world']);
      ok(b.exec({ type: 'undo' }));
      expect(converged(a, b, c)).toEqual(['FirstHello world']);
      ok(a.exec({ type: 'undo' }));
      expect(converged(a, b, c)).toEqual(['FirstHello world']);
      expect(inlineOf(a, 'First')).toEqual(['First', 'Hello world']);
    });

  for (const swap of ORDERS)
    test(`a redo of a split after a peer split the run again shows the word once (${order(swap)})`, async () => {
      const { a, b } = await replicas(oneRun('Hello world'), swap);
      selectIn(a, 0, 6, 11);
      ok(a.exec({ type: 'toggleBold' }));
      ok(a.exec({ type: 'undo' }));
      selectIn(b, 0, 6, 11);
      ok(b.exec({ type: 'toggleItalic' }));
      a.exec({ type: 'redo' });
      expect(settled(a, b)).toEqual(['Hello world']);
      // The peer's later format stays.
      expect(formatOf(a, 'world')).toEqual({ i: '1' });
    });

  for (const swap of ORDERS)
    test(`a redo after a peer deleted part of the word keeps the delete (${order(swap)})`, async () => {
      const { a, b } = await replicas(oneRun('Hello world'), swap);
      selectIn(a, 0, 6, 11);
      ok(a.exec({ type: 'toggleBold' }));
      ok(a.exec({ type: 'undo' }));
      selectIn(b, 0, 10, 11);
      ok(b.input('deleteBackward'));
      a.exec({ type: 'redo' });
      expect(settled(a, b)).toEqual(['Hello worl']);
    });
});

describe('hostile undone listings', () => {
  /** A peer that formats, and a peer that types into the format's run, then the undo. */
  async function undoneRun(swap: boolean) {
    const r = await replicas(oneRun('Hello world'), swap);
    selectIn(r.b, 0, 6, 11);
    ok(r.b.exec({ type: 'toggleItalic' }));
    selectIn(r.a, 0, 11, 11);
    ok(r.a.input('text', '!'));
    ok(r.b.exec({ type: 'undo' }));
    expect(settled(r.a, r.b)).toEqual(['Hello world!']);
    return r;
  }

  for (const swap of ORDERS)
    test(`an entry of a wrong shape is ignored (${order(swap)})`, async () => {
      const { a, b, docs } = await undoneRun(swap);
      const undone = docs[1].getMap(UNDONE_KEY);
      docs[1].transact(() => {
        for (const key of [...undone.keys()]) undone.set(key, { p: 7, u: [], a: null, c: -1 });
      }, 'hostile');
      // No entry is valid, so the run the undo unlisted no longer shows.
      expect(paragraphs(a)).toEqual(['Hello world']);
      selectIn(a, 0, 0, 0);
      ok(a.input('text', '>'));
      expect(converged(a, b)).toEqual(['>Hello world']);
    });

  for (const swap of ORDERS)
    test(`a record it shows that the core refuses refuses the part, and heals (${order(swap)})`, async () => {
      const { a, b, docs } = await undoneRun(swap);
      const nodes = docs[1].getMap(NODES_KEY);
      const undone = docs[1].getMap(UNDONE_KEY);
      // The unlisted run: a peer gives it an attribute name that XML cannot hold.
      const r = [...undone.keys()].find(
        (k) => ((nodes.get(k) as Y.Map<unknown>)?.get(F.shell) as { local?: string })?.local === 'r'
      )!;
      const attrs = (nodes.get(r) as Y.Map<unknown>).get(F.attrs) as Y.Map<unknown>;
      const good = canonicalPackageFingerprint(a.session!.store.package);
      const bad = `${ATTR_SEP}bad name`;
      docs[1].transact(() => attrs.set(bad, { ns: '', local: 'bad name', value: '1' }), 'hostile');
      for (const s of sessionsOf([a, b])) {
        expect(s.statusSnapshot().reason?.code).toBe('remote-apply-failed');
        expect(s.statusSnapshot().recovering).toBe(true);
      }
      expect(canonicalPackageFingerprint(a.session!.store.package)).toBe(good);
      docs[1].transact(() => attrs.delete(bad), 'hostile');
      expect(settled(a, b)).toEqual(['Hello world!']);
    });
});
