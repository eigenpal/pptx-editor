// Find and replace: matches in slide text, groups, tables, and notes in deck order, match case
// and whole word, stepping that selects the match on its slide, replace that keeps the
// formatting of the first replaced character, and replace all as one undo step.

import { describe, expect, test } from 'bun:test';
import {
  buildPptx,
  cell,
  para,
  run,
  shape,
  table,
} from '../../../../../scripts/fixtures/builder.ts';
import { attr, child, elements, find, textOf, xmlPartOf, type Element } from '../../store/index.ts';
import { createPptxEditor, notesText, type PptxEditorInstance } from '../index.ts';
import { matchesIn } from '../../automation/commands/find.ts';
import { findNotesMatch, onFindShown } from '../find-session.ts';
import { notesMatchRange } from '../notes-find-highlight.ts';

const IN = 914_400;

function group(inner: string): string {
  return (
    '<p:grpSp><p:nvGrpSpPr><p:cNvPr id="20" name="G"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
    `<p:grpSpPr><a:xfrm><a:off x="0" y="${4 * IN}"/><a:ext cx="${2 * IN}" cy="${IN}"/>` +
    `<a:chOff x="0" y="${4 * IN}"/><a:chExt cx="${2 * IN}" cy="${IN}"/></a:xfrm></p:grpSpPr>` +
    `${inner}</p:grpSp>`
  );
}

async function open(): Promise<PptxEditorInstance> {
  const bytes = buildPptx({
    slides: [
      {
        shapes:
          shape({
            id: 2,
            name: 'T',
            x: IN,
            y: IN,
            cx: 6 * IN,
            cy: IN,
            text:
              para(run('Alpha ') + run('Bold', 'b="1"') + run('Plain')) +
              para(run('alphabet soup')),
          }) +
          group(
            shape({
              id: 21,
              name: 'In',
              x: 0,
              y: 4 * IN,
              cx: 2 * IN,
              cy: IN,
              text: para(run('alpha')),
            })
          ) +
          table({
            id: 3,
            x: IN,
            y: 2 * IN,
            cx: 4 * IN,
            cy: IN,
            columns: [2 * IN, 2 * IN],
            tblPr: '<a:tblPr/>',
            rows: `<a:tr h="${IN / 2}">${cell('alpha cell')}${cell('other')}</a:tr>`,
          }),
      },
      {
        shapes: shape({
          id: 2,
          name: 'U',
          x: IN,
          y: IN,
          cx: 6 * IN,
          cy: IN,
          text: para(run('ALPHA')),
        }),
        notes: 'say alpha here',
      },
    ],
  });
  const e = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await e.load(bytes)).ok).toBe(true);
  return e;
}

const root = (e: PptxEditorInstance, slide: number): Element =>
  xmlPartOf(e.session!.store.package, e.session!.info.slides[slide].partName)!.root;
const shapeId = (e: PptxEditorInstance, slide: number, name: string): string =>
  elements(find(root(e, slide), 'p:spTree')!).find(
    (c) => attr(find(c, 'p:cNvPr'), 'name') === name
  )!.id;
const frameId = (e: PptxEditorInstance): string =>
  elements(find(root(e, 0), 'p:spTree')!).find((c) => c.local === 'graphicFrame')!.id;
/** The runs of the first paragraph of shape `T` on slide 1: text and bold flag. */
const runs = (e: PptxEditorInstance): [string, boolean][] =>
  elements(find(root(e, 0), 'a:p')!)
    .filter((c) => c.local === 'r')
    .map((r) => [textOf(find(r, 'a:t')), attr(child(r, 'a:rPr'), 'b') === '1']);

describe('find', () => {
  test('matches in deck order: shapes, groups, tables, then notes, slide by slide', async () => {
    const e = await open();
    const list = e.find.matches('alpha');
    expect(list.map((m) => [m.slide, m.where, m.text])).toEqual([
      [0, 'slide', 'Alpha'],
      [0, 'slide', 'alpha'],
      [0, 'slide', 'alpha'],
      [0, 'slide', 'alpha'],
      [1, 'slide', 'ALPHA'],
      [1, 'notes', 'alpha'],
    ]);
    expect(list[1]).toMatchObject({ paragraph: 1, start: 0, end: 5 });
    expect(list[2].nodeId).toBe(find(find(root(e, 0), 'p:grpSp')!, 'p:sp')!.id);
    expect(list[3].nodeId).toBe(frameId(e));
  });

  test('match case and whole word', async () => {
    const e = await open();
    expect(e.find.matches('alpha', { matchCase: true })).toHaveLength(4);
    expect(e.find.matches('alpha', { wholeWord: true })).toHaveLength(5);
    expect(e.find.matches('alpha', { matchCase: true, wholeWord: true })).toHaveLength(3);
    expect(matchesIn('a_b ab-ab', 'ab', { wholeWord: true })).toEqual([
      [4, 6],
      [7, 9],
    ]);
    expect(matchesIn('aaa', 'aa')).toEqual([[0, 2]]);
  });

  test('next selects the match as text on its slide, switching slides; previous wraps', async () => {
    const e = await open();
    const first = e.find.next('alpha');
    expect(first.ok && first.index).toBe(0);
    expect(e.snapshot().selection).toEqual({
      kind: 'text',
      nodeId: shapeId(e, 0, 'T'),
      anchor: { paragraph: 0, offset: 0 },
      focus: { paragraph: 0, offset: 5 },
    });
    expect(e.find.position('alpha')).toBe(0);
    e.find.next('alpha');
    e.find.next('alpha');
    // A match in a table cell selects its text in the cell.
    e.find.next('alpha');
    expect(e.snapshot().selection).toEqual({
      kind: 'text',
      nodeId: frameId(e),
      cell: { row: 0, column: 0 },
      anchor: { paragraph: 0, offset: 0 },
      focus: { paragraph: 0, offset: 5 },
    });
    expect(e.find.position('alpha')).toBe(3);
    const next = e.find.next('alpha');
    expect(next.ok && next.index).toBe(4);
    expect(e.snapshot().currentSlide).toBe(1);
    const notes = e.find.next('alpha');
    expect(notes.ok && notes.match?.where).toBe('notes');
    expect(e.find.position('alpha')).toBe(5);
    const wrapped = e.find.next('alpha');
    expect(wrapped.ok && wrapped.index).toBe(0);
    expect(e.snapshot().currentSlide).toBe(0);
    const back = e.find.previous('alpha');
    expect(back.ok && back.index).toBe(5);
    expect(e.find.next('zebra')).toEqual({ ok: true, match: null, index: -1, total: 0 });
  });

  test('next starts after the caret', async () => {
    const e = await open();
    e.setSelection({
      kind: 'text',
      nodeId: shapeId(e, 0, 'T'),
      anchor: { paragraph: 0, offset: 2 },
      focus: { paragraph: 0, offset: 2 },
    });
    const r = e.find.next('alpha');
    expect(r.ok && r.index).toBe(1);
  });
});

describe('replace', () => {
  test('the new text keeps the formatting of the first replaced character', async () => {
    const e = await open();
    const step = e.find.next('oldPl');
    expect(step.ok && step.match?.text).toBe('oldPl');
    const r = e.find.replace('oldPl', 'NEW');
    expect(r.ok).toBe(true);
    expect(runs(e)).toEqual([
      ['Alpha ', false],
      ['BNEW', true],
      ['ain', false],
    ]);
  });

  test('replace without a shown match selects the next match first', async () => {
    const e = await open();
    const r = e.find.replace('soup', 'stew');
    expect(r.ok && r.match?.text).toBe('soup');
    expect(e.find.matches('stew')).toHaveLength(0);
    e.find.replace('soup', 'stew');
    expect(e.find.matches('stew')).toHaveLength(1);
  });

  test('replace in speaker notes', async () => {
    const e = await open();
    const all = e.find.matches('alpha');
    expect(e.exec({ type: 'replace', match: all[5], replacement: 'beta' }).ok).toBe(true);
    expect(notesText(e.session!.store.package, 1)).toBe('say beta here');
  });

  test('replace all is one undo step and returns the count', async () => {
    const e = await open();
    const before = e.session!.store.package;
    const r = e.find.replaceAll('alpha', 'omega');
    expect(r).toEqual({ ok: true, count: 6 });
    expect(e.find.matches('alpha')).toHaveLength(0);
    expect(e.find.matches('omega').map((m) => m.text)).toEqual([
      'omega',
      'omega',
      'omega',
      'omega',
      'omega',
      'omega',
    ]);
    expect(notesText(e.session!.store.package, 1)).toBe('say omega here');
    expect(e.exec({ type: 'undo' }).ok).toBe(true);
    expect(e.session!.store.package).toBe(before);
    expect(e.find.matches('alpha')).toHaveLength(6);
    expect(e.snapshot().canUndo).toBe(false);
  });

  test('refusals are typed: stale matches, no match, view mode', async () => {
    const e = await open();
    const m = e.find.matches('Plain')[0];
    expect(e.exec({ type: 'replace', match: { ...m, text: 'Other' }, replacement: 'x' })).toEqual({
      ok: false,
      reason: 'errors.findStale',
    });
    expect(e.can({ type: 'replaceAll', query: 'zebra', replacement: 'x' })).toEqual({
      ok: false,
      reason: 'disabled.noMatch',
    });
    e.setMode('view');
    expect(e.find.replaceAll('alpha', 'x')).toEqual({ ok: false, reason: 'disabled.viewing' });
    expect(e.find.matches('alpha')).toHaveLength(6);
  });
});

test('find skips the hidden text of covered cells and finds orphan merge cells, as the layout draws them', async () => {
  const bytes = buildPptx({
    slides: [
      {
        shapes: table({
          id: 3,
          x: IN,
          y: IN,
          cx: 6 * IN,
          cy: IN,
          columns: [2 * IN, 2 * IN, 2 * IN],
          tblPr: '<a:tblPr/>',
          rows:
            `<a:tr h="${IN / 2}">${cell('alpha', { attrs: ' gridSpan="2"' })}` +
            `${cell('alpha hidden', { attrs: ' hMerge="1"' })}${cell('other')}</a:tr>` +
            `<a:tr h="${IN / 2}">${cell('x')}${cell('y')}${cell('alpha orphan', { attrs: ' hMerge="1"' })}</a:tr>`,
        }),
      },
    ],
  });
  const e = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await e.load(bytes)).ok).toBe(true);
  expect(e.find.matches('alpha')).toHaveLength(2);
  expect(e.find.matches('hidden')).toHaveLength(0);
  // A merge marker that no span reaches does not hide its cell: the cell is drawn and found.
  expect(e.find.matches('orphan')).toHaveLength(1);
});

test('each step tells its listeners, and a notes match stays shown until the slide moves on', async () => {
  const e = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  const bytes = buildPptx({ slides: [{ shapes: '', notes: 'one one\nand one' }] });
  expect((await e.load(bytes)).ok).toBe(true);
  let told = 0;
  const stop = onFindShown(e.find, () => (told += 1));
  e.find.next('one');
  e.find.next('one');
  // The steps stay on the slide with no selection: only the listener tells the find bar.
  expect(told).toBe(2);
  expect(e.find.position('one')).toBe(1);
  const m = findNotesMatch(e.find)!;
  expect([m.paragraph, m.start, m.end]).toEqual([0, 4, 7]);
  const third = e.find.next('one');
  expect(third.ok && notesMatchRange(e.snapshot().notesFormatted ?? [], third.match!)).toEqual({
    start: 12,
    end: 15,
  });
  stop();
  e.find.next('one');
  expect(told).toBe(3);
  e.setSelection({ kind: 'shapes', nodeIds: ['x'] });
  expect(findNotesMatch(e.find) === null).toBe(true);
});
