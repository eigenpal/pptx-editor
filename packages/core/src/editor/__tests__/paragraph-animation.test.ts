// Effects that animate some paragraphs of a shape (`p:spTgt/p:txEl/p:pRg`) follow their
// paragraphs through text edits: a range takes the new indexes of its paragraphs, and an
// effect whose paragraphs are all gone goes, with its build entry when no effect is left.

import { describe, expect, test } from 'bun:test';
import { buildPptx, para, run, shape } from '../../../../../scripts/fixtures/builder.ts';
import { attr, elements, find, readPptx, xmlPartOf, type Element } from '../../store/index.ts';
import { createPptxEditor, type PptxEditorInstance } from '../index.ts';

const IN = 914_400;

/** A click entrance of paragraphs `st` to `end` of shape 2, with `p:cTn/@id` `id`. */
const effect = (id: number, st: number, end: number): string =>
  '<p:par><p:cTn id="' +
  (id + 1) +
  '" fill="hold"><p:stCondLst><p:cond delay="indefinite"/></p:stCondLst><p:childTnLst>' +
  `<p:par><p:cTn id="${id}" presetID="1" presetClass="entr" presetSubtype="0" fill="hold" grpId="0" nodeType="clickEffect">` +
  '<p:stCondLst><p:cond delay="0"/></p:stCondLst><p:childTnLst>' +
  `<p:set><p:cBhvr><p:cTn id="${id + 2}" dur="1" fill="hold"><p:stCondLst><p:cond delay="0"/></p:stCondLst></p:cTn>` +
  `<p:tgtEl><p:spTgt spid="2"><p:txEl><p:pRg st="${st}" end="${end}"/></p:txEl></p:spTgt></p:tgtEl>` +
  '<p:attrNameLst><p:attrName>style.visibility</p:attrName></p:attrNameLst></p:cBhvr>' +
  '<p:to><p:strVal val="visible"/></p:to></p:set></p:childTnLst></p:cTn></p:par>' +
  '</p:childTnLst></p:cTn></p:par>';

const timing = (effects: string): string =>
  '<p:timing><p:tnLst><p:par><p:cTn id="1" dur="indefinite" restart="never" nodeType="tmRoot"><p:childTnLst>' +
  '<p:seq concurrent="1" nextAc="seek"><p:cTn id="2" dur="indefinite" nodeType="mainSeq"><p:childTnLst>' +
  effects +
  '</p:childTnLst></p:cTn><p:prevCondLst><p:cond evt="onPrev" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:prevCondLst>' +
  '<p:nextCondLst><p:cond evt="onNext" delay="0"><p:tgtEl><p:sldTgt/></p:tgtEl></p:cond></p:nextCondLst></p:seq>' +
  '</p:childTnLst></p:cTn></p:par></p:tnLst><p:bldLst><p:bldP spid="2" grpId="0" build="p"/></p:bldLst></p:timing>';

/** A deck whose shape 2 holds paragraphs A, B, and C; B and C each have an effect. */
async function open(
  effects = effect(10, 1, 1) + effect(20, 2, 2),
  texts = ['Alpha', 'Beta', 'Gamma']
): Promise<PptxEditorInstance> {
  const e = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  const text = texts.map((t) => para(t ? run(t) : '')).join('');
  const deck = buildPptx({
    slides: [
      {
        shapes: shape({ id: 2, name: 'List', x: IN, y: IN, cx: 6 * IN, cy: 3 * IN, text }),
        ext: timing(effects),
      },
    ],
  });
  expect((await e.load(deck)).ok).toBe(true);
  return e;
}

const rootOf = (e: PptxEditorInstance): Element =>
  xmlPartOf(e.session!.store.package, e.session!.info.slides[0]!.partName)!.root;
const shapeNode = (e: PptxEditorInstance): string => find(rootOf(e), 'p:sp')!.id;

/** The `[st, end]` of each paragraph range, in document order. */
function ranges(root: Element): [string | undefined, string | undefined][] {
  const out: [string | undefined, string | undefined][] = [];
  const walk = (el: Element): void => {
    if (el.local === 'pRg') out.push([attr(el, 'st'), attr(el, 'end')]);
    for (const c of elements(el)) walk(c);
  };
  walk(find(root, 'p:timing') ?? root);
  return out;
}

/** The `p:cTn/@id` of each click effect, in document order. */
function ids(root: Element): string[] {
  const out: string[] = [];
  const walk = (el: Element): void => {
    if (el.local === 'cTn' && attr(el, 'nodeType') === 'clickEffect') out.push(attr(el, 'id')!);
    for (const c of elements(el)) walk(c);
  };
  walk(find(root, 'p:timing') ?? root);
  return out;
}

/** Delete from `[p, o]` to `[p2, o2]` in shape 2. */
function remove(e: PptxEditorInstance, p: number, o: number, p2: number, o2: number): void {
  e.setSelection({
    kind: 'text',
    nodeId: shapeNode(e),
    anchor: { paragraph: p, offset: o },
    focus: { paragraph: p2, offset: o2 },
  });
  expect(e.input('deleteBackward').ok).toBe(true);
}

/** The ranges after a save and a reopen. */
async function reopened(
  e: PptxEditorInstance
): Promise<[string | undefined, string | undefined][]> {
  const bytes = e.save()!;
  expect(readPptx(bytes).ok).toBe(true);
  const again = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await again.load(bytes)).ok).toBe(true);
  return ranges(rootOf(again));
}

describe('paragraph ranges of effects', () => {
  test('deleting a paragraph before the animated ones moves their ranges', async () => {
    const e = await open();
    remove(e, 0, 0, 1, 0);
    expect(ranges(rootOf(e))).toEqual([
      ['0', '0'],
      ['1', '1'],
    ]);
    expect(await reopened(e)).toEqual([
      ['0', '0'],
      ['1', '1'],
    ]);
  });

  test('Enter before an animated paragraph moves its range', async () => {
    const e = await open();
    e.setSelection({
      kind: 'text',
      nodeId: shapeNode(e),
      anchor: { paragraph: 0, offset: 0 },
      focus: { paragraph: 0, offset: 0 },
    });
    expect(e.input('paragraph').ok).toBe(true);
    expect(ranges(rootOf(e))).toEqual([
      ['2', '2'],
      ['3', '3'],
    ]);
  });

  test('deleting the text of every animated paragraph removes its effect', async () => {
    const e = await open(effect(10, 1, 2));
    remove(e, 0, 5, 2, 5);
    expect(ranges(rootOf(e))).toEqual([]);
    expect(readPptx(e.save()!).ok).toBe(true);
    // The build entry went with the last effect of its shape.
    expect(find(rootOf(e), 'p:bldP')).toBeUndefined();
  });

  test('joining two animated paragraphs gives both effects the joined paragraph', async () => {
    const e = await open();
    remove(e, 2, 0, 2, 0);
    expect(ranges(rootOf(e))).toEqual([
      ['1', '1'],
      ['1', '1'],
    ]);
  });

  test('deleting an empty paragraph before an animated one moves its range', async () => {
    const e = await open();
    e.setSelection({
      kind: 'text',
      nodeId: shapeNode(e),
      anchor: { paragraph: 0, offset: 5 },
      focus: { paragraph: 0, offset: 5 },
    });
    expect(e.input('paragraph').ok).toBe(true);
    expect(ranges(rootOf(e))).toEqual([
      ['2', '2'],
      ['3', '3'],
    ]);
    remove(e, 1, 0, 1, 0);
    expect(ranges(rootOf(e))).toEqual([
      ['1', '1'],
      ['2', '2'],
    ]);
  });

  test('splitting an animated paragraph keeps both parts in its range', async () => {
    const e = await open();
    e.setSelection({
      kind: 'text',
      nodeId: shapeNode(e),
      anchor: { paragraph: 1, offset: 2 },
      focus: { paragraph: 1, offset: 2 },
    });
    expect(e.input('paragraph').ok).toBe(true);
    expect(ranges(rootOf(e))).toEqual([
      ['1', '2'],
      ['3', '3'],
    ]);
  });

  test('a paragraph whose text was replaced loses its effect', async () => {
    const e = await open(effect(10, 0, 0) + effect(20, 2, 2));
    e.setSelection({
      kind: 'text',
      nodeId: shapeNode(e),
      anchor: { paragraph: 0, offset: 0 },
      focus: { paragraph: 1, offset: 4 },
    });
    expect(e.input('text', 'New').ok).toBe(true);
    // Alpha and Beta became one paragraph, "New"; Gamma moved up.
    expect(ranges(rootOf(e))).toEqual([['1', '1']]);
  });

  test('a save ends a range that is past the last paragraph at it', async () => {
    const e = await open();
    const pRg = find(find(rootOf(e), 'p:timing')!, 'p:pRg')!;
    const part = e.session!.info.slides[0]!.partName;
    const r = e.session!.store.transact([
      { op: 'setAttribute', part, node: pRg.id, name: { ns: '', local: 'end' }, value: '7' },
    ]);
    expect(r.ok).toBe(true);
    expect((await reopened(e))[0]).toEqual(['1', '2']);
  });

  test('undo restores the ranges with the text', async () => {
    const e = await open();
    remove(e, 0, 0, 1, 0);
    expect(e.exec({ type: 'undo' }).ok).toBe(true);
    expect(ranges(rootOf(e))).toEqual([
      ['1', '1'],
      ['2', '2'],
    ]);
  });

  test('deleting the first of two identical paragraphs keeps the effect of the second', async () => {
    const e = await open(effect(10, 0, 0) + effect(20, 1, 1), ['Same', 'Same', 'Gamma']);
    remove(e, 0, 0, 1, 0);
    // The effect of the deleted paragraph goes; the effect of the second moves up.
    expect(ranges(rootOf(e))).toEqual([['0', '0']]);
    expect(ids(rootOf(e))).toEqual(['20']);
  });

  test('deleting the second of two identical paragraphs keeps the effect of the first', async () => {
    const e = await open(effect(10, 0, 0) + effect(20, 1, 1), ['Same', 'Same', 'Gamma']);
    remove(e, 0, 4, 1, 4);
    expect(ranges(rootOf(e))).toEqual([['0', '0']]);
    expect(ids(rootOf(e))).toEqual(['10']);
  });

  test('Enter between identical paragraphs moves only the ranges after the caret', async () => {
    const e = await open(effect(10, 0, 0) + effect(20, 1, 1), ['', '', 'Gamma']);
    e.setSelection({
      kind: 'text',
      nodeId: shapeNode(e),
      anchor: { paragraph: 1, offset: 0 },
      focus: { paragraph: 1, offset: 0 },
    });
    expect(e.input('paragraph').ok).toBe(true);
    expect(ranges(rootOf(e))).toEqual([
      ['0', '0'],
      ['2', '2'],
    ]);
  });

  test('an edit with many changes keeps the ranges of the paragraphs it did not change', async () => {
    const e = await open(effect(10, 1, 1) + effect(20, 3, 3), [
      'a'.repeat(300),
      'Beta',
      'Gamma',
      'Delta',
      'b'.repeat(300),
    ]);
    const part = e.session!.info.slides[0]!.partName;
    const paras = find(rootOf(e), 'p:txBody')!;
    const ps = elements(paras).filter((p) => p.local === 'p');
    const t = (p: Element): string => find(p, 'a:t')!.id;
    // Two far-apart changes that differ in more places than a character diff compares, and the
    // removal of Gamma: Beta stays at 1, and Delta moves from 3 to 2.
    const r = e.session!.store.transact([
      { op: 'setText', part, node: t(ps[0]!), text: 'x'.repeat(300) },
      { op: 'setText', part, node: t(ps[4]!), text: 'y'.repeat(300) },
      { op: 'removeNode', part, node: ps[2]!.id },
    ]);
    expect(r.ok).toBe(true);
    expect(ranges(rootOf(e))).toEqual([
      ['1', '1'],
      ['2', '2'],
    ]);
  });

  test('typing in an animated paragraph keeps every range', async () => {
    const e = await open();
    e.setSelection({
      kind: 'text',
      nodeId: shapeNode(e),
      anchor: { paragraph: 1, offset: 0 },
      focus: { paragraph: 1, offset: 0 },
    });
    expect(e.input('text', 'New ').ok).toBe(true);
    expect(ranges(rootOf(e))).toEqual([
      ['1', '1'],
      ['2', '2'],
    ]);
  });
});
