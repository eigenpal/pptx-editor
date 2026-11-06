// Removing a shape removes the animation entries that name it (ECMA-376 Part 1, §19.5
// `p:timing`) and the connector ends glued to it, so a saved slide names no missing shape and
// its `p:timing` stays valid against `CT_SlideTiming`. A slide the session did not
// change saves as the file wrote it.

import { describe, expect, test } from 'bun:test';
import {
  buildPptx,
  cell,
  para,
  run,
  shape,
  table,
  type FixtureSlide,
} from '../../../../../scripts/fixtures/builder.ts';
import { clickEntrances, timingWriter } from '../../../../../scripts/fixtures/timing.ts';
import {
  attr,
  canonicalPackageFingerprint,
  child,
  elements,
  find,
  NS,
  timingRemovals,
  readPptx,
  semanticDigest,
  serializeXmlPart,
  xmlPartOf,
  type Element,
  type PptxPackage,
} from '../../store/index.ts';
import { createPptxEditor, type PptxEditorInstance } from '../index.ts';

const IN = 914_400;

const { show, effect, click, mainSeq, interactive, timing } = timingWriter();

const box = (id: number, x: number, text = `S${id}`): string =>
  shape({ id, x: x * IN, y: IN, cx: IN, cy: IN, text: para(run(text)) });
const GROUP =
  '<p:grpSp><p:nvGrpSpPr><p:cNvPr id="10" name="Group 9"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>' +
  '<p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="2743200" cy="914400"/><a:chOff x="0" y="0"/><a:chExt cx="2743200" cy="914400"/></a:xfrm></p:grpSpPr>' +
  shape({ id: 11, x: 0, y: 0, cx: IN, cy: IN }) +
  shape({ id: 12, x: 2 * IN, y: 0, cx: IN, cy: IN }) +
  '</p:grpSp>';

// ---- Helpers ----

async function open(slide: FixtureSlide): Promise<PptxEditorInstance> {
  const e = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await e.load(buildPptx({ slides: [slide] }))).ok).toBe(true);
  return e;
}
const pkgOf = (e: PptxEditorInstance): PptxPackage => e.session!.store.package;
const slidePart = (e: PptxEditorInstance): string => e.session!.info.slides[0].partName;
const rootOf = (pkg: PptxPackage, part: string): Element => xmlPartOf(pkg, part)!.root;
const xmlOf = (e: PptxEditorInstance): string =>
  serializeXmlPart(xmlPartOf(pkgOf(e), slidePart(e))!);

/** The node id of the shape whose `p:cNvPr/@id` is `id`. */
function nodeOf(e: PptxEditorInstance, id: number): string {
  const walk = (el: Element): string | undefined => {
    for (const c of elements(el)) {
      const nv = elements(c).find((x) => x.ns === NS.p && x.local.startsWith('nv'));
      if (attr(child(nv, 'p:cNvPr'), 'id') === String(id)) return c.id;
      const hit = c.local === 'grpSp' ? walk(c) : undefined;
      if (hit) return hit;
    }
    return undefined;
  };
  const hit = walk(find(rootOf(pkgOf(e), slidePart(e)), 'p:spTree')!);
  if (!hit) throw new Error(`no shape ${id}`);
  return hit;
}

function deleteIds(e: PptxEditorInstance, ids: readonly number[]): void {
  e.setSelection({ kind: 'shapes', nodeIds: ids.map((id) => nodeOf(e, id)) });
  const r = e.exec({ type: 'delete' });
  expect(r.ok).toBe(true);
}

/** Every `spid` in the slide's timing. */
function spids(root: Element): string[] {
  const out: string[] = [];
  const walk = (el: Element): void => {
    const v = attr(el, 'spid');
    if (v !== undefined) out.push(v);
    for (const c of elements(el)) walk(c);
  };
  for (const t of timings(root)) walk(t);
  return out;
}

function timings(root: Element): Element[] {
  const out = elements(root).filter((c) => c.local === 'timing');
  for (const ac of elements(root).filter((c) => c.local === 'AlternateContent'))
    for (const b of elements(ac)) out.push(...elements(b).filter((c) => c.local === 'timing'));
  return out;
}

function drawingIds(root: Element): Set<string> {
  const ids = new Set<string>();
  const walk = (el: Element): void => {
    if (el.local === 'cNvPr') ids.add(attr(el, 'id')!);
    for (const c of elements(el)) walk(c);
  };
  walk(find(root, 'p:spTree')!);
  return ids;
}

const TIME_NODES = new Set([
  'par',
  'seq',
  'excl',
  'anim',
  'animClr',
  'animEffect',
  'animMotion',
  'animRot',
  'animScale',
  'cmd',
  'set',
  'audio',
  'video',
]);

/**
 * The `CT_SlideTiming` rules that a removal can break: each time node list and build list has
 * a child, each time node is a time node, each behavior has one `p:tgtEl` target, each
 * condition list has a `p:cond`, and every `spid` names a shape of the slide.
 */
function expectValidTiming(root: Element): void {
  const ids = drawingIds(root);
  for (const spid of spids(root)) expect(ids.has(spid)).toBe(true);
  const walk = (el: Element): void => {
    if (['tnLst', 'childTnLst', 'subTnLst'].includes(el.local)) {
      expect(elements(el).length).toBeGreaterThan(0);
      for (const n of elements(el)) expect(TIME_NODES.has(n.local)).toBe(true);
    }
    if (el.local === 'bldLst') expect(elements(el).length).toBeGreaterThan(0);
    if (['stCondLst', 'endCondLst', 'prevCondLst', 'nextCondLst'].includes(el.local))
      expect(elements(el).length).toBeGreaterThan(0);
    if (el.local === 'cBhvr' || el.local === 'cMediaNode')
      expect(elements(child(el, 'p:tgtEl')).length).toBe(1);
    for (const c of elements(el)) walk(c);
  };
  for (const t of timings(root)) walk(t);
}

/** Save, reopen, and check the reopened slide; returns its root. */
function saveAndReopen(e: PptxEditorInstance): Element {
  const bytes = e.save()!;
  const r = readPptx(bytes);
  expect(r.ok).toBe(true);
  if (!r.ok) throw new Error('reopen failed');
  // The save check: a second save of the reopened copy changes nothing.
  expect(semanticDigest(r.pkg)).toBe(canonicalPackageFingerprint(r.pkg));
  const root = rootOf(r.pkg, slidePart(e));
  expectValidTiming(root);
  return root;
}

function expectUndoRestores(e: PptxEditorInstance, before: string): void {
  expect(e.exec({ type: 'undo' }).ok).toBe(true);
  expect(canonicalPackageFingerprint(pkgOf(e))).toBe(before);
}

// ---- Tests ----

describe('deleting a shape removes the animation entries that name it', () => {
  test('one effect of two goes with its build entry; the other stays', async () => {
    const e = await open({
      shapes: box(2, 1) + box(3, 3),
      ext: timing(mainSeq(click(effect(2)) + click(effect(3))), [2, 3]),
    });
    const before = canonicalPackageFingerprint(pkgOf(e));
    deleteIds(e, [2]);
    const root = saveAndReopen(e);
    expect(spids(root)).toEqual(['3', '3']);
    // The click step of shape 2 went whole: one step is left in the main sequence.
    const seq = find(root, 'p:seq')!;
    expect(elements(child(child(seq, 'p:cTn'), 'p:childTnLst'))).toHaveLength(1);
    expectUndoRestores(e, before);
  });

  test('the last animated shape takes the whole p:timing with it', async () => {
    const e = await open({
      shapes: box(2, 1) + box(3, 3),
      ext: timing(mainSeq(click(effect(2))), [2]),
    });
    const before = canonicalPackageFingerprint(pkgOf(e));
    deleteIds(e, [2]);
    expect(xmlOf(e)).not.toContain('p:timing');
    const root = saveAndReopen(e);
    expect(child(root, 'p:timing')).toBeUndefined();
    expectUndoRestores(e, before);
    expect(xmlOf(e)).toContain('<p:bldP spid="2"');
  });

  test('an effect with behaviors on two shapes keeps the behavior of the shape that stays', async () => {
    const e = await open({
      shapes: box(2, 1) + box(3, 3),
      ext: timing(mainSeq(click(effect(2, show(3)))), [2, 3]),
    });
    deleteIds(e, [2]);
    const root = saveAndReopen(e);
    expect(spids(root)).toEqual(['3', '3']);
    expect(find(root, 'p:set')).toBeDefined();
  });

  test('deleting a trigger shape removes its interactive sequence only', async () => {
    const e = await open({
      shapes: box(2, 1) + box(3, 3) + box(4, 5),
      ext: timing(mainSeq(click(effect(2))) + interactive(4, 3), [2, 3]),
    });
    deleteIds(e, [4]);
    const root = saveAndReopen(e);
    // Shape 3's only effect was in the sequence, so its build entry goes too.
    expect(spids(root).sort()).toEqual(['2', '2']);
    expect(xmlOf(e)).not.toContain('interactiveSeq');
    expect(xmlOf(e)).toContain('mainSeq');
  });

  test('a time node that waits for a removed time node goes too (p:tn)', async () => {
    // Effect 3 starts when effect 2 (cTn 900) ends.
    const second =
      `<p:par><p:cTn id="901" fill="hold"><p:stCondLst><p:cond evt="onEnd" delay="0"><p:tn val="900"/></p:cond></p:stCondLst>` +
      `<p:childTnLst>${effect(3)}</p:childTnLst></p:cTn></p:par>`;
    const first = effect(2, '', 900);
    const e = await open({
      shapes: box(2, 1) + box(3, 3) + box(4, 5),
      ext: timing(mainSeq(click(first + second) + click(effect(4))), [2, 3, 4]),
    });
    deleteIds(e, [2]);
    const root = saveAndReopen(e);
    // Effect 3 went with the time node it waited for, and its build entry with it.
    expect(spids(root).sort()).toEqual(['4', '4']);
    expect(xmlOf(e)).not.toContain('<p:tn val="900"/>');
  });

  test('deleting a group removes the animations of the shapes inside it', async () => {
    const e = await open({
      shapes: GROUP + box(2, 5),
      ext: timing(mainSeq(click(effect(11)) + click(effect(2))), [2]),
    });
    deleteIds(e, [10]);
    const root = saveAndReopen(e);
    expect(spids(root)).toEqual(['2', '2']);
  });

  test('deleting the last shapes of a group removes the group and their animations', async () => {
    const e = await open({
      shapes: GROUP + box(2, 5),
      ext: timing(mainSeq(click(effect(11)) + click(effect(10))), []),
    });
    e.setSelection({
      kind: 'shapes',
      nodeIds: [nodeOf(e, 11), nodeOf(e, 12)],
      group: nodeOf(e, 10),
    });
    expect(e.exec({ type: 'delete' }).ok).toBe(true);
    expect(xmlOf(e)).not.toContain('<p:grpSp>');
    expect(xmlOf(e)).not.toContain('p:timing');
    saveAndReopen(e);
  });

  test('ungrouping removes the group animation and keeps the animations of its shapes', async () => {
    const e = await open({
      shapes: GROUP,
      ext: timing(mainSeq(click(effect(10)) + click(effect(12))), []),
    });
    const before = canonicalPackageFingerprint(pkgOf(e));
    e.setSelection({ kind: 'shapes', nodeIds: [nodeOf(e, 10)] });
    expect(e.exec({ type: 'ungroup' }).ok).toBe(true);
    const root = saveAndReopen(e);
    expect(spids(root)).toEqual(['12']);
    expectUndoRestores(e, before);
  });

  test('deleting a table removes its animation (deleteTable and delete)', async () => {
    const frame = table({
      id: 5,
      x: IN,
      y: IN,
      cx: 2 * IN,
      cy: IN,
      columns: [IN, IN],
      tblPr: '<a:tblPr firstRow="1"/>',
      rows: `<a:tr h="${IN}">${cell('a')}${cell('b')}</a:tr>`,
    });
    for (const how of ['deleteTable', 'delete'] as const) {
      const e = await open({
        shapes: frame + box(2, 5),
        ext: timing(mainSeq(click(effect(5))), []),
      });
      if (how === 'deleteTable') {
        const at = { row: 0, column: 0 };
        e.setSelection({ kind: 'cells', nodeId: nodeOf(e, 5), anchor: at, focus: at });
        expect(e.exec({ type: 'deleteTable' }).ok).toBe(true);
      } else deleteIds(e, [5]);
      expect(xmlOf(e)).not.toContain('p:timing');
      saveAndReopen(e);
    }
  });

  test('timing in mc:AlternateContent: every branch loses the shape, and the wrapper goes when empty', async () => {
    const ac = (inner: string): string =>
      '<mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">' +
      `<mc:Choice xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" Requires="p14">${inner}</mc:Choice>` +
      `<mc:Fallback>${inner}</mc:Fallback></mc:AlternateContent>`;
    const partial = await open({
      shapes: box(2, 1) + box(3, 3),
      ext: ac(timing(mainSeq(click(effect(2)) + click(effect(3))), [2, 3])),
    });
    deleteIds(partial, [2]);
    let root = saveAndReopen(partial);
    expect(spids(root)).toEqual(['3', '3', '3', '3']);
    const whole = await open({
      shapes: box(2, 1) + box(3, 3),
      ext: ac(timing(mainSeq(click(effect(2))), [2])),
    });
    deleteIds(whole, [2]);
    root = saveAndReopen(whole);
    expect(xmlOf(whole)).not.toContain('AlternateContent');
    expect(timings(root)).toHaveLength(0);
  });

  test('turning slide numbers off removes the animation of the slide-number placeholder', async () => {
    const e = createPptxEditor({ fixedMeasurer: true, substitutes: false });
    const sldNum = shape({
      id: 4,
      name: 'Slide Number Placeholder 3',
      x: 8 * IN,
      y: 6 * IN,
      cx: IN,
      cy: IN / 2,
      ph: '<p:ph type="sldNum" sz="quarter" idx="12"/>',
      text: '<a:p><a:fld id="{B6F15528-21DE-4FAA-801E-634DDDAF4B2B}" type="slidenum"><a:rPr lang="en-US"/><a:t>1</a:t></a:fld></a:p>',
    });
    const bytes = buildPptx({
      footers: true,
      slides: [{ shapes: box(2, 1) + sldNum, ext: timing(mainSeq(click(effect(4))), [4]) }],
    });
    expect((await e.load(bytes)).ok).toBe(true);
    expect(e.exec({ type: 'setSlideNumbers', show: false }).ok).toBe(true);
    expect(xmlOf(e)).not.toContain('p:timing');
    saveAndReopen(e);
  });

  test('a spid that the file already left dangling is not touched', async () => {
    const e = await open({
      shapes: box(2, 1) + box(3, 3),
      ext: timing(mainSeq(click(effect(2)) + click(effect(99))), [2]),
    });
    deleteIds(e, [3]);
    expect(xmlOf(e)).toContain('<p:spTgt spid="99"/>');
    expect(xmlOf(e)).toContain('<p:spTgt spid="2"/>');
  });

  test('a build entry with no effect in the file stays', async () => {
    const e = await open({
      shapes: box(2, 1) + box(3, 3) + box(4, 5),
      ext: timing(mainSeq(click(effect(2)) + click(effect(3))), [2, 3, 4]),
    });
    deleteIds(e, [2]);
    const root = saveAndReopen(e);
    expect(spids(root).sort()).toEqual(['3', '3', '4']);
  });

  test('a long chain of p:tn waits goes in linear time', () => {
    // Node k waits for node k + 1; the last node animates shape 2. Each node sits before the one
    // it waits for, the worst order for a pass that repeats until nothing more goes.
    const n = 20_000;
    let chain = '';
    for (let k = 0; k < n; k += 1)
      chain +=
        `<p:par><p:cTn id="${1000 + k}" fill="hold"><p:stCondLst><p:cond evt="onEnd" delay="0"><p:tn val="${1001 + k}"/></p:cond></p:stCondLst>` +
        `<p:childTnLst>${show(3)}</p:childTnLst></p:cTn></p:par>`;
    chain += effect(2, '', 1000 + n);
    const bytes = buildPptx({
      slides: [
        {
          shapes: box(2, 1) + box(3, 3) + box(4, 5),
          ext: timing(mainSeq(click(chain) + click(effect(4))), [4]),
        },
      ],
    });
    const r = readPptx(bytes);
    if (!r.ok) throw new Error('open failed');
    const root = rootOf(r.pkg, '/ppt/slides/slide1.xml');
    const started = performance.now();
    const removed = timingRemovals(root, new Set(['2']));
    const ms = performance.now() - started;
    // The whole first click step goes; the step of shape 4 and its build entry stay.
    expect(removed).toHaveLength(1);
    expect(removed[0].local).toBe('par');
    expect(spids(root).filter((s) => s === '4')).toHaveLength(2);
    // One pass for each link took about 15 s here; one pass over the graph takes milliseconds.
    expect(ms).toBeLessThan(2_000);
  });
});

describe('deleting a shape removes the connector ends glued to it', () => {
  const connector = (st: number, end: number): string =>
    '<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="7" name="Connector 6"/><p:cNvCxnSpPr>' +
    `<a:stCxn id="${st}" idx="3"/><a:endCxn id="${end}" idx="1"/></p:cNvCxnSpPr><p:nvPr/></p:nvCxnSpPr>` +
    `<p:spPr><a:xfrm><a:off x="${2 * IN}" y="${IN + IN / 2}"/><a:ext cx="${IN}" cy="0"/></a:xfrm>` +
    '<a:prstGeom prst="straightConnector1"><a:avLst/></a:prstGeom></p:spPr></p:cxnSp>';

  test('the end glued to the deleted shape goes; the connector and its other end stay', async () => {
    const e = await open({ shapes: box(2, 1) + box(3, 3) + connector(2, 3) });
    const before = canonicalPackageFingerprint(pkgOf(e));
    deleteIds(e, [3]);
    expect(xmlOf(e)).toContain('<a:stCxn id="2" idx="3"/>');
    expect(xmlOf(e)).not.toContain('a:endCxn');
    expect(xmlOf(e)).toContain('<p:cxnSp>');
    saveAndReopen(e);
    expectUndoRestores(e, before);
  });

  test('an end that the file already left dangling stays', async () => {
    const e = await open({ shapes: box(2, 1) + box(3, 3) + box(4, 5) + connector(2, 99) });
    deleteIds(e, [4]);
    expect(xmlOf(e)).toContain('<a:endCxn id="99" idx="1"/>');
    expect(xmlOf(e)).toContain('<a:stCxn id="2" idx="3"/>');
  });
});

describe('saving keeps the p:timing that the session did not change', () => {
  // `CT_TimeNodeList` and `CT_BuildList` need a child, but a file can still hold empty ones.
  const EMPTY_STEP = (): string => {
    const w = timingWriter();
    return w.timing(w.mainSeq(w.click('') + w.click(w.effect(2))), [2]);
  };
  const savedSlide = (e: PptxEditorInstance): Element => {
    const r = readPptx(e.save()!);
    if (!r.ok) throw new Error('reopen failed');
    return rootOf(r.pkg, slidePart(e));
  };
  const emptyLists = (root: Element): number => {
    let count = 0;
    const walk = (el: Element): void => {
      if (['childTnLst', 'bldLst'].includes(el.local) && elements(el).length === 0) count += 1;
      for (const c of elements(el)) walk(c);
    };
    for (const t of timings(root)) walk(t);
    return count;
  };

  test('a deck with animations saves unchanged with no edits', async () => {
    for (const ext of [clickEntrances([2, 3]), EMPTY_STEP()]) {
      const e = await open({ shapes: box(2, 1) + box(3, 3), ext });
      const opened = canonicalPackageFingerprint(pkgOf(e));
      expect(semanticDigest(pkgOf(e))).toBe(opened);
      expect(serializeXmlPart({ ...xmlPartOf(pkgOf(e), slidePart(e))!, root: savedSlide(e) })).toBe(
        xmlOf(e)
      );
    }
  });

  test('an edit outside p:timing keeps the empty lists that the file wrote', async () => {
    const e = await open({ shapes: box(2, 1) + box(3, 3), ext: EMPTY_STEP() });
    expect(emptyLists(rootOf(pkgOf(e), slidePart(e)))).toBe(1);
    deleteIds(e, [3]);
    expect(emptyLists(savedSlide(e))).toBe(1);
  });

  test('an edit inside p:timing saves it without empty lists', async () => {
    const e = await open({
      shapes: box(2, 1) + box(3, 3),
      ext: (() => {
        const w = timingWriter();
        return w.timing(
          w.mainSeq(w.click('') + w.click(w.effect(2)) + w.click(w.effect(3))),
          [2, 3]
        );
      })(),
    });
    deleteIds(e, [3]);
    // The removal itself leaves the empty step that the file wrote; the save takes it.
    expect(emptyLists(rootOf(pkgOf(e), slidePart(e)))).toBe(1);
    const root = saveAndReopen(e);
    expect(emptyLists(root)).toBe(0);
    expect(spids(root)).toEqual(['2', '2']);
  });
});
