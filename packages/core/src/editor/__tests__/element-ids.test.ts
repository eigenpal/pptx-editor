// Ids that the editor mints for new elements: shape ids (`p:cNvPr/@id`, ECMA-376 Part 1,
// 20.1.10.21 `ST_DrawingElementId`, `xsd:unsignedInt`), slide ids (`ST_SlideId`), comment
// indexes (`ST_Index`), and slide creation ids. Every id in a file is attacker-controlled: a
// hostile id never pushes a new id out of its range or onto an id in use.

import { describe, expect, test } from 'bun:test';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { buildPptx } from '../../../../../scripts/fixtures/builder.ts';
import { buildSlideStructureDeck } from '../../../../../scripts/fixtures/slide-structure.ts';
import {
  attr,
  elements,
  path,
  readPptx,
  shapeTreeOf,
  xmlPartOf,
  type Element,
  type PptxPackage,
} from '../../store/index.ts';
import { createPptxEditor, type PptxEditorInstance } from '../index.ts';
import { freeIds, MAX_MINTED_ID, unsignedIntId } from '../../automation/commands/package-edits.ts';
import { CommandRefusal } from '../../automation/commands/op-batch.ts';
import { treeOf } from '../../automation/commands/slide-structure.ts';

describe('unsignedIntId', () => {
  test('reads every xsd:unsignedInt lexical form in range', () => {
    const cases: [string, number][] = [
      ['0', 0],
      ['7', 7],
      ['0010', 10],
      ['4294967295', 4_294_967_295],
      [`${'0'.repeat(500)}42`, 42],
      // whiteSpace="collapse", an optional "+", and "-0" (XML Schema Part 2, 3.3.22).
      [' 5 ', 5],
      ['\t5\n', 5],
      ['+5', 5],
      ['-0', 0],
      ['-000', 0],
    ];
    for (const [raw, n] of cases) expect({ raw, id: unsignedIntId(raw) }).toEqual({ raw, id: n });
  });

  test('refuses every value that is not an xsd:unsignedInt in range', () => {
    for (const raw of [
      undefined,
      '',
      ' ',
      '+',
      '4294967296',
      '1e400',
      '1e3',
      '-1',
      '+-5',
      '5 5',
      '0x10',
      '1.5',
      'NaN',
      'Infinity',
      '9'.repeat(400),
      '١٢',
    ])
      expect({ raw, id: unsignedIntId(raw) }).toEqual({ raw, id: undefined });
    expect(unsignedIntId('255', 256)).toBeUndefined();
    expect(unsignedIntId('2147483648', 0, MAX_MINTED_ID)).toBeUndefined();
  });
});

describe('freeIds', () => {
  test('counts up from one past the largest id in range', () => {
    const next = freeIds(new Set([1, 2, 7, 4_294_967_295]), 1, MAX_MINTED_ID, 1);
    expect([next(), next(), next()]).toEqual([8, 9, 10]);
  });

  test('past the maximum, takes the lowest free ids, then refuses when the range is full', () => {
    const next = freeIds(new Set([1, 3, 5]), 1, 5);
    expect([next(), next()]).toEqual([2, 4]);
    expect(() => next()).toThrow(CommandRefusal);
  });

  test('a nearly full range costs one lookup per used id, not one per id of the range', () => {
    const ids = new Set<number>();
    for (let i = 1; i <= 100_000; i += 1) if (i !== 50_000) ids.add(i);
    // Counts the lookups instead of timing them, so the test is not flaky on a slow machine.
    let lookups = 0;
    const used = new (class extends Set<number> {
      override has(n: number): boolean {
        lookups += 1;
        return super.has(n);
      }
    })(ids);
    lookups = 0;
    const next = freeIds(used, 1, MAX_MINTED_ID, MAX_MINTED_ID);
    expect(next()).toBe(50_000);
    expect(next()).toBe(100_001);
    expect(lookups).toBeLessThanOrEqual(ids.size + 2);
  });

  test('an `after` beyond the maximum never gives an id past it', () => {
    const next = freeIds(new Set([1]), 1, 3, 4_294_967_295);
    expect([next(), next()]).toEqual([2, 3]);
    expect(() => next()).toThrow(CommandRefusal);
  });
});

const PNG = Uint8Array.from([137, 80, 78, 71, 13, 10, 26, 10]);

/** A rectangle with a raw `p:cNvPr/@id` value. */
const rect = (id: string, x = 0): string =>
  `<p:sp><p:nvSpPr><p:cNvPr id="${id}" name="R"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>` +
  `<p:spPr><a:xfrm><a:off x="${x}" y="0"/><a:ext cx="914400" cy="914400"/></a:xfrm>` +
  '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:sp>';

/** The hostile ids of the issue, on one slide. */
const HOSTILE = [
  '4294967295',
  '4294967294',
  '1e400',
  '-1',
  '0',
  '0x10',
  '0007',
  '9'.repeat(400),
  'NaN',
  '12',
  '12',
];

async function open(bytes: Uint8Array): Promise<PptxEditorInstance> {
  const e = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await e.load(bytes)).ok).toBe(true);
  return e;
}

const hostileDeck = (ids: readonly string[] = HOSTILE): Uint8Array =>
  buildPptx({
    footers: true,
    slides: [{ layout: 'blank', shapes: ids.map((id, i) => rect(id, i * 100_000)).join('') }],
  });

const pkgOf = (e: PptxEditorInstance): PptxPackage => e.session!.store.package;
const slidePart = (e: PptxEditorInstance): string => e.session!.info.slides[0].partName;

/** Every `p:cNvPr/@id` of the slide's shape tree, raw. */
function rawIds(pkg: PptxPackage, part: string): string[] {
  const out: string[] = [];
  const walk = (el: Element): void => {
    if (el.local === 'cNvPr') out.push(attr(el, 'id') ?? '');
    for (const c of elements(el)) walk(c);
  };
  walk(shapeTreeOf(pkg, part)!);
  return out;
}

/** The ids that `run` adds: digits, in range, and unique on the slide (also after reopen). */
function expectMinted(e: PptxEditorInstance, run: () => void, count = 1): number[] {
  const before = rawIds(pkgOf(e), slidePart(e));
  run();
  const after = rawIds(pkgOf(e), slidePart(e));
  const added = [...after];
  for (const id of before) added.splice(added.indexOf(id), 1);
  expect(added.length).toBeGreaterThanOrEqual(count);
  const numbers = after.map((id) => unsignedIntId(id));
  for (const id of added) {
    expect(id).toMatch(/^[1-9]\d*$/);
    const n = Number(id);
    expect(n).toBeLessThanOrEqual(MAX_MINTED_ID);
    expect(numbers.filter((m) => m === n)).toHaveLength(1);
  }
  const r = readPptx(e.save()!);
  expect(r.ok).toBe(true);
  return added.map(Number);
}

const shapeNodes = (e: PptxEditorInstance): string[] =>
  elements(shapeTreeOf(pkgOf(e), slidePart(e))!)
    .filter((el) => el.kind === 'sp')
    .map((el) => el.id);

describe('new shape ids on a slide with hostile ids', () => {
  test('insert text box, shape, line, and picture', async () => {
    const e = await open(hostileDeck());
    expect(expectMinted(e, () => e.exec({ type: 'insertTextBox' }))).toEqual([13]);
    expectMinted(e, () => e.exec({ type: 'insertShape', preset: 'ellipse' }));
    expectMinted(e, () => e.exec({ type: 'insertLine', kind: 'line' }));
    expectMinted(e, () =>
      e.exec({ type: 'insertImage', bytes: PNG, mime: 'image/png', width: 10, height: 10 })
    );
  });

  test('insert table', async () => {
    const e = await open(hostileDeck());
    expectMinted(e, () => expect(e.exec({ type: 'insertTable', rows: 2, cols: 2 }).ok).toBe(true));
  });

  test('duplicate, copy and paste, and group', async () => {
    const e = await open(hostileDeck());
    const [a, b] = shapeNodes(e);
    e.setSelection({ kind: 'shapes', nodeIds: [a!, b!] });
    expectMinted(e, () => expect(e.exec({ type: 'duplicate' }).ok).toBe(true), 2);
    e.setSelection({ kind: 'shapes', nodeIds: [a!, b!] });
    const payload = e.copy()!;
    expectMinted(e, () => expect(e.exec({ type: 'paste', payload }).ok).toBe(true), 2);
    e.setSelection({ kind: 'shapes', nodeIds: [a!, b!] });
    expectMinted(e, () => expect(e.exec({ type: 'group' }).ok).toBe(true));
  });

  test('slide numbers', async () => {
    const e = await open(hostileDeck());
    expectMinted(e, () => expect(e.exec({ type: 'setSlideNumbers', show: true }).ok).toBe(true));
  });

  test('a layout change adds its empty placeholders', async () => {
    const e = await open(hostileDeck());
    const current = e.session!.info.slides[0].layoutPart;
    const other = e.snapshot().layouts!.find((l) => l.id !== current)!;
    expectMinted(e, () =>
      expect(e.exec({ type: 'applyLayout', indices: [0], layoutId: other.id }).ok).toBe(true)
    );
  });

  test('an id with whitespace or a "+" sign is still in use', async () => {
    const e = await open(hostileDeck(['2', ' 3 ', '+4']));
    expect(expectMinted(e, () => e.exec({ type: 'insertTextBox' }))).toEqual([5]);
  });

  test('a slide whose ids reach the maximum takes the lowest free ids', async () => {
    const e = await open(hostileDeck(['1', '2', String(MAX_MINTED_ID), '4']));
    expect(expectMinted(e, () => e.exec({ type: 'insertTextBox' }))).toEqual([3]);
    expect(expectMinted(e, () => e.exec({ type: 'insertTextBox' }))).toEqual([5]);
  });
});

describe('connector ends of copied shapes', () => {
  test('an end that names "0005" stays glued to the copy of the shape with id 5', async () => {
    const connector =
      '<p:cxnSp><p:nvCxnSpPr><p:cNvPr id="9" name="C"/><p:cNvCxnSpPr>' +
      '<a:stCxn id="5" idx="0"/></p:cNvCxnSpPr><p:nvPr/></p:nvCxnSpPr>' +
      '<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="914400" cy="0"/></a:xfrm>' +
      '<a:prstGeom prst="line"><a:avLst/></a:prstGeom></p:spPr></p:cxnSp>';
    const e = await open(
      buildPptx({ slides: [{ layout: 'blank', shapes: rect('0005') + connector }] })
    );
    const tree = (): Element => shapeTreeOf(pkgOf(e), slidePart(e))!;
    const ids = elements(tree())
      .filter((el) => el.kind === 'sp' || el.kind === 'cxnSp')
      .map((el) => el.id);
    e.setSelection({ kind: 'shapes', nodeIds: ids });
    expect(e.exec({ type: 'duplicate' }).ok).toBe(true);
    // Each copy follows its source.
    const copyOf = (kind: string): Element => elements(tree()).filter((el) => el.kind === kind)[1]!;
    const shapeId = attr(path(copyOf('sp'), 'p:nvSpPr', 'p:cNvPr'), 'id');
    const end = path(copyOf('cxnSp'), 'p:nvCxnSpPr', 'p:cNvCxnSpPr', 'a:stCxn');
    expect(shapeId).not.toBe('0005');
    expect(attr(end, 'id')).toBe(shapeId);
  });
});

describe('slide ids', () => {
  test('a new slide takes a free ST_SlideId when the file holds hostile slide ids', async () => {
    const files = unzipSync(buildPptx({ slides: [{ shapes: '' }, { shapes: '' }] }));
    const xml = strFromU8(files['ppt/presentation.xml']!);
    const ids = [...xml.matchAll(/<p:sldId id="(\d+)"/g)].map((m) => m[1]);
    files['ppt/presentation.xml'] = strToU8(
      xml.replace(`id="${ids[0]}"`, 'id="2147483647"').replace(`id="${ids[1]}"`, 'id="1e400"')
    );
    const e = await open(zipSync(files));
    for (let i = 0; i < 3; i += 1) expect(e.exec({ type: 'addSlide' }).ok).toBe(true);
    const list = path(xmlPartOf(pkgOf(e), e.session!.info.presentationPart)!.root, 'p:sldIdLst');
    const all = elements(list).map((s) => attr(s, 'id')!);
    for (const id of all) {
      expect(id).toMatch(/^\d+$/);
      expect(Number(id)).toBeGreaterThanOrEqual(256);
      expect(Number(id)).toBeLessThanOrEqual(2_147_483_647);
    }
    expect(new Set(all).size).toBe(all.length);
  });
});

describe('comment indexes and creation ids of a duplicated slide', () => {
  const hostileComments = (idx: readonly [string, string], lastIdx: string, creation: string) => {
    const files = unzipSync(buildSlideStructureDeck());
    const swap = (name: string, edit: (xml: string) => string): void => {
      files[name] = strToU8(edit(strFromU8(files[name]!)));
    };
    swap('ppt/comments/comment1.xml', (xml) =>
      xml.replace('idx="1"', `idx="${idx[0]}"`).replace('idx="2"', `idx="${idx[1]}"`)
    );
    swap('ppt/commentAuthors.xml', (xml) => xml.replace('lastIdx="2"', `lastIdx="${lastIdx}"`));
    swap('ppt/slides/slide1.xml', (xml) => xml.replace('val="1000001"', `val="${creation}"`));
    return zipSync(files);
  };

  const copied = async (bytes: Uint8Array) => {
    const e = await open(bytes);
    expect(e.exec({ type: 'duplicateSlide', index: 0 }).ok).toBe(true);
    const pkg = pkgOf(e);
    const comments = [...pkg.parts.keys()].filter((n) => /\/comments\/comment\d+\.xml$/.test(n));
    const idx = comments.flatMap((n) =>
      elements(treeOf(pkg, n)!.root).map((cm) => attr(cm, 'idx')!)
    );
    const author = elements(treeOf(pkg, '/ppt/commentAuthors.xml')!.root)[0];
    const creation = (i: number): string | undefined => {
      const root = xmlPartOf(pkg, e.session!.info.slides[i].partName)!.root;
      return attr(elements(elements(path(root, 'p:extLst'))[0])[0], 'val');
    };
    expect(readPptx(e.save()!).ok).toBe(true);
    return { idx, lastIdx: attr(author, 'lastIdx')!, creation: [creation(0), creation(1)] };
  };

  test('an index at the unsignedInt maximum: the copies take the lowest free indexes', async () => {
    const out = await copied(hostileComments(['4294967295', '1'], '4294967295', '4294967295'));
    expect(out.idx).toEqual(['4294967295', '1', '2', '3']);
    expect(out.lastIdx).toBe('4294967295');
    expect(out.creation[1]).toBe('1');
  });

  test('indexes that are not numbers set no maximum', async () => {
    const out = await copied(hostileComments(['1e400', '-1'], 'NaN', '-5'));
    expect(out.idx).toEqual(['1e400', '-1', '1', '2']);
    expect(out.lastIdx).toBe('2');
    expect(out.creation[1]).toMatch(/^[1-9]\d*$/);
  });

  test('a reply that names its parent as "001" by author "00" follows the copied parent', async () => {
    const files = unzipSync(buildSlideStructureDeck());
    const name = 'ppt/comments/comment1.xml';
    files[name] = strToU8(
      strFromU8(files[name]!).replace(
        '<p15:parentCm authorId="0" idx="1"/>',
        '<p15:parentCm authorId="00" idx="001"/>'
      )
    );
    const e = await open(zipSync(files));
    expect(e.exec({ type: 'duplicateSlide', index: 0 }).ok).toBe(true);
    const pkg = pkgOf(e);
    const copy = [...pkg.parts.keys()].find(
      (n) => /\/comments\/comment\d+\.xml$/.test(n) && n !== `/${name}`
    )!;
    const parent: string[] = [];
    const cms: string[] = [];
    const walk = (el: Element): void => {
      if (el.local === 'cm') cms.push(attr(el, 'idx')!);
      if (el.local === 'parentCm') parent.push(attr(el, 'idx')!);
      for (const c of elements(el)) walk(c);
    };
    walk(treeOf(pkg, copy)!.root);
    expect(cms).toEqual(['3', '4']);
    expect(parent).toEqual(['3']);
  });
});
