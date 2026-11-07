// Slide size (`setSlideSize`): the new `p:sldSz`, the boxes it scales on slides, layouts, and
// masters, the table grid, the `keep` mode, refusals, undo, and save and reopen.

import { describe, expect, test } from 'bun:test';
import { readFileSync } from 'node:fs';
import { join } from 'node:path';
import {
  attr,
  canonicalForm,
  child,
  elements,
  find,
  presentationInfo,
  readPptx,
  serializeXmlPart,
  shapeTreeOf,
  xmlPartOf,
  type Element,
  type PptxPackage,
} from '../../store/index.ts';
import { createPptxEditor, type PptxEditorInstance } from '../index.ts';

const FIXTURES = join(import.meta.dir, '../../../../../fixtures');

async function open(name: string): Promise<PptxEditorInstance> {
  const e = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await e.load(readFileSync(join(FIXTURES, `${name}.pptx`)))).ok).toBe(true);
  return e;
}

const pkgOf = (e: PptxEditorInstance): PptxPackage => e.session!.store.package;
const pkg = pkgOf;

/** The XML text of a part, whether the package holds it parsed or as bytes. */
function partText(p: PptxPackage, name: string): string {
  const part = p.parts.get(name);
  if (!part) throw new Error(`no part ${name}`);
  return part.kind === 'xml' ? serializeXmlPart(part.xml) : new TextDecoder().decode(part.bytes);
}

function reopen(e: PptxEditorInstance): PptxPackage {
  const r = readPptx(e.save()!);
  if (!r.ok) throw new Error('reopen');
  return r.pkg;
}

interface Box {
  readonly x: number;
  readonly y: number;
  readonly cx: number;
  readonly cy: number;
}

/** The own box of every top-level shape of a part's shape tree, by shape name. */
function boxes(pkg: PptxPackage, part: string): Map<string, Box> {
  const out = new Map<string, Box>();
  for (const el of elements(shapeTreeOf(pkg, part))) {
    const xfrm =
      el.local === 'graphicFrame'
        ? child(el, 'p:xfrm')
        : child(child(el, el.local === 'grpSp' ? 'p:grpSpPr' : 'p:spPr'), 'a:xfrm');
    const off = child(xfrm, 'a:off');
    const ext = child(xfrm, 'a:ext');
    const name = attr(find(el, 'p:cNvPr'), 'name');
    if (!off || !ext || !name) continue;
    out.set(name, {
      x: Number(attr(off, 'x')),
      y: Number(attr(off, 'y')),
      cx: Number(attr(ext, 'cx')),
      cy: Number(attr(ext, 'cy')),
    });
  }
  return out;
}

function gridOf(pkg: PptxPackage, part: string): number[] {
  const tbl = find(shapeTreeOf(pkg, part), 'a:tbl');
  return elements(child(tbl, 'a:tblGrid')).map((c: Element) => Number(attr(c, 'w')));
}

const sizeOf = (pkg: PptxPackage): [number, number] => {
  const info = presentationInfo(pkg);
  return [info.width, info.height];
};

describe('setSlideSize', () => {
  test('scales every box by the smaller ratio and centers it on the new slide', async () => {
    const e = await open('tables');
    const before = pkgOf(e);
    const [w, h] = sizeOf(before);
    expect([w, h]).toEqual([12_192_000, 6_858_000]);
    const slide = presentationInfo(before).slides[0].partName;
    const layout = [...before.parts.keys()].find((n) => n.startsWith('/ppt/slideLayouts/'))!;
    const was = boxes(before, slide);
    const wasLayout = boxes(before, layout);
    const wasGrid = gridOf(before, slide);
    expect(was.size).toBeGreaterThan(0);

    const r = e.exec({ type: 'setSlideSize', width: 9_144_000, height: 6_858_000 });
    expect(r.ok).toBe(true);
    expect(e.snapshot().slideWidth).toBe(9_144_000);
    expect(e.snapshot().slideHeight).toBe(6_858_000);

    // 4:3 from 16:9: the width ratio is the smaller one, 0.75, and the content is centered
    // vertically: (6858000 - 6858000 * 0.75) / 2 = 857250.
    const s = 0.75;
    for (const pkg of [pkgOf(e), reopen(e)]) {
      expect(sizeOf(pkg)).toEqual([9_144_000, 6_858_000]);
      for (const [name, b] of was) {
        const now = boxes(pkg, slide).get(name)!;
        expect(now.x).toBe(Math.round(b.x * s));
        expect(now.y).toBe(Math.round(b.y * s + 857_250));
        expect(now.cx).toBe(Math.round(b.cx * s));
        expect(now.cy).toBe(Math.round(b.cy * s));
      }
      for (const [name, b] of wasLayout)
        expect(boxes(pkg, layout).get(name)!.cx).toBe(Math.round(b.cx * s));
      expect(gridOf(pkg, slide)).toEqual(wasGrid.map((v) => Math.round(v * s)));
      expect(attr(child(xmlPartOf(pkg, '/ppt/presentation.xml')!.root, 'p:sldSz'), 'type')).toBe(
        undefined
      );
    }

    // One undo step restores the size and every box.
    expect(e.exec({ type: 'undo' }).ok).toBe(true);
    expect(sizeOf(pkgOf(e))).toEqual([w, h]);
    expect(boxes(pkgOf(e), slide)).toEqual(was);
    expect(gridOf(pkgOf(e), slide)).toEqual(wasGrid);
  });

  test('text scales with its box: sizes, spacing, insets, and the text styles', async () => {
    const e = await open('text-runs');
    const sizes = (pkg: PptxPackage): number[] => {
      const out: number[] = [];
      for (const [, p] of pkg.parts) {
        if (p.kind !== 'xml' || !['sld', 'sldLayout', 'sldMaster'].includes(p.xml.root.local))
          continue;
        const walk = (el: Element): void => {
          if (['rPr', 'defRPr', 'endParaRPr'].includes(el.local) && attr(el, 'sz'))
            out.push(Number(attr(el, 'sz')));
          for (const c of elements(el)) walk(c);
        };
        walk(p.xml.root);
      }
      const defaults = child(xmlPartOf(pkg, '/ppt/presentation.xml')!.root, 'p:defaultTextStyle');
      const walk = (el: Element): void => {
        if (el.local === 'defRPr' && attr(el, 'sz')) out.push(Number(attr(el, 'sz')));
        for (const c of elements(el)) walk(c);
      };
      if (defaults) walk(defaults);
      return out;
    };
    const was = sizes(pkgOf(e));
    expect(was.length).toBeGreaterThan(3);
    // 16:9 to 4:3: the factor is 0.75.
    expect(e.exec({ type: 'setSlideSize', width: 9_144_000, height: 6_858_000 }).ok).toBe(true);
    expect(sizes(reopen(e))).toEqual(was.map((v) => Math.max(100, Math.round(v * 0.75))));
    expect(e.exec({ type: 'undo' }).ok).toBe(true);
    expect(sizes(pkgOf(e))).toEqual(was);
  });

  test('diagram drawings scale in their frame, and chart text scales', async () => {
    const drawing = (p: PptxPackage): string => partText(p, '/ppt/diagrams/drawing1.xml');
    const e = await open('diagrams');
    const was = drawing(pkg(e));
    expect(e.exec({ type: 'setSlideSize', width: 9_144_000, height: 6_858_000 }).ok).toBe(true);
    const now = drawing(reopen(e));
    // The first shape: off (390, 208259) and ext (1582031, 786760), scaled by 0.75 from the
    // frame's origin.
    expect(was).toContain('x="390" y="208259"');
    expect(now).toContain('x="293" y="156194"');
    expect(now).toContain('cx="1186523" cy="590070"');
    // A second resize scales the drawing that the first one wrote back.
    expect(e.exec({ type: 'setSlideSize', width: 6_858_000, height: 5_143_500 }).ok).toBe(true);
    expect(drawing(reopen(e))).toContain('cx="889892" cy="442553"');
    // Diagram text keeps its size and insets.
    const sizes0 = [...was.matchAll(/ sz="(\d+)"| lIns="(\d+)"/g)].map((m) => m[0]);
    expect([...now.matchAll(/ sz="(\d+)"| lIns="(\d+)"/g)].map((m) => m[0])).toEqual(sizes0);
    const c = await open('charts');
    const sizes = (p: PptxPackage): string[] =>
      [...partText(p, '/ppt/charts/chart1.xml').matchAll(/ sz="(\d+)"/g)].map((m) => m[1]!);
    const before = sizes(pkg(c));
    expect(before).toContain('1862');
    expect(c.exec({ type: 'setSlideSize', width: 9_144_000, height: 6_858_000 }).ok).toBe(true);
    const once = before.map((v) => Math.round(Number(v) * 0.75));
    expect(sizes(reopen(c))).toEqual(once.map(String));
    expect(c.exec({ type: 'setSlideSize', width: 6_858_000, height: 5_143_500 }).ok).toBe(true);
    expect(sizes(reopen(c))).toEqual(once.map((v) => String(Math.round(v * 0.75))));
  });

  test('keep leaves every box where it is', async () => {
    const e = await open('tables');
    const slide = presentationInfo(pkgOf(e)).slides[0].partName;
    const was = boxes(pkgOf(e), slide);
    expect(
      e.exec({ type: 'setSlideSize', width: 12_192_000, height: 9_144_000, content: 'keep' }).ok
    ).toBe(true);
    const pkg = reopen(e);
    expect(sizeOf(pkg)).toEqual([12_192_000, 9_144_000]);
    expect(boxes(pkg, slide)).toEqual(was);
  });

  test('a group scales its own box; its children keep their child coordinates', async () => {
    const e = await open('groups-transforms');
    const slide = presentationInfo(pkgOf(e)).slides[0].partName;
    /** The transform of each child of each group, in child coordinates. */
    const children = (p: PptxPackage): string[] =>
      elements(shapeTreeOf(p, slide))
        .filter((el) => el.local === 'grpSp')
        .flatMap((g) => elements(g).filter((c) => c.local !== 'grpSpPr'))
        .map((c) => find(c, 'a:xfrm'))
        .filter((x): x is Element => x !== undefined)
        .map((x) => canonicalForm(x));
    const was = children(pkgOf(e));
    expect(was.length).toBeGreaterThan(0);
    expect(e.exec({ type: 'setSlideSize', width: 6_096_000, height: 3_429_000 }).ok).toBe(true);
    expect(children(reopen(e))).toEqual(was);
  });

  test('refuses a size out of 1 to 56 inches, and an unchanged size', async () => {
    const e = await open('tables');
    const [w, h] = sizeOf(pkgOf(e));
    const small = e.exec({ type: 'setSlideSize', width: 914_399, height: h });
    expect(small.ok).toBe(false);
    if (!small.ok) expect(small.reason).toBe('errors.slideSizeRange');
    const large = e.exec({ type: 'setSlideSize', width: w, height: 51_206_401 });
    expect(large.ok).toBe(false);
    const same = e.exec({ type: 'setSlideSize', width: w, height: h });
    expect(same.ok).toBe(false);
    if (!same.ok) expect(same.reason).toBe('disabled.noSizeChange');
    expect(e.can({ type: 'setSlideSize', width: w, height: 9_144_000 }).ok).toBe(true);
    expect(sizeOf(pkgOf(e))).toEqual([w, h]);
  });
});
