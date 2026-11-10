import { describe, expect, test } from 'bun:test';
import { strFromU8, strToU8, unzipSync, zipSync } from 'fflate';
import { buildPptx, shape } from '../../../../../scripts/fixtures/builder.ts';
import { serializeXmlPart, xmlPartOf } from '../../store/index.ts';
import { createPptxEditor, type PptxEditorInstance, type ShapeFormatPatch } from '../index.ts';
import { OpBatch } from '../../automation/commands/op-batch.ts';
import {
  patchShapeFormatProperties,
  readShapeFormatProperties,
} from '../../automation/commands/format-detail.ts';

const GRADIENT =
  '<a:gradFill rotWithShape="0" flip="xy"><a:gsLst>' +
  '<a:gs pos="0"><a:schemeClr val="accent2"><a:lumMod val="70000"/><a:alpha val="65000"/></a:schemeClr></a:gs>' +
  '<a:gs pos="100000"><a:srgbClr val="FFFFFF"/></a:gs></a:gsLst>' +
  '<a:lin ang="2700000" scaled="1"/><a:tileRect l="-20000"/></a:gradFill>';
const EFFECTS =
  '<a:effectLst><a:outerShdw blurRad="25400" dist="101600" dir="2700000" sx="87000" kx="120000" algn="br" rotWithShape="0" xmlns:x="urn:test" x:kept="yes">' +
  '<a:schemeClr val="accent1"><a:shade val="80000"/><a:alpha val="55000"/></a:schemeClr>' +
  '<x:custom value="kept"/></a:outerShdw>' +
  '<a:reflection blurRad="12700" dist="25400" stA="60000" endA="1000" stPos="2000" endPos="70000" sx="89000" sy="-100000" dir="5400000" fadeDir="5400000" rotWithShape="0"/>' +
  '</a:effectLst>';
function deck(): Uint8Array {
  return buildPptx({
    slides: [
      {
        layout: 'blank',
        shapes: shape({
          id: 2,
          x: 0,
          y: 0,
          cx: 914400,
          cy: 914400,
          fill: GRADIENT,
          effects: EFFECTS,
        }),
      },
      { shapes: shape({ id: 2, x: 0, y: 0, cx: 914400, cy: 914400 }) },
    ],
  });
}
async function open(bytes = deck()): Promise<PptxEditorInstance> {
  const e = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await e.load(bytes)).ok).toBe(true);
  e.setSelection({ kind: 'shapes', nodeIds: [e.session!.layout(0)!.items[0]!.element.nodeId] });
  return e;
}
function detail(e: PptxEditorInstance) {
  const read = e.shapeFormat();
  if (!read.ok) throw new Error(read.reason);
  return read.format;
}
function patch(e: PptxEditorInstance, value: ShapeFormatPatch) {
  return e.exec({ type: 'patchShapeFormat', target: detail(e).target, patch: value });
}
const xml = (e: PptxEditorInstance) =>
  serializeXmlPart(xmlPartOf(e.session!.store.package, e.session!.info.slides[0]!.partName)!);

describe('partial shape formatting', () => {
  test('an unsupported effect graph does not hide an independent gradient', async () => {
    const files = unzipSync(deck());
    files['ppt/slides/slide1.xml'] = strToU8(
      strFromU8(files['ppt/slides/slide1.xml']!).replace(
        EFFECTS,
        '<a:effectDag><a:cont type="sib" name="kept"/></a:effectDag>'
      )
    );
    const e = await open(zipSync(files));
    expect(detail(e).effectsReason).toBe('errors.unsupported');
    expect(detail(e).gradient?.angle).toBe(45);
    expect(patch(e, { kind: 'gradient', angle: 100 }).ok).toBe(true);
    const after = xml(e);
    expect(after).toContain('<a:effectDag><a:cont type="sib" name="kept"/></a:effectDag>');
    expect(patch(e, { kind: 'shadow', distance: 4 }).ok).toBe(false);
    expect(xml(e)).toBe(after);
    e.destroy();
  });

  test('empty and invalid patches neither change XML nor create history', async () => {
    const e = await open();
    const before = xml(e);
    const stop = detail(e).gradient!.stops[0]!.id;
    for (const value of [
      { kind: 'gradient' },
      { kind: 'gradient', angle: undefined },
      { kind: 'gradient', stop: { id: stop } },
      { kind: 'shadow', distance: undefined },
      { kind: 'reflection', startOpacity: undefined },
      { kind: 'gradient', rotWithShape: 'false' },
      { kind: 'shadow', opacity: 1.1 },
    ] as unknown as ShapeFormatPatch[])
      expect(patch(e, value).ok).toBe(false);
    expect(xml(e)).toBe(before);
    expect(e.can({ type: 'undo' }).ok).toBe(false);
    e.destroy();
  });

  test('keeps foreign elements whose local names match colors and transforms', async () => {
    const files = unzipSync(deck());
    files['ppt/slides/slide1.xml'] = strToU8(
      strFromU8(files['ppt/slides/slide1.xml']!)
        .replace('<a:gs pos="0">', '<a:gs pos="0" xmlns:x="urn:test"><x:srgbClr val="foreign"/>')
        .replace(
          '<a:schemeClr val="accent2">',
          '<a:schemeClr val="accent2" xmlns:x="urn:test" x:tracking="kept">'
        )
        .replace(
          '<a:lumMod val="70000"/>',
          '<a:lumMod val="70000"/><x:alpha xmlns:x="urn:test" val="foreign"/>'
        )
    );
    const e = await open(zipSync(files));
    const id = detail(e).gradient!.stops[0]!.id;
    expect(patch(e, { kind: 'gradient', stop: { id, opacity: 0.25 } }).ok).toBe(true);
    expect(xml(e)).toContain('<x:srgbClr val="foreign"/>');
    expect(xml(e)).toContain('val="foreign"');
    expect(patch(e, { kind: 'gradient', stop: { id, color: '#112233' } }).ok).toBe(true);
    expect(xml(e)).toContain('<x:srgbClr val="foreign"/>');
    expect(xml(e)).toContain('<x:alpha');
    expect(xml(e)).toContain('x:tracking="kept"');
    expect(detail(e).gradient!.stops[0]!.opacity).toBe(0.25);
    e.destroy();
  });

  test('reads exact numeric values without converting custom colors', async () => {
    const e = await open();
    const f = detail(e);
    expect(f.gradient?.angle).toBe(45);
    expect(f.gradient?.stops[0]?.opacity).toBe(0.65);
    expect(f.gradient?.stops[0]?.color).toBeUndefined();
    expect(f.shadow).toMatchObject({ distance: 8, blur: 2, angle: 45, opacity: 0.55 });
    expect(f.reflection).toMatchObject({ distance: 2, startOpacity: 0.6, endPosition: 0.7 });
    e.destroy();
  });

  test('patches angle and stop opacity while preserving source gradient properties', async () => {
    const e = await open();
    const before = xml(e);
    expect(patch(e, { kind: 'gradient', angle: 90 }).ok).toBe(true);
    expect(xml(e)).toBe(before.replace('ang="2700000"', 'ang="5400000"'));
    const stop = detail(e).gradient!.stops[0]!;
    expect(patch(e, { kind: 'gradient', stop: { id: stop.id, opacity: 0.3 } }).ok).toBe(true);
    expect(xml(e)).toContain(
      '<a:schemeClr val="accent2"><a:lumMod val="70000"/><a:alpha val="30000"/></a:schemeClr>'
    );
    expect(e.exec({ type: 'undo' }).ok).toBe(true);
    expect(detail(e).gradient!.stops[0]!.opacity).toBe(0.65);
    expect(e.exec({ type: 'undo' }).ok).toBe(true);
    expect(xml(e)).toBe(before);
    e.destroy();
  });

  test('patches custom shadow and reflection without replacing other attributes or unknown XML', async () => {
    const e = await open();
    const before = xml(e);
    expect(patch(e, { kind: 'shadow', distance: 12 }).ok).toBe(true);
    expect(xml(e)).toBe(before.replace('dist="101600"', 'dist="152400"'));
    expect(patch(e, { kind: 'reflection', endOpacity: 0.15 }).ok).toBe(true);
    expect(xml(e)).toContain('endA="15000"');
    expect(xml(e)).toContain('sx="89000"');
    const saved = e.save()!;
    const reopened = await open(saved);
    expect(detail(reopened).shadow?.distance).toBe(12);
    expect(detail(reopened).reflection?.endOpacity).toBe(0.15);
    expect(xml(reopened)).toContain('x:kept="yes"');
    expect(xml(reopened)).toContain('<x:custom value="kept"/>');
    e.destroy();
    reopened.destroy();
  });

  test('adds and removes stops without rewriting other stops or gradient metadata', async () => {
    const e = await open();
    const before = xml(e);
    const originalIds = detail(e).gradient!.stops.map((s) => s.id);
    expect(
      patch(e, {
        kind: 'gradient',
        addStop: { position: 0.5, color: { theme: 'accent3' }, opacity: 0.4 },
      }).ok
    ).toBe(true);
    const stops = detail(e).gradient!.stops;
    expect(stops).toHaveLength(3);
    expect(stops[1]).toMatchObject({ position: 0.5, color: { theme: 'accent3' }, opacity: 0.4 });
    expect([stops[0]!.id, stops[2]!.id]).toEqual(originalIds);
    expect(patch(e, { kind: 'gradient', removeStop: stops[1]!.id }).ok).toBe(true);
    expect(xml(e)).toBe(before);
    expect(patch(e, { kind: 'gradient', removeStop: stops[0]!.id }).ok).toBe(false);
    expect(xml(e)).toBe(before);
    e.destroy();
  });

  test('refuses invalid values, missing stops, foreign documents, and another active part without writes', async () => {
    const e = await open();
    const before = xml(e);
    for (const bad of [NaN, Infinity, -1, 151])
      expect(patch(e, { kind: 'shadow', blur: bad }).ok).toBe(false);
    expect(patch(e, { kind: 'gradient', stop: { id: 'missing', position: 0.5 } }).ok).toBe(false);
    expect(xml(e)).toBe(before);
    const target = detail(e).target;
    e.goToSlide(1);
    expect(
      e.exec({ type: 'patchShapeFormat', target, patch: { kind: 'shadow', blur: 4 } }).ok
    ).toBe(false);
    await e.load(deck());
    expect(
      e.exec({ type: 'patchShapeFormat', target, patch: { kind: 'shadow', blur: 4 } }).ok
    ).toBe(false);
    e.destroy();
  });

  test('copies inherited formatting before a partial override and keeps the source part unchanged', async () => {
    const files = unzipSync(deck());
    const slide = strFromU8(files['ppt/slides/slide1.xml']!);
    const inherited = slide
      .replace(/<p:sld\b/, '<p:sldLayout')
      .replace('</p:sld>', '</p:sldLayout>');
    files['ppt/slideLayouts/slideLayout2.xml'] = strToU8(
      inherited.replace('<p:nvPr></p:nvPr>', '<p:nvPr><p:ph type="body" idx="4"/></p:nvPr>')
    );
    files['ppt/slides/slide1.xml'] = strToU8(
      slide
        .replace(GRADIENT, '')
        .replace(EFFECTS, '')
        .replace('<p:nvPr></p:nvPr>', '<p:nvPr><p:ph type="body" idx="4"/></p:nvPr>')
    );
    const e = await open(zipSync(files));
    const originalLayout = serializeXmlPart(
      xmlPartOf(e.session!.store.package, '/ppt/slideLayouts/slideLayout2.xml')!
    );
    // Operations in one automation batch can use a stop captured before materialization.
    const batch = new OpBatch(e.session!.store.package);
    const target = detail(e).target;
    const sourceStop = detail(e).gradient!.stops[0]!.id;
    patchShapeFormatProperties(batch, target.part, target.nodeId, { kind: 'gradient', angle: 90 });
    patchShapeFormatProperties(batch, target.part, target.nodeId, {
      kind: 'gradient',
      stop: { id: sourceStop, opacity: 0.2 },
    });
    const batched = readShapeFormatProperties(batch, target.part, target.nodeId);
    expect(batched.gradient?.angle).toBe(90);
    expect(batched.gradient?.stops[0]?.opacity).toBe(0.2);
    // A new batch must not reinterpret an old source id as another stop's position.
    expect(() =>
      patchShapeFormatProperties(new OpBatch(batch.pkg), target.part, target.nodeId, {
        kind: 'gradient',
        stop: { id: sourceStop, opacity: 0.3 },
      })
    ).toThrow();
    expect(detail(e).gradient?.angle).toBe(45);
    expect(patch(e, { kind: 'gradient', angle: 80 }).ok).toBe(true);
    expect(patch(e, { kind: 'shadow', blur: 5 }).ok).toBe(true);
    expect(xml(e)).toContain('flip="xy"');
    expect(xml(e)).toContain('x:kept="yes"');
    expect(
      serializeXmlPart(xmlPartOf(e.session!.store.package, '/ppt/slideLayouts/slideLayout2.xml')!)
    ).toBe(originalLayout);
    e.destroy();
  });

  test('keeps theme parts unchanged when overriding a theme gradient and shadow', async () => {
    const files = unzipSync(deck());
    files['ppt/slides/slide1.xml'] = strToU8(
      strFromU8(files['ppt/slides/slide1.xml']!)
        .replace(GRADIENT, '')
        .replace(EFFECTS, '')
        .replace(
          '</p:spPr>',
          '</p:spPr><p:style><a:lnRef idx="0"><a:schemeClr val="accent1"/></a:lnRef><a:fillRef idx="1"><a:schemeClr val="accent4"/></a:fillRef><a:effectRef idx="3"><a:schemeClr val="accent3"/></a:effectRef><a:fontRef idx="minor"><a:schemeClr val="tx1"/></a:fontRef></p:style>'
        )
    );
    for (const name of ['ppt/theme/theme1.xml', 'ppt/theme/theme2.xml']) {
      files[name] = strToU8(
        strFromU8(files[name]!).replace(
          /(<a:fillStyleLst>)<a:solidFill>.*?<\/a:solidFill>/,
          `$1${GRADIENT.replace('val="accent2"', 'val="phClr"')}`
        )
      );
    }
    const e = await open(zipSync(files));
    const before = ['theme1', 'theme2'].map((n) =>
      serializeXmlPart(xmlPartOf(e.session!.store.package, `/ppt/theme/${n}.xml`)!)
    );
    expect(detail(e).gradient?.angle).toBe(45);
    expect(detail(e).shadow).not.toBeNull();
    expect(patch(e, { kind: 'gradient', angle: 120 }).ok).toBe(true);
    expect(patch(e, { kind: 'shadow', distance: 15 }).ok).toBe(true);
    expect(xml(e)).toContain(
      '<a:schemeClr val="accent4"><a:lumMod val="70000"/><a:alpha val="65000"/></a:schemeClr>'
    );
    expect(
      ['theme1', 'theme2'].map((n) =>
        serializeXmlPart(xmlPartOf(e.session!.store.package, `/ppt/theme/${n}.xml`)!)
      )
    ).toEqual(before);
    e.destroy();
  });
});
