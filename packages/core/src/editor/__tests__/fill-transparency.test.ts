import { describe, expect, test } from 'bun:test';
import { buildPptx, shape } from '../../../../../scripts/fixtures/builder.ts';
import { serializeXmlPart, xmlPartOf } from '../../store/index.ts';
import {
  createPptxEditor,
  formatFillTransparency,
  type PptxEditorInstance,
  type ShapeFormatPatch,
} from '../index.ts';

const box = { x: 0, y: 0, cx: 914400, cy: 914400 };
const THEMED =
  '<a:solidFill><a:schemeClr val="accent2"><a:lumMod val="75000"/></a:schemeClr></a:solidFill>';
const STYLE =
  '<p:style><a:lnRef idx="0"><a:schemeClr val="accent1"/></a:lnRef>' +
  '<a:fillRef idx="1"><a:schemeClr val="accent1"/></a:fillRef>' +
  '<a:effectRef idx="0"><a:schemeClr val="accent1"/></a:effectRef>' +
  '<a:fontRef idx="minor"><a:schemeClr val="tx1"/></a:fontRef></p:style>';
const GRADIENT =
  '<a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs>' +
  '<a:gs pos="100000"><a:srgbClr val="0000FF"/></a:gs></a:gsLst><a:lin ang="0"/></a:gradFill>';

function deck(): Uint8Array {
  return buildPptx({
    slides: [
      {
        layout: 'blank',
        shapes:
          shape({ id: 2, name: 'Themed', ...box, fill: THEMED }) +
          shape({ id: 3, name: 'Styled', ...box, style: STYLE }) +
          shape({ id: 4, name: 'Gradient', ...box, fill: GRADIENT }) +
          shape({ id: 5, name: 'Empty', ...box, fill: '<a:noFill/>' }) +
          shape({
            id: 6,
            name: 'Half',
            ...box,
            fill: '<a:solidFill><a:srgbClr val="112233"><a:alpha val="50000"/></a:srgbClr></a:solidFill>',
          }),
      },
    ],
  });
}
async function open(): Promise<PptxEditorInstance> {
  const e = createPptxEditor({ fixedMeasurer: true, substitutes: false });
  expect((await e.load(deck())).ok).toBe(true);
  return e;
}
const idOf = (e: PptxEditorInstance, name: string) =>
  e.session!.layout(0)!.items.find((i) => i.element.name === name)!.element.nodeId;
const select = (e: PptxEditorInstance, ...names: string[]) =>
  e.setSelection({ kind: 'shapes', nodeIds: names.map((n) => idOf(e, n)) });
function detail(e: PptxEditorInstance) {
  const read = e.shapeFormat();
  if (!read.ok) throw new Error(read.reason);
  return read.format;
}
const patch = (e: PptxEditorInstance, value: ShapeFormatPatch) =>
  e.exec({ type: 'patchShapeFormat', target: detail(e).target, patch: value });
const xml = (e: PptxEditorInstance) =>
  serializeXmlPart(xmlPartOf(e.session!.store.package, e.session!.info.slides[0]!.partName)!);

describe('solid fill transparency', () => {
  test('reads the solid fill opacity, and no solid detail for other fills', async () => {
    const e = await open();
    select(e, 'Themed');
    expect(detail(e).solid).toEqual({ opacity: 1 });
    select(e, 'Half');
    expect(detail(e).solid).toEqual({ color: '#112233', opacity: 0.5 });
    select(e, 'Styled');
    expect(detail(e).solid).toEqual({ color: { theme: 'accent1' }, opacity: 1 });
    for (const name of ['Gradient', 'Empty']) {
      select(e, name);
      expect(detail(e).solid).toBeNull();
    }
    e.destroy();
  });

  test('writes alpha into the source color, keeps its transforms, and undoes in one step', async () => {
    const e = await open();
    select(e, 'Themed');
    const before = xml(e);
    expect(patch(e, { kind: 'solid', opacity: 0.6 }).ok).toBe(true);
    expect(xml(e)).toContain(
      '<a:schemeClr val="accent2"><a:lumMod val="75000"/><a:alpha val="60000"/></a:schemeClr>'
    );
    expect(detail(e).solid?.opacity).toBe(0.6);
    expect(patch(e, { kind: 'solid', opacity: 0.25 }).ok).toBe(true);
    expect(xml(e)).toContain('<a:alpha val="25000"/>');
    expect(e.exec({ type: 'undo' }).ok).toBe(true);
    expect(detail(e).solid?.opacity).toBe(0.6);
    expect(e.exec({ type: 'undo' }).ok).toBe(true);
    expect(xml(e)).toBe(before);
    e.destroy();
  });

  test('copies a style fill into the shape before it sets the alpha', async () => {
    const e = await open();
    select(e, 'Styled');
    expect(patch(e, { kind: 'solid', opacity: 0.7 }).ok).toBe(true);
    expect(xml(e)).toContain(
      '<a:solidFill><a:schemeClr val="accent1"><a:alpha val="70000"/></a:schemeClr></a:solidFill>'
    );
    expect(detail(e).solid).toEqual({ color: { theme: 'accent1' }, opacity: 0.7 });
    e.destroy();
  });

  test('refuses gradient, no fill, and invalid values without a write', async () => {
    const e = await open();
    const before = xml(e);
    select(e, 'Gradient');
    expect(patch(e, { kind: 'solid', opacity: 0.5 }).ok).toBe(false);
    select(e, 'Empty');
    expect(patch(e, { kind: 'solid', opacity: 0.5 }).ok).toBe(false);
    select(e, 'Themed');
    for (const bad of [NaN, -0.1, 1.1, undefined])
      expect(patch(e, { kind: 'solid', opacity: bad } as ShapeFormatPatch).ok).toBe(false);
    expect(xml(e)).toBe(before);
    expect(e.can({ type: 'undo' }).ok).toBe(false);
    e.destroy();
  });

  test('a new fill color keeps the current transparency', async () => {
    const e = await open();
    select(e, 'Half');
    expect(e.exec({ type: 'setShapeFill', color: { theme: 'accent3' } }).ok).toBe(true);
    expect(xml(e)).toContain(
      '<a:solidFill><a:schemeClr val="accent3"><a:alpha val="50000"/></a:schemeClr></a:solidFill>'
    );
    expect(detail(e).solid).toEqual({ color: { theme: 'accent3' }, opacity: 0.5 });
    expect(e.exec({ type: 'setShapeFill', color: null }).ok).toBe(true);
    expect(detail(e).solid).toBeNull();
    e.destroy();
  });
});

describe('formatFillTransparency', () => {
  test('gives the transparency percent of one solid shape and its patch', async () => {
    const e = await open();
    select(e, 'Half');
    const field = formatFillTransparency(e)!;
    expect(field).toMatchObject({
      id: 'fill.transparency',
      labelKey: 'inspector.transparency',
      value: 50,
      unit: '%',
      min: 0,
      max: 100,
      target: detail(e).target,
    });
    expect(field.patch(35)).toEqual({ kind: 'solid', opacity: 0.65 });
    select(e, 'Themed');
    expect(formatFillTransparency(e)?.value).toBe(0);
    e.destroy();
  });

  test('is null without one solid shape', async () => {
    const e = await open();
    expect(formatFillTransparency(null)).toBeNull();
    expect(formatFillTransparency(e)).toBeNull();
    for (const name of ['Gradient', 'Empty']) {
      select(e, name);
      expect(formatFillTransparency(e)).toBeNull();
    }
    select(e, 'Themed', 'Half');
    expect(formatFillTransparency(e)).toBeNull();
    e.destroy();
  });
});
