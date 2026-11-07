import { describe, expect, test } from 'bun:test';
import { parseXmlPart } from '../../store/index.ts';
import { resolveEffects, type EffectPlacement } from '../../cascade/index.ts';
import { apply, boxMatrix } from '../../layout/index.ts';
import { effectMatrix } from '../effect-transforms.ts';

const IN = 914_400;
const base: EffectPlacement = {
  blur: 0,
  dist: 0,
  dir: 0,
  sx: 1,
  sy: 1,
  kx: 0,
  ky: 0,
  align: 'b',
  rotWithShape: false,
};
const box = { x: IN, y: IN, cx: 2 * IN, cy: IN, rot: 0, flipH: false, flipV: false };

const near = (a: { x: number; y: number }, b: { x: number; y: number }): void => {
  expect(Math.abs(a.x - b.x)).toBeLessThan(1);
  expect(Math.abs(a.y - b.y)).toBeLessThan(1);
};

describe('effect transform', () => {
  test('an offset moves by dist at dir', () => {
    const m = effectMatrix({ ...base, dist: 100, dir: 90 }, boxMatrix(box), box.cx, box.cy);
    near(apply(m, IN, IN), { x: IN, y: IN + 100 });
  });

  test('scale and skew keep the alignment point fixed', () => {
    const e = { ...base, sy: 0.23, kx: 13.34, align: 'br' as const };
    const m = effectMatrix(e, boxMatrix(box), box.cx, box.cy);
    near(apply(m, 3 * IN, 2 * IN), { x: 3 * IN, y: 2 * IN });
    // The top-right corner drops to 23 % of the height and leans by tan(kx) of that height.
    const top = apply(m, 3 * IN, IN);
    near(top, { x: 3 * IN - Math.tan((13.34 * Math.PI) / 180) * 0.23 * IN, y: 2 * IN - 0.23 * IN });
  });

  test('a reflection mirrors about the bottom edge', () => {
    const m = effectMatrix({ ...base, sy: -1, align: 'bl' }, boxMatrix(box), box.cx, box.cy);
    near(apply(m, IN, IN), { x: IN, y: 3 * IN });
    near(apply(m, IN, 2 * IN), { x: IN, y: 2 * IN });
  });

  test('rotWithShape turns the offset with the shape', () => {
    const rotated = boxMatrix({ ...box, rot: 90 });
    const center = apply(rotated, box.cx / 2, box.cy / 2);
    const still = effectMatrix({ ...base, dist: 100, dir: 0 }, rotated, box.cx, box.cy);
    const turns = effectMatrix(
      { ...base, dist: 100, dir: 0, rotWithShape: true },
      rotated,
      box.cx,
      box.cy
    );
    near(apply(still, center.x, center.y), { x: center.x + 100, y: center.y });
    near(apply(turns, center.x, center.y), { x: center.x, y: center.y + 100 });
  });
});

describe('effect model', () => {
  const parse = (xml: string) => {
    const r = parseXmlPart(
      '/x.xml',
      `<p:spPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><a:effectLst>${xml}</a:effectLst></p:spPr>`
    );
    if (!r.ok) throw new Error('bad xml');
    return resolveEffects(r.part.root, {
      scheme: new Map(),
      colorMap: new Map(),
    } as never)!;
  };

  test('schema defaults: bottom alignment and rotWithShape', () => {
    const fx = parse('<a:outerShdw blurRad="10"><a:srgbClr val="000000"/></a:outerShdw>');
    expect(fx.outerShadow?.align).toBe('b');
    expect(fx.outerShadow?.rotWithShape).toBe(true);
  });

  test('every reflection attribute resolves', () => {
    const fx = parse(
      '<a:reflection blurRad="6350" stA="50000" stPos="10000" endA="300" endPos="55000" dist="5000" dir="5400000" fadeDir="5400000" sy="-100000" kx="60000" algn="bl" rotWithShape="0"/>'
    );
    expect(fx.reflection).toMatchObject({
      blur: 6350,
      startAlpha: 0.5,
      startPos: 0.1,
      endPos: 0.55,
      dist: 5000,
      dir: 90,
      fadeDir: 90,
      sy: -1,
      kx: 1,
      align: 'bl',
      rotWithShape: false,
    });
  });

  test('every effect length is clamped like the outer shadow', () => {
    // All four are ST_PositiveCoordinate: 0 to 27 273 042 316 900 EMU.
    const huge = '27273042316900';
    const fx = parse(
      `<a:outerShdw blurRad="${huge}" dist="-5"><a:srgbClr val="000000"/></a:outerShdw>` +
        `<a:innerShdw blurRad="${huge}" dist="-5" dir="-5400000"><a:srgbClr val="000000"/></a:innerShdw>` +
        `<a:glow rad="${huge}"><a:srgbClr val="000000"/></a:glow><a:softEdge rad="-12700"/>`
    );
    const max = fx.outerShadow!.blur;
    expect(max).toBeLessThan(1e9);
    expect(fx.outerShadow!.dist).toBe(0);
    expect(fx.innerShadow).toMatchObject({ blur: max, dist: 0 });
    expect(fx.innerShadow!.dir).toBeGreaterThanOrEqual(0);
    expect(fx.glow!.radius).toBe(max);
    expect(fx.softEdge!.radius).toBe(0);
  });

  test('unknown effects are reported', () => {
    expect(parse('<a:blur rad="1000"/><a:presetShdw prst="shdw1"/>').unsupported).toEqual([
      'presetShdw',
    ]);
  });

  test('blur and preset shadows resolve, with a second copy for the double presets', () => {
    const fx = parse(
      '<a:blur rad="38100" grow="0"/>' +
        '<a:prstShdw prst="shdw13" dist="50800" dir="2700000"><a:srgbClr val="000000"><a:alpha val="50000"/></a:srgbClr></a:prstShdw>'
    );
    expect(fx.blur).toEqual({ radius: 38100, grow: false });
    expect(fx.outerShadow).toMatchObject({ blur: 0, dist: 50800, dir: 45, sx: 1, sy: 1 });
    expect(fx.secondShadow).toMatchObject({ dist: 101600 });
    expect(fx.secondShadow!.color.a).toBeCloseTo(0.3, 5);
    const lean = parse(
      '<a:prstShdw prst="shdw7" dist="38100"><a:srgbClr val="000000"/></a:prstShdw>'
    );
    expect(lean.outerShadow).toMatchObject({ sy: -0.5, kx: -60, align: 'b' });
    expect(lean.secondShadow).toBeUndefined();
    // An unknown preset is a plain copy; a hostile distance is clamped.
    const odd = parse(
      '<a:prstShdw prst="nope" dist="99999999999999"><a:srgbClr val="000000"/></a:prstShdw>'
    );
    expect(odd.outerShadow!.dist).toBeLessThan(1e9);
  });
});
