import { describe, expect, test } from 'bun:test';
import {
  applyColorTransforms,
  BLACK,
  drawingColorBase,
  isDrawingColorModel,
  parseHex,
  resolveDrawingColor,
  toCss,
  toHex,
  type DrawingColorTransform,
  type Rgba,
} from '../drawingml-color.ts';

const ACCENT1 = parseHex('4472C4')!;
const ACCENT2 = parseHex('ED7D31')!;

const t = (name: string, val?: string): DrawingColorTransform => ({ name, val });
const apply = (base: Rgba, ...transforms: DrawingColorTransform[]): string =>
  toHex(applyColorTransforms(base, transforms));
const base = (model: string, attrs: Record<string, string>): Rgba | undefined =>
  drawingColorBase(
    model,
    (name) => attrs[name],
    (name) => (name === 'accent1' ? ACCENT1 : undefined)
  );

// Expected values are sampled from the reference PDF of PPTX Editor's `colors` fixture, slide 1.
describe('color transforms (ECMA-376 Part 1, 20.1.2.3)', () => {
  test('inv inverts linear components', () => {
    expect(apply(ACCENT1, t('inv'))).toBe('#f8ebb3');
    expect(apply(parseHex('FFFFFF')!, t('inv'))).toBe('#000000');
  });

  test('gray uses Rec. 709 weights on sRGB components', () => {
    expect(apply(ACCENT2, t('gray'))).toBe('#8f8f8f');
  });

  test('HSL and linear transforms keep their measured values', () => {
    expect(apply(ACCENT1, t('lumMod', '75000'))).toBe('#2f5597');
    expect(apply(ACCENT1, t('lumMod', '60000'), t('lumOff', '40000'))).toBe('#8faadc');
    expect(apply(ACCENT2, t('shade', '50000'))).toBe('#ae5a21');
    expect(apply(ACCENT2, t('tint', '40000'))).toBe('#f8d7cd');
  });

  test('the worked shade example of 20.1.2.3.31 gives 00BC00', () => {
    expect(apply(parseHex('00FF00')!, t('shade', '50000'))).toBe('#00bc00');
  });

  test('transforms apply in document order', () => {
    expect(apply(ACCENT1, t('lumMod', '50000'), t('inv'))).not.toBe(
      apply(ACCENT1, t('inv'), t('lumMod', '50000'))
    );
  });

  test('alpha transforms set, scale, and offset the alpha', () => {
    const out = applyColorTransforms(ACCENT1, [t('alpha', '50000'), t('alphaMod', '50000')]);
    expect(out.a).toBe(0.25);
    expect(toCss(out)).toBe('rgba(68,114,196,0.25)');
  });

  test('an unknown transform is skipped', () => {
    expect(apply(ACCENT1, t('unknown', '50000'))).toBe('#4472c4');
  });
});

// The `a:ST_Percentage` family is a union of an `xsd:int` (`50000`) and the `s:ST_Percentage`
// pattern (`50%`), in the Strict and the Transitional schemas alike.
describe('the two spellings of a percentage', () => {
  test('a percentage with a % sign draws the same as thousandths of a percent', () => {
    expect(apply(ACCENT1, t('lumMod', '75%'))).toBe(apply(ACCENT1, t('lumMod', '75000')));
    expect(apply(ACCENT1, t('lumMod', '60%'), t('lumOff', '40%'))).toBe('#8faadc');
    expect(apply(ACCENT2, t('shade', '50%'))).toBe('#ae5a21');
    expect(apply(ACCENT2, t('tint', '40%'))).toBe('#f8d7cd');
    expect(apply(ACCENT1, t('satMod', '-12.5%'))).toBe(apply(ACCENT1, t('satMod', '-12500')));
    expect(applyColorTransforms(ACCENT1, [t('alpha', '12.5%')]).a).toBe(0.125);
  });

  test('lumMod 50% does not draw black', () => {
    expect(apply(ACCENT1, t('lumMod', '50%'))).not.toBe('#000000');
  });

  test('scrgbClr and hslClr read both spellings', () => {
    expect(toHex(base('scrgbClr', { r: '50%', g: '25%', b: '0%' })!)).toBe('#bc8900');
    expect(toHex(base('hslClr', { hue: '7200000', sat: '100%', lum: '50%' })!)).toBe('#00ff00');
  });

  test('an angle has no percentage spelling', () => {
    expect(apply(ACCENT1, t('hueOff', '50%'))).toBe('#4472c4');
  });

  test('an xsd:int can carry a leading plus sign', () => {
    expect(apply(ACCENT1, t('hueOff', '+7200000'))).toBe(apply(ACCENT1, t('hueOff', '7200000')));
    expect(apply(ACCENT2, t('tint', '+40000'))).toBe('#f8d7cd');
  });
});

describe('hostile values', () => {
  const isByte = (v: number): boolean => Number.isInteger(v) && v >= 0 && v <= 255;

  test('a malformed or huge value never gives NaN, Infinity, or a channel out of range', () => {
    for (const val of [
      'NaN',
      'Infinity',
      '1e5',
      '0x10',
      '50 %',
      '%',
      '.5%',
      '+50%',
      `${'9'.repeat(11)}%`,
      `${'9'.repeat(10_000)}`,
      `1.${'0'.repeat(10_000)}%`,
      `${'9'.repeat(10)}.${'9'.repeat(10)}%`,
    ]) {
      for (const name of ['lumMod', 'lumOff', 'hueOff', 'hueMod', 'tint', 'alphaMod', 'red']) {
        const out = applyColorTransforms(ACCENT1, [t(name, val)]);
        expect([out.r, out.g, out.b].every(isByte)).toBe(true);
        expect(out.a >= 0 && out.a <= 1).toBe(true);
      }
    }
  });

  test('the shared black cannot be changed by a caller', () => {
    expect(Object.isFrozen(BLACK)).toBe(true);
  });
});

describe('resolveDrawingColor', () => {
  test('resolves the base color and then the transforms, from a neutral element', () => {
    const attrs: Record<string, string> = { val: 'accent1' };
    const out = resolveDrawingColor(
      { model: 'schemeClr', readAttribute: (n) => attrs[n], transforms: [t('lumMod', '75%')] },
      (name) => (name === 'accent1' ? ACCENT1 : undefined)
    );
    expect(toHex(out!)).toBe('#2f5597');
  });

  test('an unresolved base gives undefined', () => {
    const out = resolveDrawingColor(
      { model: 'schemeClr', readAttribute: () => 'accent6', transforms: [] },
      () => undefined
    );
    expect(out).toBeUndefined();
  });
});

describe('color models', () => {
  test('the six DrawingML color models, and nothing else', () => {
    for (const model of ['srgbClr', 'schemeClr', 'sysClr', 'prstClr', 'scrgbClr', 'hslClr'])
      expect(isDrawingColorModel(model)).toBe(true);
    expect(isDrawingColorModel('solidFill')).toBe(false);
    expect(isDrawingColorModel('__proto__')).toBe(false);
  });

  test('each model resolves its base color', () => {
    expect(base('srgbClr', { val: '4472C4' })).toEqual(ACCENT1);
    expect(base('schemeClr', { val: 'accent1' })).toEqual(ACCENT1);
    expect(base('schemeClr', { val: 'accent6' })).toBeUndefined();
    expect(toHex(base('sysClr', { val: 'windowText', lastClr: '123456' })!)).toBe('#123456');
    expect(toHex(base('sysClr', { val: 'window' })!)).toBe('#ffffff');
    expect(toHex(base('prstClr', { val: 'cornflowerBlue' })!)).toBe('#6495ed');
    expect(toHex(base('scrgbClr', { r: '50000', g: '25000', b: '0' })!)).toBe('#bc8900');
    expect(toHex(base('hslClr', { hue: '7200000', sat: '100000', lum: '50000' })!)).toBe('#00ff00');
  });

  test('a preset name from the file cannot reach the prototype', () => {
    expect(base('prstClr', { val: '__proto__' })).toBeUndefined();
    expect(base('prstClr', { val: 'constructor' })).toBeUndefined();
  });

  test('a malformed hex or model reads as undefined', () => {
    expect(base('srgbClr', { val: '44 72C4' })).toBeUndefined();
    expect(base('solidFill', { val: '4472C4' })).toBeUndefined();
  });
});
