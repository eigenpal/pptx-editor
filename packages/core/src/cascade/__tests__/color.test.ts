import { describe, expect, test } from 'bun:test';
import { parseXmlPart, type Element } from '../../store/index.ts';
import { resolveColor, toHex, type ColorContext } from '../color.ts';

const A = 'http://schemas.openxmlformats.org/drawingml/2006/main';

const ctx: ColorContext = {
  scheme: new Map([
    ['dk1', { r: 0, g: 0, b: 0, a: 1 }],
    ['accent1', { r: 0x44, g: 0x72, b: 0xc4, a: 1 }],
    ['accent2', { r: 0xed, g: 0x7d, b: 0x31, a: 1 }],
  ]),
  colorMap: new Map([['tx1', 'dk1']]),
};

function color(xml: string): string {
  const parsed = parseXmlPart('/c.xml', xml.replace(/^<a:(\w+)/, `<a:$1 xmlns:a="${A}"`));
  if (!parsed.ok) throw new Error(parsed.reason);
  const rgba = resolveColor(parsed.part.root as Element, ctx);
  if (!rgba) throw new Error('unresolved');
  return toHex(rgba);
}

// Expected values are sampled from the reference PDF of the `colors` fixture, slide 1.
describe('color transforms (ECMA-376 Part 1, 20.1.2.3)', () => {
  test('inv inverts linear components', () => {
    expect(color('<a:schemeClr val="accent1"><a:inv/></a:schemeClr>')).toBe('#f8ebb3');
  });

  test('inv of white and black stays exact', () => {
    expect(color('<a:srgbClr val="FFFFFF"><a:inv/></a:srgbClr>')).toBe('#000000');
    expect(color('<a:srgbClr val="000000"><a:inv/></a:srgbClr>')).toBe('#ffffff');
  });

  test('gray uses Rec. 709 weights on sRGB components', () => {
    expect(color('<a:schemeClr val="accent2"><a:gray/></a:schemeClr>')).toBe('#8f8f8f');
  });

  test('the other transforms of the fixture keep their measured values', () => {
    expect(color('<a:schemeClr val="accent1"><a:lumMod val="75000"/></a:schemeClr>')).toBe(
      '#2f5597'
    );
    expect(
      color(
        '<a:schemeClr val="accent1"><a:lumMod val="60000"/><a:lumOff val="40000"/></a:schemeClr>'
      )
    ).toBe('#8faadc');
    expect(color('<a:schemeClr val="accent2"><a:shade val="50000"/></a:schemeClr>')).toBe(
      '#ae5a21'
    );
    expect(color('<a:schemeClr val="accent2"><a:tint val="40000"/></a:schemeClr>')).toBe('#f8d7cd');
  });

  test('a percentage written as 75% draws the same as 75000', () => {
    expect(color('<a:schemeClr val="accent1"><a:lumMod val="75%"/></a:schemeClr>')).toBe('#2f5597');
  });

  test('tx1 resolves through the color map, and phClr without a style color does not', () => {
    expect(color('<a:schemeClr val="tx1"/>')).toBe('#000000');
    expect(() => color('<a:schemeClr val="phClr"/>')).toThrow('unresolved');
  });
});
