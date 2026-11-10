import { expect, test } from 'bun:test';
import { ARROWHEAD_STYLE, ShapeArrowheadStyle } from '../model/enums.ts';

test('endpoint styles map public PascalCase values to DrawingML and back', () => {
  const pairs = [
    [ShapeArrowheadStyle.none, 'none'],
    [ShapeArrowheadStyle.triangle, 'triangle'],
    [ShapeArrowheadStyle.open, 'arrow'],
    [ShapeArrowheadStyle.stealth, 'stealth'],
    [ShapeArrowheadStyle.oval, 'oval'],
    [ShapeArrowheadStyle.diamond, 'diamond'],
  ] as const;
  for (const [publicValue, fileValue] of pairs) {
    expect(ARROWHEAD_STYLE.toFile(publicValue, 'beginArrowheadStyle')).toBe(fileValue);
    expect(ARROWHEAD_STYLE.fromFile(fileValue)).toBe(publicValue);
  }
  for (const invalid of ['arrow', 'OpenArrow', 'triangle', '__proto__']) {
    expect(() => ARROWHEAD_STYLE.toFile(invalid, 'endArrowheadStyle')).toThrow();
  }
});
